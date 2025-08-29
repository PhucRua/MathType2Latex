import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import unzipper from "unzipper";
import { execFileSync } from "child_process";
import { XMLParser } from "fast-xml-parser";
import mammoth from "mammoth";

// Import CommonJS trong ESM
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const { MathMLToLaTeX } = require("mathml-to-latex");
const CFB = require("cfb");

const app = express();

/* ---------- CORS ---------- */
const allow = (process.env.ALLOWED_ORIGINS || "")
  .split(",").map(s => s.trim()).filter(Boolean);
app.use(cors({
  origin: (origin, cb) => {
    if (!origin || allow.length === 0 || allow.includes(origin)) return cb(null, true);
    return cb(new Error("CORS blocked"));
  }
}));

app.get("/health", (_, res) => res.json({ ok: true }));

/* ---------- Upload ---------- */
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 }
});

/* ---------- DOCX unzip helpers ---------- */
async function openDocx(buffer) { return await unzipper.Open.buffer(buffer); }
async function readEntry(entry) {
  const chunks = [];
  return new Promise((resolve, reject) => {
    entry.stream().on("data", c => chunks.push(c))
      .on("end", () => resolve(Buffer.concat(chunks)))
      .on("error", reject);
  });
}

/* ---------- XML helpers ---------- */
function mapRelIdToEmbedding(relsXml) {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "" });
  const rels = parser.parse(relsXml)?.Relationships?.Relationship || [];
  const map = {};
  (Array.isArray(rels) ? rels : [rels]).forEach(r => {
    if (r.Target && r.Id && String(r.Target).startsWith("embeddings/")) {
      map[r.Id] = "word/" + String(r.Target).replace(/^\.?\//, "");
    }
  });
  return map;
}
function findOleRelIds(documentXml) {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "" });
  const doc = parser.parse(documentXml);
  const ids = new Set();
  (function walk(obj) {
    if (!obj || typeof obj !== "object") return;
    for (const k of Object.keys(obj)) {
      const v = obj[k];
      if (k === "o:OLEObject" && v?.["r:id"]) ids.add(v["r:id"]);
      if (k === "v:imagedata" && v?.["r:id"]) ids.add(v["r:id"]);
      if (typeof v === "object") walk(v);
    }
  })(doc);
  return Array.from(ids);
}
function mapProgIdFromDocXml(documentXml) {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "" });
  const doc = parser.parse(documentXml);
  const map = {};
  (function walk(obj) {
    if (!obj || typeof obj !== "object") return;
    for (const k of Object.keys(obj)) {
      const v = obj[k];
      if (k === "o:OLEObject" && (v?.["r:id"] || v?.["r:linkByRef"])) {
        const rid = v["r:id"] || v["r:linkByRef"];
        if (rid) map[rid] = v.ProgID || v.progId || "";
      }
      if (typeof v === "object") walk(v);
    }
  })(doc);
  return map;
}

/* ---------- Debug: list streams ---------- */
function listCfbStreams(binBuffer) {
  try {
    const cf = CFB.read(binBuffer, { type: "buffer" });
    return cf.FileIndex.filter(fi => fi.content).map(fi => ({
      name: fi.name, size: (fi.content?.length || 0)
    })).slice(0, 30);
  } catch { return []; }
}

/* ---------- Heuristics: MTEF ---------- */
function isLikelyMTEF(buf) {
  if (!buf || buf.length < 4) return false;
  const major = buf[0], platform = buf[1];
  return major >= 2 && major <= 8 && platform >= 0 && platform <= 7;
}

/* Parse stream Ole10Native → trả payload gốc (thường là bytes MTEF hoặc file nhúng) */
function extractFromOle10Native(buf) {
  const b = Buffer.from(buf);
  function tryFrom(off) {
    try {
      let i = off;
      const readZ = () => {
        let j = i;
        while (j < b.length && b[j] !== 0) j++;
        const s = b.subarray(i, j).toString("binary");
        i = j + 1; // skip NUL
        return s;
      };
      /* format: [name]\0 [src]\0 [temp]\0 [DWORD len] [data...] */
      const _name = readZ();
      const _src  = readZ();
      const _tmp  = readZ();
      if (i + 4 > b.length) return null;
      const len = b.readUInt32LE(i); i += 4;
      if (len > 0 && i + len <= b.length) {
        return b.subarray(i, i + len);
      }
      return null;
    } catch { return null; }
  }
  return tryFrom(0) || tryFrom(4) || tryFrom(6);
}

/* ---------- Extract MTEF from OLE ---------- */
function extractMTEFFromOLE(binBuffer, progId = "") {
  let cf = null;
  try { cf = CFB.read(binBuffer, { type: "buffer" }); } catch { cf = null; }

  // 1) Ưu tiên các stream MathType quen thuộc
  const CANDIDATE_RX = [
    /Equation Native/i,
    /MathType Equation/i,
    /Mathtype Equation/i,
    /^Equation$/i,
    /EqnData/i,
    /Equation Data/i
  ];
  if (cf) {
    for (const rx of CANDIDATE_RX) {
      const hit = cf.FileIndex.find(fi => fi.content && rx.test(fi.name));
      if (hit) {
        const buf = Buffer.from(hit.content);
        if (isLikelyMTEF(buf)) return buf;
        const j = buf.indexOf(Buffer.from("MTEF"));
        if (j >= 0 && isLikelyMTEF(buf.subarray(j + 4))) return buf.subarray(j + 4);
        return buf; // để Ruby tự parse
      }
    }
    // 2) Thử stream Ole10Native (rất hay gặp với DSMT4)
    const ole10 = cf.FileIndex.find(fi => fi.content && /^\x01Ole10Native$/i.test(fi.name));
    if (ole10) {
      const payload = extractFromOle10Native(ole10.content);
      if (payload) {
        if (isLikelyMTEF(payload)) return payload;
        const k = payload.indexOf(Buffer.from("MTEF"));
        if (k >= 0 && isLikelyMTEF(payload.subarray(k + 4))) return payload.subarray(k + 4);
      }
    }
    // 3) Quét tất cả stream tìm byte-pattern giống MTEF
    for (const fi of cf.FileIndex) {
      if (!fi.content) continue;
      const buf = Buffer.from(fi.content);
      if (isLikelyMTEF(buf)) return buf;
      const k = buf.indexOf(Buffer.from("MTEF"));
      if (k >= 0 && isLikelyMTEF(buf.subarray(k + 4))) return buf.subarray(k + 4);
    }
  }

  // 4) Fallback: quét ngoài .bin
  try {
    const b = Buffer.from(binBuffer);
    if (isLikelyMTEF(b)) return b;
    const t = b.indexOf(Buffer.from("MTEF"));
    if (t >= 0 && isLikelyMTEF(b.subarray(t + 4))) return b.subarray(t + 4);
  } catch {}

  return null;
}

/* ---------- Convert ---------- */
function convertMtefBinToMathMLAndTeX(binBuffer, tmpName, progId) {
  const mtef = extractMTEFFromOLE(binBuffer, progId);
  if (!mtef) {
    return { mathml: "", latex: "", error: "no_mtef_found", error_detail: "" };
  }

  const tmpPath = path.join(os.tmpdir(), tmpName);
  fs.writeFileSync(tmpPath, mtef);

  let mathml = "";
  let error = "";
  let error_detail = "";
  try {
    mathml = execFileSync("ruby", [path.join(process.cwd(), "mt2mml.rb"), tmpPath], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"]
    });
    if (!mathml || !mathml.trim().startsWith("<")) {
      error = "converter_empty_mathml";
    }
  } catch (e) {
    error = "ruby_converter_error";
    try {
      error_detail = (e && (e.stderr ? e.stderr.toString("utf8") : e.message || "")) || "";
    } catch { error_detail = ""; }
    mathml = "";
  } finally {
    try { fs.unlinkSync(tmpPath); } catch {}
  }

  let latex = "";
  if (mathml && mathml.trim().startsWith("<")) {
    try { latex = MathMLToLaTeX.convert(mathml); }
    catch { error = error || "latex_convert_failed"; }
  }

  return { mathml, latex, error, error_detail };
}

/* ---------- API ---------- */
app.post("/convert", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const directory = await openDocx(req.file.buffer);
    const docEntry = directory.files.find(f => f.path === "word/document.xml");
    const relEntry = directory.files.find(f => f.path === "word/_rels/document.xml.rels");
    const docXml = docEntry ? (await readEntry(docEntry)).toString("utf8") : "";
    const relXml = relEntry ? (await readEntry(relEntry)).toString("utf8") : "";

    const relMap    = relXml ? mapRelIdToEmbedding(relXml) : {};
    const oleRelIds = findOleRelIds(docXml);
    const progIdMap = docXml ? mapProgIdFromDocXml(docXml) : {};

    const bins = {};
    for (const file of directory.files) {
      if (file.path.startsWith("word/embeddings/") && file.path.endsWith(".bin")) {
        bins[file.path] = await readEntry(file);
      }
    }

    const equations = [];
    for (const rId of oleRelIds) {
      const embPath = relMap[rId];
      if (embPath && bins[embPath]) {
        const name = path.basename(embPath);
        const progId = progIdMap[rId] || "";
        const { mathml, latex, error, error_detail } =
          convertMtefBinToMathMLAndTeX(bins[embPath], name, progId);
        equations.push({
          rId, embPath, name, progId,
          mathml, latex, error, error_detail,
          streams: listCfbStreams(bins[embPath])
        });
      }
    }

    const htmlResult = await mammoth.convertToHtml({ buffer: req.file.buffer });
    const html = htmlResult.value || "";

    res.json({ ok: true, count: equations.length, equations, htmlFallback: html });
  } catch (e) {
    res.status(500).json({ error: e.message || String(e) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("Server listening on", PORT));
