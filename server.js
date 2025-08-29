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

// Import CommonJS packages trong ESM
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const { MathMLToLaTeX } = require("mathml-to-latex"); // CJS
const CFB = require("cfb");                            // CJS

const app = express();

// ---- CORS ----
const allow = (process.env.ALLOWED_ORIGINS || "")
  .split(",").map(s => s.trim()).filter(Boolean);
app.use(cors({
  origin: (origin, cb) => {
    if (!origin || allow.length === 0 || allow.includes(origin)) return cb(null, true);
    return cb(new Error("CORS blocked"));
  }
}));

// ---- Healthcheck ----
app.get("/health", (_, res) => res.json({ ok: true }));

// ---- Upload ----
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 } // 15MB
});

// ---- Helpers: DOCX unzip ----
async function openDocx(buffer) {
  return await unzipper.Open.buffer(buffer);
}
async function readEntry(entry) {
  const chunks = [];
  return new Promise((resolve, reject) => {
    entry.stream()
      .on("data", (c) => chunks.push(c))
      .on("end", () => resolve(Buffer.concat(chunks)))
      .on("error", reject);
  });
}

// ---- XML helpers ----
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
      if (k === "v:imagedata" && v?.["r:id"]) ids.add(v["r:id"]); // đôi khi dùng VML
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
  return map; // {'rId6':'Equation.DSMT4', ...}
}

// ---- Debug: liệt kê stream OLE ----
function listCfbStreams(binBuffer) {
  try {
    const cf = CFB.read(binBuffer, { type: "buffer" });
    return cf.FileIndex.filter(fi => fi.content).map(fi => ({
      name: fi.name, size: (fi.content?.length || 0)
    })).slice(0, 20);
  } catch { return []; }
}

// ---- Nhận diện MTEF “hợp lý” ----
function isLikelyMTEF(buf) {
  if (!buf || buf.length < 4) return false;
  const major = buf[0], platform = buf[1];
  return major >= 2 && major <= 8 && platform >= 0 && platform <= 7;
}

// ---- Bóc MTEF từ OLE (ưu tiên MathType DSMT4 / “Equation Native”) ----
function extractMTEFFromOLE(binBuffer, progId = "") {
  let cf = null;
  try { cf = CFB.read(binBuffer, { type: "buffer" }); } catch { cf = null; }

  const CANDIDATE_RX = [
    /Equation Native/i,       // phổ biến nhất
    /MathType Equation/i,
    /Mathtype Equation/i,
    /^Equation$/i,
    /EqnData/i,
    /Equation Data/i
  ];

  // 1) Nếu parse được CFB → thử các stream ưu tiên (không cắt “MTEF” để tránh hỏng data)
  if (cf) {
    for (const rx of CANDIDATE_RX) {
      const hit = cf.FileIndex.find(fi => fi.content && rx.test(fi.name));
      if (hit) {
        const buf = Buffer.from(hit.content);
        if (isLikelyMTEF(buf)) return buf;
        const sig = Buffer.from("MTEF");
        const j = buf.indexOf(sig);
        if (j >= 0 && isLikelyMTEF(buf.subarray(j + 4))) return buf.subarray(j + 4);
        return buf; // để Ruby thử parse “as-is”
      }
    }
    // 2) Quét tất cả stream để tìm vùng “giống MTEF”
    for (const fi of cf.FileIndex) {
      if (!fi.content) continue;
      const buf = Buffer.from(fi.content);
      if (isLikelyMTEF(buf)) return buf;
      const k = buf.indexOf(Buffer.from("MTEF"));
      if (k >= 0 && isLikelyMTEF(buf.subarray(k + 4))) return buf.subarray(k + 4);
    }
  }

  // 3) Fallback: tìm trong toàn bộ .bin (khi không phải CFB chuẩn)
  try {
    const b = Buffer.from(binBuffer);
    if (isLikelyMTEF(b)) return b;
    const t = b.indexOf(Buffer.from("MTEF"));
    if (t >= 0 && isLikelyMTEF(b.subarray(t + 4))) return b.subarray(t + 4);
  } catch {}

  return null;
}

// ---- Chuyển MTEF → MathML → LaTeX ----
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

// ---- API chính ----
app.post("/convert", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const directory = await openDocx(req.file.buffer);

    const docEntry = directory.files.find(f => f.path === "word/document.xml");
    const relEntry = directory.files.find(f => f.path === "word/_rels/document.xml.rels");
    const docXml = docEntry ? (await readEntry(docEntry)).toString("utf8") : "";
    const relXml = relEntry ? (await readEntry(relEntry)).toString("utf8") : "";

    const relMap   = relXml ? mapRelIdToEmbedding(relXml) : {};
    const oleRelIds = findOleRelIds(docXml);
    const progIdMap = docXml ? mapProgIdFromDocXml(docXml) : {};

    // Toàn bộ OLE nhị phân
    const bins = {};
    for (const file of directory.files) {
      if (file.path.startsWith("word/embeddings/") && file.path.endsWith(".bin")) {
        bins[file.path] = await readEntry(file);
      }
    }

    // Chuyển từng OLE → MathML/LaTeX
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

    // HTML fallback toàn văn (không render công thức)
    const htmlResult = await mammoth.convertToHtml({ buffer: req.file.buffer });
    const html = htmlResult.value || "";

    res.json({ ok: true, count: equations.length, equations, htmlFallback: html });
  } catch (e) {
    res.status(500).json({ error: e.message || String(e) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("Server listening on", PORT));
