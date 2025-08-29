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

// Import CJS trong ESM
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const { MathMLToLaTeX } = require("mathml-to-latex");
const CFB = require("cfb"); // <— parser OLE/Compound File

const app = express();

// CORS whitelist (để trống = allow all cho dev)
const allow = (process.env.ALLOWED_ORIGINS || "")
  .split(",")
  .map(s => s.trim())
  .filter(Boolean);

app.use(cors({
  origin: (origin, cb) => {
    if (!origin || allow.length === 0 || allow.includes(origin)) return cb(null, true);
    return cb(new Error("CORS blocked"));
  }
}));

app.get("/health", (_, res) => res.json({ ok: true }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 } // 15MB
});

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

/** Tìm và trả về phần MTEF “sạch” bên trong OLE .bin */
function extractMTEFFromOLE(binBuffer) {
  // 1) Thử parse chuẩn bằng CFB
  try {
    const cf = CFB.read(binBuffer, { type: "buffer" });
    // Các tên stream phổ biến chứa dữ liệu MathType
    const PATTERN = /Equation Native|MathType Equation|Mathtype Equation|Equation|MTEquation|MTEF/i;

    // Tìm stream khớp
    const entry = cf.FileIndex.find(fi => PATTERN.test(fi.name));
    if (entry && entry.content) {
      let buf = Buffer.from(entry.content); // Uint8Array -> Buffer
      // Trong stream này thường có “MTEF” ở đâu đó, cắt từ vị trí đó
      const sig = Buffer.from("MTEF");
      const idx = buf.indexOf(sig);
      if (idx >= 0) return buf.subarray(idx);
      // Nếu không thấy “MTEF”, có thể stream đã là MTEF thuần → trả nguyên
      return buf;
    }
  } catch (e) {
    // bỏ qua, fallback bên dưới
  }

  // 2) Fallback: tìm trực tiếp “MTEF” trong toàn bộ .bin
  try {
    const buf = Buffer.from(binBuffer);
    const sig = Buffer.from("MTEF");
    const idx = buf.indexOf(sig);
    if (idx >= 0) return buf.subarray(idx);
  } catch (e) {
    // ignore
  }

  return null; // không tìm thấy MTEF
}

function convertMtefBinToMathMLAndTeX(binBuffer, tmpName) {
  // Bóc MTEF
  const mtef = extractMTEFFromOLE(binBuffer);
  if (!mtef) {
    return { mathml: "", latex: "", error: "no_mtef_found" };
  }

  // Ghi ra file tạm để Ruby gem đọc
  const tmpPath = path.join(os.tmpdir(), tmpName);
  fs.writeFileSync(tmpPath, mtef);

  let mathml = "";
  let error = "";
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
    mathml = "";
  } finally {
    try { fs.unlinkSync(tmpPath); } catch {}
  }

  let latex = "";
  if (mathml && mathml.trim().startsWith("<")) {
    try {
      latex = MathMLToLaTeX.convert(mathml);
    } catch (e) {
      // nếu lỗi chuyển LaTeX, vẫn trả MathML
      error = error || "latex_convert_failed";
    }
  }

  return { mathml, latex, error };
}

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
  // Tìm r:id của đối tượng OLE MathType (trong document.xml)
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "" });
  const doc = parser.parse(documentXml);
  const ids = new Set();

  function walk(obj) {
    if (!obj || typeof obj !== "object") return;
    for (const k of Object.keys(obj)) {
      const v = obj[k];
      if (k === "o:OLEObject" && v?.["r:id"]) ids.add(v["r:id"]);
      if (k === "v:imagedata" && v?.["r:id"]) ids.add(v["r:id"]); // 1 số case dùng VML
      if (typeof v === "object") walk(v);
    }
  }
  walk(doc);
  return Array.from(ids);
}

app.post("/convert", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const directory = await openDocx(req.file.buffer);

    // Document + relationships
    const docEntry = directory.files.find(f => f.path === "word/document.xml");
    const relEntry = directory.files.find(f => f.path === "word/_rels/document.xml.rels");
    const docXml = docEntry ? (await readEntry(docEntry)).toString("utf8") : "";
    const relXml = relEntry ? (await readEntry(relEntry)).toString("utf8") : "";

    const relMap = relXml ? mapRelIdToEmbedding(relXml) : {};
    const oleRelIds = findOleRelIds(docXml);

    // Toàn bộ nhị phân OLE
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
        const { mathml, latex, error } = convertMtefBinToMathMLAndTeX(bins[embPath], name);
        equations.push({ rId, embPath, name, mathml, latex, error });
      }
    }

    // HTML fallback của cả tài liệu (không render công thức)
    const htmlResult = await mammoth.convertToHtml({ buffer: req.file.buffer });
    const html = htmlResult.value || "";

    res.json({
      ok: true,
      count: equations.length,
      equations,
      htmlFallback: html
    });
  } catch (e) {
    res.status(500).json({ error: e.message || String(e) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("Server listening on", PORT));
