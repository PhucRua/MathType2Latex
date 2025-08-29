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

// ✅ Import CommonJS trong file ESM
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const { MathMLToLaTeX } = require("mathml-to-latex"); // ← dùng require để lấy named export CJS

const app = express();

// CORS whitelist qua env ALLOWED_ORIGINS (phân tách dấu phẩy). Trống = allow all (dev).
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

function convertMtefBinToMathMLAndTeX(binBuffer, tmpName) {
  const tmpPath = path.join(os.tmpdir(), tmpName);
  fs.writeFileSync(tmpPath, binBuffer);

  let mathml = "";
  try {
    // Gọi script Ruby dùng gem mathtype_to_mathml (chuyển MTEF → MathML)
    mathml = execFileSync("ruby", [path.join(process.cwd(), "mt2mml.rb"), tmpPath], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"]
    });
  } catch (e) {
    mathml = "";
  } finally {
    try { fs.unlinkSync(tmpPath); } catch {}
  }

  let latex = "";
  if (mathml && mathml.trim().startsWith("<")) {
    try {
      latex = MathMLToLaTeX.convert(mathml); // ✅ dùng API đúng của gói
    } catch {}
  }
  return { mathml, latex };
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
  // Tìm r:id của OLE MathType trong document.xml
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "" });
  const doc = parser.parse(documentXml);
  const ids = new Set();

  function walk(obj) {
    if (!obj || typeof obj !== "object") return;
    for (const k of Object.keys(obj)) {
      const v = obj[k];
      if (k === "o:OLEObject" && v?.["r:id"]) ids.add(v["r:id"]);
      if (k === "v:imagedata" && v?.["r:id"]) ids.add(v["r:id"]); // một số file dùng VML
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
        const { mathml, latex } = convertMtefBinToMathMLAndTeX(bins[embPath], name);
        equations.push({ rId, embPath, name, mathml, latex });
      }
    }

    // HTML fallback của cả tài liệu (không render công thức)
    const htmlResult = await mammoth.convertToHtml({ buffer: req.file.buffer });
    const html = htmlResult.value || "";

    res.json({ ok: true, count: equations.length, equations, htmlFallback: html });
  } catch (e) {
    res.status(500).json({ error: e.message || String(e) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("Server listening on", PORT));
