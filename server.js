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

/* ---------------- CORS ---------------- */
const allow = (process.env.ALLOWED_ORIGINS || "")
  .split(",").map(s => s.trim()).filter(Boolean);
app.use(cors({
  origin: (origin, cb) => {
    if (!origin || allow.length === 0 || allow.includes(origin)) return cb(null, true);
    return cb(new Error("CORS blocked"));
  }
}));

app.get("/health", (_, res) => res.json({ ok: true }));

/* ---------------- Upload ---------------- */
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 } // 15MB
});

/* ---------------- Utils ---------------- */
const toArr = (x) => Array.isArray(x) ? x : (x == null ? [] : [x]);
const escHtml = (s) => String(s).replace(/[&<>]/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]));
const escAttr = (s) => String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;');
const textOf = (v) => {
  if (typeof v === "string") return v;
  if (v && typeof v === "object") return (v["#text"] ?? v.text ?? v._ ?? "");
  return "";
};

/* ---------------- DOCX unzip helpers ---------------- */
async function openDocx(buffer) { return await unzipper.Open.buffer(buffer); }
async function readEntry(entry) {
  const chunks = [];
  return new Promise((resolve, reject) => {
    entry.stream().on("data", c => chunks.push(c))
      .on("end", () => resolve(Buffer.concat(chunks)))
      .on("error", reject);
  });
}

/* ---------------- XML helpers ---------------- */
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
      if (k === "v:imagedata" && v?.["r:id"]) ids.add(v["r:id"]); // VML preview
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

/* ---------------- Debug: list OLE streams ---------------- */
function listCfbStreams(binBuffer) {
  try {
    const cf = CFB.read(binBuffer, { type: "buffer" });
    return cf.FileIndex.filter(fi => fi.content).map(fi => ({
      name: fi.name, size: (fi.content?.length || 0)
    })).slice(0, 30);
  } catch { return []; }
}

/* ---------------- Convert OLE .bin → MathML → LaTeX ---------------- */
function convertOleBinToMathMLAndTeX(oleBinBuffer, tmpName) {
  const tmpPath = path.join(os.tmpdir(), tmpName);
  fs.writeFileSync(tmpPath, oleBinBuffer);

  let mathml = "", latex = "", error = "", error_detail = "";
  try {
    mathml = execFileSync("ruby", [path.join(process.cwd(), "mt2mml.rb"), tmpPath], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"]
    });
    if (!mathml || !mathml.trim().startsWith("<")) error = "converter_empty_mathml";
  } catch (e) {
    error = "ruby_converter_error";
    try { error_detail = (e && (e.stderr ? e.stderr.toString("utf8") : e.message || "")) || ""; } catch {}
  } finally {
    try { fs.unlinkSync(tmpPath); } catch {}
  }

  if (mathml && mathml.trim().startsWith("<")) {
    try { latex = MathMLToLaTeX.convert(mathml); }
    catch { error = error || "latex_convert_failed"; }
  }
  return { mathml, latex, error, error_detail };
}

/* ---------------- Build inline HTML: chèn công thức đúng vị trí (INLINE) ---------------- */
function buildInlineHtml(documentXml, equations){
  const parser = new XMLParser({ ignoreAttributes:false, attributeNamePrefix:"" });
  const root = parser.parse(documentXml);

  const eqByRid = {};
  (equations||[]).forEach(e => { if (e && e.rId) eqByRid[e.rId] = e; });

  // Ép mọi MathML về inline
  const forceInlineMathML = (mml) =>
    String(mml)
      .replace(/<math\b([^>]*?)display="block"/gi,'<math$1display="inline"')
      .replace(/<math\b(?![^>]*display=)/gi,'<math display="inline"');

  const body = root?.["w:document"]?.["w:body"];
  const paras = toArr(body?.["w:p"]);
  const htmlOut = [];

  const renderEqInline = (eq, rid) => {
    if (eq && eq.latex && eq.latex.trim()){
      const tex = eq.latex;
      // INLINE render với \( ... \)
      return `<span class="eq-inline" data-tex="${escAttr(tex)}">\\(${escHtml(tex)}\\)</span>`;
    }
    if (eq && eq.mathml && eq.mathml.trim()){
      return `<span class="eq-inline" data-has-mml="1">${forceInlineMathML(eq.mathml)}</span>`;
    }
    return `<span class="eq-inline-missing" title="${escAttr(rid || (eq && eq.error) || 'missing')}">[công&nbsp;thức]</span>`;
  };

  const walk = (node, out) => {
    if (node == null) return;

    if (typeof node === "string") { out.push(escHtml(node)); return; }

    // Text run
    if (node["w:t"] != null) {
      const wt = node["w:t"];
      if (Array.isArray(wt)) wt.forEach(x => out.push(escHtml(textOf(x))));
      else out.push(escHtml(textOf(wt)));
      return;
    }

    // Break / tab
    if (node["w:br"] != null) out.push("<br/>");
    if (node["w:tab"] != null) out.push("&emsp;");

    // OLE patterns
    if (node["w:object"] && node["w:object"]["o:OLEObject"]) {
      const rid = node["w:object"]["o:OLEObject"]["r:id"] || node["w:object"]["o:OLEObject"]["r:linkByRef"];
      out.push(renderEqInline(eqByRid[rid], rid)); return;
    }
    if (node["w:pict"] && node["w:pict"]["v:imagedata"] && node["w:pict"]["v:imagedata"]["r:id"]) {
      const rid = node["w:pict"]["v:imagedata"]["r:id"];
      out.push(renderEqInline(eqByRid[rid], rid)); return;
    }
    if (node["o:OLEObject"] && (node["o:OLEObject"]["r:id"] || node["o:OLEObject"]["r:linkByRef"])) {
      const rid = node["o:OLEObject"]["r:id"] || node["o:OLEObject"]["r:linkByRef"];
      out.push(renderEqInline(eqByRid[rid], rid)); return;
    }

    // Walk children
    if (node["w:p"]) toArr(node["w:p"]).forEach(ch => walk(ch, out));
    if (node["w:r"]) toArr(node["w:r"]).forEach(ch => walk(ch, out));

    for (const k of Object.keys(node)) {
      if (k === "w:p" || k === "w:r" || k === "w:t" || k === "w:br" || k === "w:tab" ||
          k === "w:object" || k === "w:pict" || k === "o:OLEObject") continue;
      const v = node[k];
      if (typeof v === "object") {
        if (Array.isArray(v)) v.forEach(ch => walk(ch, out));
        else walk(v, out);
      }
    }
  };

  for (const p of paras) {
    const buf = [];
    walk(p, buf);
    htmlOut.push(`<p>${buf.join("") || "&nbsp;"}</p>`);
  }

  return `
  <style>
    .eq-inline{padding:2px 4px;border-radius:6px}
    .eq-inline code{background:#0b1020;color:#d1e7ff;border-radius:6px;padding:4px 6px}
    .eq-inline-missing{color:#dc2626;border-bottom:1px dotted #dc2626}
  </style>
  ${htmlOut.join("\n")}
  `;
}

/* ---------------- Build plain text: chèn công thức $...$ ---------------- */
function buildPlainText(documentXml, equations){
  const parser = new XMLParser({ ignoreAttributes:false, attributeNamePrefix:"" });
  const root = parser.parse(documentXml);

  const eqByRid = {};
  (equations||[]).forEach(e => { if (e && e.rId) eqByRid[e.rId] = e; });

  const body = root?.["w:document"]?.["w:body"];
  const paras = toArr(body?.["w:p"]);
  const lines = [];

  const renderEqText = (eq) => {
    if (eq && eq.latex && eq.latex.trim()){
      // INLINE plain text: $ ... $
      return `$${eq.latex}$`;
    }
    if (eq && eq.mathml && eq.mathml.trim()){
      // Nếu không convert được LaTeX, để placeholder (có thể thay bằng chuyển đổi khác nếu muốn)
      return `[MATHML]`;
    }
    return `[MATH?]`;
  };

  const walk = (node, outArr) => {
    if (node == null) return;

    if (typeof node === "string") { outArr.push(node); return; }

    // Text run
    if (node["w:t"] != null) {
      const wt = node["w:t"];
      if (Array.isArray(wt)) wt.forEach(x => outArr.push(textOf(x)));
      else outArr.push(textOf(wt));
      return;
    }

    // Break / tab
    if (node["w:br"] != null) outArr.push("\n");
    if (node["w:tab"] != null) outArr.push("\t");

    // OLE patterns
    if (node["w:object"] && node["w:object"]["o:OLEObject"]) {
      const rid = node["w:object"]["o:OLEObject"]["r:id"] || node["w:object"]["o:OLEObject"]["r:linkByRef"];
      outArr.push(renderEqText(eqByRid[rid])); return;
    }
    if (node["w:pict"] && node["w:pict"]["v:imagedata"] && node["w:pict"]["v:imagedata"]["r:id"]) {
      const rid = node["w:pict"]["v:imagedata"]["r:id"];
      outArr.push(renderEqText(eqByRid[rid])); return;
    }
    if (node["o:OLEObject"] && (node["o:OLEObject"]["r:id"] || node["o:OLEObject"]["r:linkByRef"])) {
      const rid = node["o:OLEObject"]["r:id"] || node["o:OLEObject"]["r:linkByRef"];
      outArr.push(renderEqText(eqByRid[rid])); return;
    }

    // Walk children
    if (node["w:p"]) toArr(node["w:p"]).forEach(ch => walk(ch, outArr));
    if (node["w:r"]) toArr(node["w:r"]).forEach(ch => walk(ch, outArr));

    for (const k of Object.keys(node)) {
      if (k === "w:p" || k === "w:r" || k === "w:t" || k === "w:br" || k === "w:tab" ||
          k === "w:object" || k === "w:pict" || k === "o:OLEObject") continue;
      const v = node[k];
      if (typeof v === "object") {
        if (Array.isArray(v)) v.forEach(ch => walk(ch, outArr));
        else walk(v, outArr);
      }
    }
  };

  for (const p of paras) {
    const buf = [];
    walk(p, buf);
    // Gộp chuỗi, rút gọn khoảng trắng đa ký tự về 1 khoảng (nhẹ nhàng, tránh dính từ)
    const line = buf.join("").replace(/\u00A0/g, " ");
    lines.push(line);
  }
  return lines.join("\n\n");
}

/* ---------------- API ---------------- */
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

    // Thu thập các OLE .bin
    const bins = {};
    for (const file of directory.files) {
      if (file.path.startsWith("word/embeddings/") && file.path.endsWith(".bin")) {
        bins[file.path] = await readEntry(file);
      }
    }

    // Convert từng OLE
    const equations = [];
    for (const rId of oleRelIds) {
      const embPath = relMap[rId];
      if (embPath && bins[embPath]) {
        const name = path.basename(embPath);
        const progId = progIdMap[rId] || "";
        const { mathml, latex, error, error_detail } =
          convertOleBinToMathMLAndTeX(bins[embPath], name);
        equations.push({
          rId, embPath, name, progId,
          mathml, latex, error, error_detail,
          streams: listCfbStreams(bins[embPath])
        });
      }
    }

    // Fallback HTML + Inline HTML + Plain Text
    const htmlResult  = await mammoth.convertToHtml({ buffer: req.file.buffer });
    const htmlFallback = htmlResult.value || "";
    const inlineHtml   = buildInlineHtml(docXml, equations);
    const plainText    = buildPlainText(docXml, equations);

    res.json({ ok: true, count: equations.length, equations, htmlFallback, inlineHtml, plainText });
  } catch (e) {
    res.status(500).json({ error: e.message || String(e) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("Server listening on", PORT));
