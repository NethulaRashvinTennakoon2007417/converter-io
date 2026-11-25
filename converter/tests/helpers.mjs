import fs from "node:fs"; import path from "node:path"; import { createRequire } from "node:module"; import { fileURLToPath } from "node:url";
const require = createRequire(import.meta.url), here = path.dirname(fileURLToPath(import.meta.url)), root = path.join(here, "..");
export const NM = process.env.NM || path.join(here, "node_modules");
export const PDFLib = require(NM + "/pdf-lib"), JSZip = require(NM + "/jszip"), fontkit = require(NM + "/@pdf-lib/fontkit");
const pdfjs = require(NM + "/pdfjs-dist/legacy/build/pdf.js");
export const fixture = n => path.join(here, "fixtures", n);
export async function convert(file) {
  const { docxToPdf } = await import("../docxToPdf.js");
  return docxToPdf(fs.readFileSync(file), { PDFLib, JSZip, fontkit, loadFont: async n => fs.readFileSync(path.join(root, "fonts", n + ".ttf")) });
}
/* Per page: text lines (grouped by baseline y, sorted by x) with their y, via pdf.js getTextContent(). */
export async function readPdf(bytes) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), useSystemFonts: false, verbosity: 0 }).promise, pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const pg = await doc.getPage(i), tc = await pg.getTextContent(), ops = await pg.getOperatorList(), rows = new Map();
    for (const it of tc.items) { if (!it.str.trim()) continue; const y0 = Math.round(it.transform[5] * 10) / 10, key = [...rows.keys()].find(k => Math.abs(k - y0) < 5), y = key ?? y0; (rows.get(y) || rows.set(y, []).get(y)).push({ x: it.transform[4], w: it.width, s: it.str }); }
    const join = its => { its.sort((a, b) => a.x - b.x); let o = ""; its.forEach((t, k) => { if (k && t.x - (its[k - 1].x + its[k - 1].w) > 0.8 && !/\s$/.test(o)) o += " "; o += t.s; }); return o.replace(/\s+/g, " ").trim(); };
    const lines = [...rows].sort((a, b) => b[0] - a[0]).map(([y, its]) => ({ y, text: join(its) }));
    const images = ops.fnArray.filter(f => f === pdfjs.OPS.paintImageXObject || f === pdfjs.OPS.paintInlineImageXObject).length;
    pages.push({ lines, text: lines.map(l => l.text).join("\n"), images, height: pg.view[3] });
  }
  return pages;
}
/* Count image XObjects and embedded font files in the PDF structure itself. */
export async function inspectPdf(bytes) {
  const d = await PDFLib.PDFDocument.load(bytes, { updateMetadata: false }); let images = 0, fontFiles = 0;
  for (const [, o] of d.context.enumerateIndirectObjects()) {
    const dict = o.dict || o; if (!dict || !dict.get) continue;
    const g = n => dict.get(PDFLib.PDFName.of(n));
    if (g("Subtype") === PDFLib.PDFName.of("Image")) images++;
    if (g("FontFile2") || g("FontFile3") || g("FontFile")) fontFiles++;
  }
  return { images, fontFiles };
}
