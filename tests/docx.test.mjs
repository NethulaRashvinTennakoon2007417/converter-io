import test from "node:test"; import assert from "node:assert/strict";
import { convert, readPdf, inspectPdf, fixture } from "./helpers.mjs";
import { formatNumber, loadNumbering, createCounter, parseXml, lineHeightPt, spacingGap, linesThatFit } from "../docxToPdf.js";

/* ---------- pure helpers ---------- */
test("formatNumber covers every Word numbering format", () => {
  assert.equal(formatNumber("decimalZero", 1), "01"); assert.equal(formatNumber("decimalZero", 12), "12");
  assert.equal(formatNumber("lowerLetter", 2), "b"); assert.equal(formatNumber("lowerLetter", 27), "aa"); assert.equal(formatNumber("upperLetter", 3), "C");
  assert.equal(formatNumber("lowerRoman", 4), "iv"); assert.equal(formatNumber("upperRoman", 1994), "MCMXCIV");
  assert.equal(formatNumber("bullet", 1), "\u2022"); assert.equal(formatNumber("decimal", 7), "7");
});
test("numbering: counters advance, deeper levels reset, lvlText substitutes %1..%9, startOverride honoured", () => {
  const W = 'xmlns:w="x"';
  const num = loadNumbering(parseXml(`<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimalZero"/><w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
      <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/></w:lvl></w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    <w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="6"/></w:lvlOverride></w:num></w:numbering>`));
  const c = createCounter(num), seq = [[1, 0], [1, 0], [1, 1], [1, 1], [1, 2], [1, 0], [1, 1]].map(([n, l]) => c.next(n, l).label);
  assert.deepEqual(seq, ["01.", "02.", "02.1", "02.2", "(a)", "03.", "03.1"]);
  assert.equal(c.next(2, 0).label, "06.");
});
test("line height follows Word: 12pt Times at 1.15 = 15.86pt, exact and atLeast honoured", () => {
  assert.ok(Math.abs(lineHeightPt("serif", 12, { line: "276", lineRule: "auto" }) - 15.86) < 0.05);
  assert.equal(lineHeightPt("serif", 12, { line: "300", lineRule: "exact" }), 15);
  assert.ok(lineHeightPt("serif", 12, { line: "100", lineRule: "atLeast" }) > 13.7);
});
test("contextualSpacing removes the gap only between paragraphs of the same style", () => {
  const prev = { sid: "L", ctx: true, after: 8 };
  assert.equal(spacingGap(prev, { sid: "L", ctx: true, before: 13.55 }), 0);
  assert.equal(spacingGap(prev, { sid: "N", ctx: false, before: 13.55 }), 8 + 13.55);
});
test("widow/orphan control and keep-lines", () => {
  const h = Array(10).fill(10);
  assert.equal(linesThatFit(h, 1000), 10);
  assert.equal(linesThatFit(h, 95), 8);                       // 9 fit, but 1 would be left alone -> 8
  assert.equal(linesThatFit(h, 15), 0);                       // 1 line at the bottom would be an orphan
  assert.equal(linesThatFit(h, 95, { keepLines: true }), 0);
});

/* ---------- regression gate for the Policies.docx fixture ---------- */
let out, pages, info;
test.before(async () => { out = await convert(fixture("Policies.docx")); pages = await readPdf(out.bytes); info = await inspectPdf(out.bytes); });

test("GATE: output is real text, never a picture of text", () => {
  const chars = pages.reduce((a, p) => a + p.text.replace(/\s/g, "").length, 0);
  assert.ok(chars > 5000, `only ${chars} extractable characters`);
  pages.forEach((p, i) => { assert.equal(p.images, 0, `page ${i + 1} paints an image`); assert.ok(p.text.length > 150, `page ${i + 1} has almost no text`); });
  assert.equal(info.images, 0, "PDF contains image objects");
  assert.ok(info.fontFiles >= 2, "fonts are not embedded");
});
test("GATE: 4 pages, under 400 KB", () => { assert.equal(out.pages, 4); assert.equal(pages.length, 4); assert.ok(out.bytes.length < 400 * 1024, `${out.bytes.length} bytes`); });
test("GATE: running header once per page, first thing on every page", () => {
  pages.forEach((p, i) => { assert.equal(p.text.split("D.M.D.S.DISSANAYAKE").length - 1, 1, `page ${i + 1}`); assert.equal(p.lines[0].text, "D.M.D.S.DISSANAYAKE", `page ${i + 1} first line`); });
});
test("GATE: numbering is decimalZero, then 1. / 1.1 / (a)", () => {
  assert.ok(pages[0].lines.some(l => /^01\.\s+General$/.test(l.text)), "01. General");
  assert.ok(pages[0].lines.some(l => /^02\.\s+Principal/.test(l.text)) && pages[0].lines.some(l => /^03\.\s+Registered Office/.test(l.text)));
  assert.ok(pages[1].lines.some(l => /^1\.1\s+Basis of Preparation/.test(l.text)), "1.1 on page 2");
  assert.ok(pages[1].lines.some(l => /^\(a\)\s+Depreciation/.test(l.text)), "(a) on page 2");
  assert.ok(pages[1].lines.some(l => /^2\.\s+Valuation/.test(l.text)));
  assert.ok(pages[2].lines.some(l => /^3\.\s+Liabilities/.test(l.text)) || pages[3].lines.some(l => /^3\.\s+Liabilities/.test(l.text)));
  assert.ok(pages[3].lines.some(l => /^6\.\s+Comparatives/.test(l.text)), "numbering restarts at 6 for numId 10");
});
test("GATE: the Word page break survives (page 1 ends at the address block, page 2 starts at General Policies)", () => {
  const body = pages[0].lines.filter(l => l.y < pages[0].height - 95).map(l => l.text);
  assert.equal(body.at(-1), "Kottawa."); assert.match(body.at(-3), /No 3\/12, 3rd Lane, Athurugirya Road,/);
  assert.match(pages[1].lines[3].text, /^1\.\s+General Policies/);
  assert.match(pages[2].lines[3].text, /^2\.2\s+Investments/);
  assert.match(pages[3].lines[3].text, /^or receivable net of sales discounts/);
});
test("GATE: body line pitch on page 2 is 15.8 +/- 1.5 pt", () => {
  const ys = pages[1].lines.filter(l => l.y < pages[1].height - 95).map(l => l.y), d = [];
  for (let i = 1; i < ys.length; i++) { const g = ys[i - 1] - ys[i]; if (g < 20) d.push(g); }
  d.sort((a, b) => a - b); const med = d[Math.floor(d.length / 2)];
  assert.ok(d.length > 20 && Math.abs(med - 15.8) <= 1.5, `median pitch ${med}`);
});

test("META: the gate itself catches a rasterised page (the old html2canvas behaviour)", async () => {
  const { PDFLib } = await import("./helpers.mjs");
  const d = await PDFLib.PDFDocument.create(), png = await d.embedPng(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64"));
  d.addPage([595, 842]).drawImage(png, { x: 0, y: 0, width: 595, height: 842 });
  const bytes = await d.save(), [pg] = await readPdf(bytes), info = await inspectPdf(bytes);
  assert.ok(pg.images > 0 && info.images > 0 && pg.text.length < 150 && info.fontFiles === 0);
});
