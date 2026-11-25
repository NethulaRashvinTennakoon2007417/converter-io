import fs from "node:fs"; import { readPdf, inspectPdf } from "./helpers.mjs";
const b = fs.readFileSync(process.argv[2]), pages = await readPdf(b), info = await inspectPdf(b);
console.log(JSON.stringify({ bytes: b.length, pages: pages.length, chars: pages.reduce((a, p) => a + p.text.replace(/\s/g, "").length, 0), pageImages: pages.map(p => p.images), imageObjects: info.images, fontFiles: info.fontFiles, firstLines: pages.map(p => p.lines[0].text), p1end: pages[0].lines.filter(l => l.y < pages[0].height - 95).slice(-3).map(l => l.text) }));
