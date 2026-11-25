import fs from "node:fs";import path from "node:path";import {createRequire} from "node:module";import {fileURLToPath} from "node:url";
const require=createRequire(import.meta.url),here=path.dirname(fileURLToPath(import.meta.url)),root=path.join(here,"..");
const NM=process.env.NM||path.join(here,"node_modules");
const PDFLib=require(NM+"/pdf-lib"),JSZip=require(NM+"/jszip"),fontkit=require(NM+"/@pdf-lib/fontkit");
const {docxToPdf}=await import("../docxToPdf.js");
const docx=process.argv[2]||path.join(here,"fixtures/Policies.docx"),out=process.argv[3]||path.join(here,"out.pdf");
const t=Date.now();
const r=await docxToPdf(fs.readFileSync(docx),{PDFLib,JSZip,fontkit,loadFont:async n=>fs.readFileSync(path.join(root,"fonts",n+".ttf"))});
fs.writeFileSync(out,r.bytes);console.log("pages",r.pages,"bytes",r.bytes.length,"ms",Date.now()-t,"warnings",r.warnings);
