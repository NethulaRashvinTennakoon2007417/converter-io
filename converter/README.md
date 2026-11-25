# converter.io

Free, ad-free file converter that runs entirely in the browser. Static site, no build step.

Tools: merge, split, rotate, compress, images to PDF, PDF to images, OCR (English/Sinhala/Tamil), Word to PDF.

## Word to PDF
`docxToPdf.js` reads the .docx itself (styles, numbering, headers/footers, tables, pictures) and writes a real
text PDF with pdf-lib. Fonts in `fonts/` are Tinos, Arimo and Cousine (same metrics as Times New Roman, Arial and
Courier New, SIL Open Font License), trimmed to Latin coverage and embedded in the PDF.
Not supported yet: Sinhala/Tamil text, text boxes, footnotes, tables nested in tables.

## Deploy
1. Push this folder to GitHub.
2. In Vercel: Add New Project, import the repo, Application Preset "Other", leave build settings empty, Deploy.
`.vercelignore` keeps `tests/` off the public site.

## Tests (optional, needs Node 20+)
```
cd tests && npm install && npm test
```
The gate converts `tests/fixtures/Policies.docx` and fails if the PDF is a picture, lacks text/fonts, loses its
header or numbering, changes page count or breaks, or drifts from Word's line spacing.
**The fixture contains a real company name and address: keep the repo private or delete the fixture first.**
