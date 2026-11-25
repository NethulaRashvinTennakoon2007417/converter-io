/* docxToPdf.js - Word (.docx) to a real, text-based PDF. 100% client side, no DOM needed.
   Pipeline: parse OOXML -> resolve styles/numbering -> build paragraph/table blocks -> lay out lines
   -> paginate (headers, footers, widow/orphan, keep-with-next, page breaks) -> draw with pdf-lib.
   Pure helpers (formatNumber, createCounter, lineHeightPt, spacingGap, linesThatFit) are exported for tests. */

const TW = 1 / 20, SQUEEZE = 0.25;
/* ---------- 1. XML ---------- */
const ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const dec = s => s.indexOf("&") < 0 ? s : s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => e[0] === "#" ? String.fromCodePoint(/^#x/i.test(e) ? parseInt(e.slice(2), 16) : +e.slice(1)) : (ENT[e] ?? m));
export function parseXml(src) {
  const root = { n: "#", a: {}, c: [] }, st = [root];
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<(\/?)([\w:.-]+)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(src))) {
    const top = st[st.length - 1];
    if (m[6] !== undefined) { if (top.n === "w:t" || top.n === "w:instrText") top.c.push(dec(m[6])); continue; }
    if (m[1] !== undefined) { top.c.push(m[1]); continue; }
    if (m[3] === undefined) continue;
    if (m[2]) { if (st.length > 1) st.pop(); continue; }
    const a = {};
    for (const x of m[4].matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) a[x[1]] = dec(x[2] ?? x[3]);
    const el = { n: m[3], a, c: [] };
    top.c.push(el);
    if (!m[5]) st.push(el);
  }
  return root.c.find(x => typeof x !== "string");
}
const kid = (n, name) => n && n.c.find(x => x.n === name);
const kids = (n, name) => n ? n.c.filter(x => x.n === name) : [];
const av = (n, k) => n && n.a[k];
const on = n => !!n && !/^(0|false|off)$/i.test(n.a["w:val"] ?? "1");

/* ---------- 2. Numbering ---------- */
const roman = n => { let s = ""; for (const [v, r] of [[1000,"M"],[900,"CM"],[500,"D"],[400,"CD"],[100,"C"],[90,"XC"],[50,"L"],[40,"XL"],[10,"X"],[9,"IX"],[5,"V"],[4,"IV"],[1,"I"]]) while (n >= v) { s += r; n -= v; } return s; };
export function formatNumber(fmt, n) {
  switch (fmt) {
    case "decimalZero": return String(n).padStart(2, "0");
    case "lowerLetter": case "upperLetter": { const s = String.fromCharCode(97 + (n - 1) % 26).repeat(Math.floor((n - 1) / 26) + 1); return fmt === "upperLetter" ? s.toUpperCase() : s; }
    case "lowerRoman": return roman(n).toLowerCase();
    case "upperRoman": return roman(n);
    case "bullet": return "\u2022";
    case "none": return "";
    default: return String(n);
  }
}
export function loadNumbering(root) {
  const abs = {}, nums = {};
  for (const a of kids(root, "w:abstractNum")) {
    const lv = {};
    for (const l of kids(a, "w:lvl")) {
      const i = +av(l, "w:ilvl");
      lv[i] = { start: +(av(kid(l, "w:start"), "w:val") ?? 1), fmt: av(kid(l, "w:numFmt"), "w:val") || "decimal", text: av(kid(l, "w:lvlText"), "w:val") ?? `%${i + 1}.`,
        jc: av(kid(l, "w:lvlJc"), "w:val") || "left", suff: av(kid(l, "w:suff"), "w:val") || "tab", isLgl: !!kid(l, "w:isLgl"),
        restart: kid(l, "w:lvlRestart") ? +av(kid(l, "w:lvlRestart"), "w:val") : undefined, pPr: kid(l, "w:pPr"), rPr: kid(l, "w:rPr") };
    }
    abs[av(a, "w:abstractNumId")] = lv;
  }
  for (const n of kids(root, "w:num")) {
    const ov = {};
    for (const o of kids(n, "w:lvlOverride")) ov[+av(o, "w:ilvl")] = { start: kid(o, "w:startOverride") ? +av(kid(o, "w:startOverride"), "w:val") : undefined };
    nums[av(n, "w:numId")] = { abs: av(kid(n, "w:abstractNumId"), "w:val"), ov };
  }
  return { abs, nums };
}
/* One counter set per list. Lists that share an abstractNum continue each other (as in Word) unless a startOverride splits them. */
export function createCounter(num) {
  const state = {};
  return { next(numId, ilvl = 0) {
    const nm = num.nums[numId], levels = nm && num.abs[nm.abs], L = levels && levels[ilvl];
    if (!L) return null;
    const key = Object.values(nm.ov).some(o => o.start !== undefined) ? "n" + numId : "a" + nm.abs, c = state[key] ??= [];
    const startOf = i => nm.ov[i]?.start ?? levels[i]?.start ?? 1;
    c[ilvl] = c[ilvl] === undefined ? startOf(ilvl) : c[ilvl] + 1;
    for (let j = ilvl + 1; j < 9; j++) if (levels[j]?.restart !== 0) c[j] = undefined;
    const label = L.fmt === "bullet" ? "\u2022" : L.text.replace(/%([1-9])/g, (m, d) => { const j = d - 1; return formatNumber(L.isLgl ? "decimal" : (levels[j] || L).fmt, c[j] === undefined ? startOf(j) : c[j]); });
    return { label, lvl: L };
  } };
}

/* ---------- 3. Spacing / pagination helpers (pure) ---------- */
export const METRICS = { serif: { asc: .891, desc: .216, gap: .042 }, sans: { asc: .905, desc: .212, gap: .033 }, mono: { asc: .833, desc: .3, gap: 0 } };
export const natural = (cls, size) => size * (METRICS[cls].asc + METRICS[cls].desc + METRICS[cls].gap);
/* w:spacing line/lineRule -> absolute points. auto multiplies the font's natural line height (240 = single). */
export function lineHeightPt(cls, size, sp = {}) {
  const nat = natural(cls, size), line = sp.line === undefined ? undefined : +sp.line;
  if (line === undefined) return nat;
  if (sp.lineRule === "exact") return line * TW;
  if (sp.lineRule === "atLeast") return Math.max(nat, line * TW);
  return nat * line / 240;
}
/* Space between two paragraphs: after(prev) + before(cur), with contextualSpacing suppression for equal styles. */
export function spacingGap(prev, cur) {
  const sameStyle = prev && prev.sid === cur.sid;
  const after = prev && !(prev.ctx && sameStyle) ? prev.after : 0;
  const before = cur.ctx && sameStyle ? 0 : cur.before;
  return after + before;
}
/* How many of `heights` start on this page, honouring widow/orphan control and keep-lines. */
export function linesThatFit(heights, room, { widow = true, keepLines = false, atStart = true } = {}) {
  let k = 0, h = 0;
  while (k < heights.length && h + heights[k] <= room + 0.01) h += heights[k++];
  if (k === heights.length) return k;
  if (keepLines && atStart) return 0;
  if (widow) { if (atStart && k < 2) return 0; if (heights.length - k < 2) k = Math.max(0, heights.length - 2); if (atStart && k < 2) return 0; }
  return k;
}

/* ---------- 4. Properties and styles ---------- */
const numAttr = (n, k) => n && n.a[k] !== undefined ? +n.a[k] : undefined;
function pProps(pPr) {
  const p = {};
  if (!pPr) return p;
  const sp = kid(pPr, "w:spacing");
  if (sp) { p.sp = {}; for (const [k, v] of Object.entries(sp.a)) if (!/Lines$|Autospacing$/.test(k)) p.sp[k.slice(2)] = v; }
  const ind = kid(pPr, "w:ind");
  if (ind) { p.ind = {}; const g = (...ks) => { for (const k of ks) if (ind.a[k] !== undefined) return +ind.a[k]; }; p.ind = { left: g("w:left", "w:start"), right: g("w:right", "w:end"), hanging: g("w:hanging"), firstLine: g("w:firstLine") }; for (const k in p.ind) if (p.ind[k] === undefined) delete p.ind[k]; }
  const jc = av(kid(pPr, "w:jc"), "w:val"); if (jc) p.jc = jc;
  for (const [k, e] of [["keepNext", "w:keepNext"], ["keepLines", "w:keepLines"], ["widow", "w:widowControl"], ["pbb", "w:pageBreakBefore"], ["ctx", "w:contextualSpacing"]]) if (kid(pPr, e)) p[k] = on(kid(pPr, e));
  const tabs = kid(pPr, "w:tabs"); if (tabs) p.tabs = kids(tabs, "w:tab").map(t => ({ pos: +t.a["w:pos"], val: t.a["w:val"] || "left" }));
  const np = kid(pPr, "w:numPr"); if (np) p.numPr = { numId: av(kid(np, "w:numId"), "w:val"), ilvl: +(av(kid(np, "w:ilvl"), "w:val") ?? 0) };
  const ps = av(kid(pPr, "w:pStyle"), "w:val"); if (ps) p.pStyle = ps;
  return p;
}
function mergeP(a, b) {
  const o = { ...a, ...b };
  if (a.sp || b.sp) o.sp = { ...a.sp, ...b.sp };
  if (a.ind || b.ind) { const i = { ...a.ind, ...b.ind }; if (b.ind?.hanging !== undefined) delete i.firstLine; if (b.ind?.firstLine !== undefined) delete i.hanging; o.ind = i; }
  if (a.tabs || b.tabs) { const m = new Map(); for (const t of [...(a.tabs || []), ...(b.tabs || [])]) t.val === "clear" ? m.delete(t.pos) : m.set(t.pos, t); o.tabs = [...m.values()]; }
  return o;
}
function rProps(rPr) {
  const r = {};
  if (!rPr) return r;
  const f = kid(rPr, "w:rFonts");
  if (f) { const fo = f.a["w:ascii"] || f.a["w:hAnsi"], th = f.a["w:asciiTheme"] || f.a["w:hAnsiTheme"]; if (fo) r.font = fo; if (th) r.theme = th; }
  for (const [k, e] of [["b", "w:b"], ["i", "w:i"], ["strike", "w:strike"], ["caps", "w:caps"], ["hide", "w:vanish"]]) if (kid(rPr, e)) r[k] = on(kid(rPr, e));
  const u = kid(rPr, "w:u"); if (u) r.u = av(u, "w:val") !== "none";
  const sz = numAttr(kid(rPr, "w:sz"), "w:val"); if (sz) r.sz = sz;
  const c = av(kid(rPr, "w:color"), "w:val"); if (c) r.color = c === "auto" ? "000000" : c;
  const v = av(kid(rPr, "w:vertAlign"), "w:val"); if (v) r.vert = v;
  const cs = av(kid(rPr, "w:rStyle"), "w:val"); if (cs) r.rStyle = cs;
  return r;
}
function loadStyles(root) {
  const dd = kid(root, "w:docDefaults"), S = { byId: {}, defP: pProps(kid(kid(dd, "w:pPrDefault"), "w:pPr")), defR: rProps(kid(kid(dd, "w:rPrDefault"), "w:rPr")), defPara: null };
  for (const s of kids(root, "w:style")) {
    const id = av(s, "w:styleId"), type = av(s, "w:type");
    S.byId[id] = { id, type, based: av(kid(s, "w:basedOn"), "w:val"), p: pProps(kid(s, "w:pPr")), r: rProps(kid(s, "w:rPr")), tbl: kid(s, "w:tblPr") };
    if (av(s, "w:default") === "1" && type === "paragraph") S.defPara = id;
  }
  return S;
}
function chain(S, id) { const out = [], seen = new Set(); while (id && S.byId[id] && !seen.has(id)) { seen.add(id); out.unshift(S.byId[id]); id = S.byId[id].based; } return out; }
const styleP = (S, id) => chain(S, id || S.defPara).reduce((a, s) => mergeP(a, s.p), S.defP);
const styleR = (S, pid, cid) => { let r = { ...S.defR }; for (const s of chain(S, pid || S.defPara)) r = { ...r, ...s.r }; for (const s of chain(S, cid)) r = { ...r, ...s.r }; return r; };

/* ---------- 5. Fonts (metric-compatible, embedded as subsets) ---------- */
const SERIF = /times|cambria|georgia|garamond|palatino|book antiqua|century|minion|constantia|serif$/i, MONO = /courier|consolas|mono|lucida console|menlo/i;
const EXACT = /^(times new roman|arial|courier new)$/i;
function fontKey(env, r) {
  let fam = r.font || (r.theme && env.theme[/major/i.test(r.theme) ? "major" : "minor"]) || "Arial";
  const cls = MONO.test(fam) ? "mono" : SERIF.test(fam) && !/sans/i.test(fam) ? "serif" : "sans";
  if (!EXACT.test(fam)) env.warn.add(`Font "${fam}" is not available here, so a similar font is used and line breaks may differ slightly.`);
  const name = { serif: "Tinos", sans: "Arimo", mono: "Cousine" }[cls] + "-" + ({ "00": "Regular", "10": "Bold", "01": "Italic", "11": "BoldItalic" })[(r.b ? 1 : 0) + "" + (r.i ? 1 : 0)];
  env.need.add(name);
  return { name, cls };
}

/* ---------- 6. Build blocks from OOXML ---------- */
function buildSect(el, env) {
  const g = (n, k) => numAttr(kid(el, n), k);
  const pg = kid(el, "w:pgSz"), mar = kid(el, "w:pgMar");
  const refs = {};
  for (const h of kids(el, "w:headerReference")) refs["h" + h.a["w:type"]] = h.a["r:id"];
  for (const f of kids(el, "w:footerReference")) refs["f" + f.a["w:type"]] = f.a["r:id"];
  const s = { W: (numAttr(pg, "w:w") || 11906) * TW, H: (numAttr(pg, "w:h") || 16838) * TW, refs, titlePg: !!kid(el, "w:titlePg"), start: g("w:pgNumType", "w:start"), brk: av(kid(el, "w:type"), "w:val") || "nextPage" };
  const m = k => (numAttr(mar, "w:" + k) ?? 1440) * TW;
  Object.assign(s, { top: m("top"), bottom: m("bottom"), left: m("left"), right: m("right"), header: m("header"), footer: m("footer") });
  if (av(pg, "w:orient") === "landscape" && s.W < s.H) [s.W, s.H] = [s.H, s.W];
  return s;
}
function buildBlocks(env, nodes, counter) {
  const out = [];
  for (const n of nodes) {
    if (n.n === "w:p") {
      out.push(buildPara(env, n, counter));
      const sp = kid(kid(n, "w:pPr"), "w:sectPr");
      if (sp) out.push({ type: "sect", sect: buildSect(sp, env) });
    } else if (n.n === "w:tbl") out.push(buildTable(env, n, counter));
    else if (n.n === "w:sdt") out.push(...buildBlocks(env, kid(n, "w:sdtContent")?.c.filter(x => typeof x !== "string") || [], counter));
  }
  return out;
}
function buildPara(env, pNode, counter) {
  const pPrN = kid(pNode, "w:pPr"), direct = pProps(pPrN);
  const base = styleP(env.S, direct.pStyle), numPr = direct.numPr || base.numPr;
  const markR = { ...styleR(env.S, direct.pStyle), ...rProps(kid(pPrN, "w:rPr")) };
  let pr = base, label = null;
  if (numPr && numPr.numId && numPr.numId !== "0") {
    const r = counter.next(numPr.numId, numPr.ilvl);
    if (r) { const lp = pProps(r.lvl.pPr); pr = mergeP(pr, { ind: lp.ind, tabs: lp.tabs }); label = { text: r.label, suff: r.lvl.suff, r: { ...markR, ...rProps(r.lvl.rPr) } }; fontKey(env, label.r); }
  }
  pr = mergeP(pr, direct);
  fontOf(env, markR);
  const items = [], rBase = styleR(env.S, direct.pStyle);
  inlines(env, pNode, rBase, items, { fld: null });
  return { type: "p", pr, sid: direct.pStyle || env.S.defPara, items, label, markR };
}
function inlines(env, node, rBase, items, st) {
  for (const c of node.c) {
    if (typeof c === "string") continue;
    if (c.n === "w:r") runItems(env, c, rBase, items, st);
    else if (c.n === "w:hyperlink" || c.n === "w:smartTag" || c.n === "w:ins" || c.n === "w:sdtContent" || c.n === "w:customXml") inlines(env, c, rBase, items, st);
    else if (c.n === "w:sdt") inlines(env, kid(c, "w:sdtContent") || c, rBase, items, st);
    else if (c.n === "w:fldSimple") { const f = /\b(PAGE|NUMPAGES)\b/.exec(c.a["w:instr"] || ""); if (f) items.push({ k: "fld", f: f[1], ...fontOf(env, { ...rBase, ...rProps(kid(kid(c, "w:r"), "w:rPr")) }) }); else inlines(env, c, rBase, items, st); }
  }
}
function fontOf(env, r) { const { name, cls } = fontKey(env, r); let sz = (r.sz || 22) / 2, rise = 0; if (r.vert === "superscript") { rise = sz * .35; sz *= .65; } else if (r.vert === "subscript") { rise = -sz * .15; sz *= .65; } return { fk: name, cls, sz, rise, r }; }
function runItems(env, run, rBase, items, st) {
  const own = rProps(kid(run, "w:rPr")); let r = { ...rBase };
  for (const cs of chain(env.S, own.rStyle)) r = { ...r, ...cs.r };
  r = { ...r, ...own };
  if (r.hide) return;
  for (const c of run.c) {
    if (typeof c === "string") continue;
    switch (c.n) {
      case "w:t": if (st.fld && st.fld.phase === "result" && st.fld.swallow) break; words(env, r, c.c.join(""), items); break;
      case "w:tab": items.push({ k: "tab", ...fontOf(env, r) }); break;
      case "w:br": items.push({ k: av(c, "w:type") === "page" || av(c, "w:type") === "column" ? "pb" : "br", ...fontOf(env, r) }); break;
      case "w:cr": items.push({ k: "br", ...fontOf(env, r) }); break;
      case "w:noBreakHyphen": words(env, r, "-", items); break;
      case "w:lastRenderedPageBreak": items.push({ k: "soft" }); break;
      case "w:instrText": if (st.fld) st.fld.instr += c.c.join(""); break;
      case "w:fldChar": {
        const t = av(c, "w:fldCharType");
        if (t === "begin") st.fld = { instr: "", phase: "instr", swallow: false };
        else if (t === "separate" && st.fld) { st.fld.phase = "result"; const f = /\b(PAGE|NUMPAGES)\b/.exec(st.fld.instr); if (f) { st.fld.swallow = true; items.push({ k: "fld", f: f[1], ...fontOf(env, r) }); } }
        else if (t === "end") st.fld = null;
        break;
      }
      case "w:drawing": case "w:pict": image(env, c, items); break;
      case "w:footnoteReference": case "w:endnoteReference": env.warn.add("Footnotes and endnotes are not drawn yet."); break;
    }
  }
}
function words(env, r, text, items) {
  const f = fontOf(env, r); if (r.caps) text = text.toUpperCase();
  for (const m of text.matchAll(/([ \t\r\n]+)|([^ \t\r\n]+)/g)) items.push(m[1] ? { k: "sp", s: " ".repeat(m[1].length), ...f } : { k: "w", s: m[2].replace(/\u00a0/g, " ").replace(/\u00ad/g, ""), ...f });
}
function findDeep(n, name, acc = []) { for (const c of n.c) if (typeof c !== "string") { if (c.n === name) acc.push(c); findDeep(c, name, acc); } return acc; }
function image(env, node, items) {
  const blip = findDeep(node, "a:blip")[0], ext = findDeep(node, "wp:extent")[0];
  if (!blip || !ext) { env.warn.add("A drawing, shape or text box could not be drawn."); return; }
  if (findDeep(node, "wp:anchor").length) env.warn.add("Floating pictures are placed in the text line.");
  const target = env.rels[blip.a["r:embed"]], data = target && env.media[target];
  const mime = /\.png$/i.test(target || "") ? "png" : /\.jpe?g$/i.test(target || "") ? "jpg" : null;
  if (!data || !mime) { env.warn.add("A picture in an unsupported format was skipped."); return; }
  items.push({ k: "img", w: +ext.a.cx / 12700, h: +ext.a.cy / 12700, data, mime });
}
function buildTable(env, tbl, counter) {
  const tblPr = kid(tbl, "w:tblPr"), sid = av(kid(tblPr, "w:tblStyle"), "w:val");
  const bd = {};
  for (const t of [...chain(env.S, sid).map(s => s.tbl), tblPr]) { const b = kid(t, "w:tblBorders"); if (b) for (const e of b.c) if (typeof e !== "string") bd[e.n.slice(2)] = !/^(nil|none)$/.test(e.a["w:val"] || "single"); }
  const grid = kids(kid(tbl, "w:tblGrid"), "w:gridCol").map(g => +g.a["w:w"] * TW);
  const rows = kids(tbl, "w:tr").map(tr => {
    const cells = []; let gi = 0;
    for (const tc of kids(tr, "w:tc")) {
      const pr = kid(tc, "w:tcPr"), span = numAttr(kid(pr, "w:gridSpan"), "w:val") || 1;
      const w = grid.slice(gi, gi + span).reduce((a, b) => a + b, 0) || (numAttr(kid(pr, "w:tcW"), "w:w") || 2000) * TW;
      const own = kid(pr, "w:tcBorders"), borders = {}, any = Object.values(bd).some(Boolean);
      for (const e of ["top", "left", "bottom", "right"]) { const o = kid(own, "w:" + e); borders[e] = o ? !/^(nil|none)$/.test(o.a["w:val"]) : any; }
      const sub = kids(tc, "w:p").length || kids(tc, "w:tbl").length ? tc.c.filter(x => typeof x !== "string") : [];
      if (kids(tc, "w:tbl").length) env.warn.add("Tables inside table cells are skipped.");
      cells.push({ x0: grid.slice(0, gi).reduce((a, b) => a + b, 0), w, blocks: buildBlocks(env, sub.filter(x => x.n !== "w:tbl"), counter), borders, fill: av(kid(pr, "w:shd"), "w:fill"), vmerge: kid(pr, "w:vMerge") && av(kid(pr, "w:vMerge"), "w:val") !== "restart" });
      gi += span;
    }
    return { cells };
  });
  return { type: "t", rows, grid, border: bd };
}

/* ---------- 7. Layout ---------- */
const hexRgb = (PL, h) => { const n = parseInt(h || "000000", 16); return PL.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); };
function widthOf(env, it, s = it.s) {
  const key = it.fk + "|" + it.sz + "|" + s, c = env.wc.get(key); if (c !== undefined) return c;
  const w = env.fonts[it.fk].widthOfTextAtSize(env.clean(it.fk, s), it.sz); env.wc.set(key, w); return w;
}
function layoutPara(env, b, W, fields) {
  const pr = b.pr, ind = pr.ind || {}, L = (ind.left || 0) * TW, R = (ind.right || 0) * TW, hang = (ind.hanging || 0) * TW, firstLine = (ind.firstLine || 0) * TW;
  const right = W - R, startX = L - hang + firstLine, jc = pr.jc, dtab = 720 * TW;
  const stops = (pr.tabs || []).map(t => ({ pos: t.pos * TW, val: t.val })).sort((a, c) => a.pos - c.pos);
  const nextStop = x => { let best = { pos: Math.ceil((x + .01) / dtab) * dtab, val: "left" }; if (hang > 0 && L > x + .01 && L < best.pos) best = { pos: L, val: "left" }; for (const s of stops) if (s.pos > x + .01 && s.pos < best.pos) best = s; return best; };
  const src = b.items.map(i => i.k === "fld" ? { k: "w", s: String(fields[i.f] ?? 0), fk: i.fk, cls: i.cls, sz: i.sz, rise: i.rise, r: i.r } : i);
  const lines = []; let line, x, pbNext = false, softNext = false;
  const newLine = first => { line = { items: [], first, h: 0, desc: 0, nat: 0, end: "wrap", pb: pbNext, soft: softNext }; pbNext = softNext = false; x = first ? startX : L; lines.push(line); };
  newLine(true);
  const put = (it, w) => { it.x = x; it.w = w; line.items.push(it); x += w; if (it.k === "w" || it.k === "tab" || it.k === "img") { const nat = it.k === "img" ? it.h : natural(it.cls, it.sz); line.nat = Math.max(line.nat, nat); line.desc = Math.max(line.desc, it.k === "img" ? 0 : METRICS[it.cls].desc * it.sz); } };
  const endLine = end => { line.end = end; };
  if (b.label) {
    const li = { k: "w", lbl: true, s: b.label.text, ...fontOf(env, b.label.r) }, w = widthOf(env, li);
    put(li, w);
    if (b.label.suff === "tab") { const s = nextStop(x); x = Math.max(x, s.pos); } else if (b.label.suff === "space") x += widthOf(env, { ...li, s: " " });
  }
  for (let i = 0; i < src.length; i++) {
    const it = src[i];
    if (it.k === "soft") { if (line.items.filter(q => q.k !== "sp" && !q.lbl).length) { endLine("wrap"); newLine(false); softNext = false; line.soft = true; } else line.soft = true; continue; }
    if (it.k === "br") { endLine("br"); newLine(false); continue; }
    if (it.k === "pb") { if (i === src.length - 1) { b.pbAfter = true; continue; } endLine("br"); pbNext = true; newLine(false); continue; }
    if (it.k === "tab") {
      const s = nextStop(x); let tx = s.pos;
      if (s.val === "right" || s.val === "center") { let w = 0; for (let j = i + 1; j < src.length && src[j].k !== "tab" && src[j].k !== "br" && src[j].k !== "pb"; j++) if (src[j].k === "w" || src[j].k === "sp") w += widthOf(env, src[j]); tx = Math.max(x, s.pos - (s.val === "center" ? w / 2 : w)); }
      put({ ...it, k: "tab" }, Math.max(0, Math.min(tx, right) - x)); continue;
    }
    if (it.k === "img") { const w = Math.min(it.w, Math.max(right - L, 10)); if (x + w > right && line.items.some(q => q.k !== "sp")) { endLine("wrap"); newLine(false); } put({ ...it, h: it.h * (w / it.w), w }, w); continue; }
    const w = it.k === "sp" ? widthOf(env, { ...it, s: " " }) * it.s.length : widthOf(env, it);
    if (it.k === "sp") { put(it, w); continue; }
    if (x + w > right + .01 && jc === "both") { // Word squeezes spaces up to ~25% before wrapping a justified line
      let seen = false, cnt = 0; for (const q of line.items) { if (q.k === "w") seen = true; else if (q.k === "sp" && seen) cnt += q.s.length; }
      if (x + w - right <= cnt * widthOf(env, { ...it, s: " " }) * SQUEEZE) { put(it, w); continue; }
    }
    if (x + w > right + .01) {
      const hasContent = line.items.some(q => q.k === "w" || q.k === "img");
      if (hasContent) { while (line.items.length && line.items[line.items.length - 1].k === "sp") { const sp = line.items.pop(); x -= sp.w; } endLine("wrap"); newLine(false); }
      if (x + w > right + .01) { // longer than a whole line: split by characters
        let rest = it.s;
        while (rest) { let n = 1; while (n < rest.length && x + widthOf(env, it, rest.slice(0, n + 1)) <= right) n++; const part = rest.slice(0, n); put({ ...it, s: part }, widthOf(env, it, part)); rest = rest.slice(n); if (rest) { endLine("wrap"); newLine(false); } }
        continue;
      }
    }
    put(it, w);
  }
  // line heights (paragraph mark counts on the last line), then alignment
  const mf = fontOf(env, b.markR), markNat = natural(mf.cls, mf.sz);
  lines.forEach((ln, idx) => {
    let nat = ln.nat, desc = ln.desc;
    if (!nat || idx === lines.length - 1) { nat = Math.max(nat, markNat); desc = Math.max(desc, METRICS[mf.cls].desc * mf.sz); }
    const sp = pr.sp || {}, auto = sp.line === undefined || (sp.lineRule || "auto") === "auto";
    ln.h = sp.line === undefined ? nat : sp.lineRule === "exact" ? +sp.line * TW : sp.lineRule === "atLeast" ? Math.max(nat, +sp.line * TW) : nat * (+sp.line) / 240;
    ln.desc = desc;
    const vis = ln.items.filter(q => !(q.k === "sp" && q === ln.items[ln.items.length - 1]));
    const endX = vis.length ? vis[vis.length - 1].x + vis[vis.length - 1].w : 0, slack = right - endX;
    const last = idx === lines.length - 1 || ln.end === "br";
    if (jc === "both" && slack !== 0 && (slack < 0 || !last)) {
      const sps = []; let seen = false;
      for (const q of ln.items) { if (q.k === "w") seen = true; else if (q.k === "sp" && seen && q !== ln.items[ln.items.length - 1]) sps.push(q); }
      const cnt = sps.reduce((a, q) => a + q.s.length, 0);
      if (cnt) { let sh = 0; const per = slack / cnt; for (const q of ln.items) { q.x += sh; if (sps.includes(q)) { q.w += per * q.s.length; sh += per * q.s.length; } } }
    } else if ((jc === "right" || jc === "center") && slack > 0 && vis.length) { const d = jc === "right" ? slack : slack / 2; for (const q of ln.items) q.x += d; }
  });
  return lines;
}
function layoutTable(env, t, W, counter) {
  const total = t.grid.reduce((a, b) => a + b, 0), k = total > W ? W / total : 1, PAD = 5.4;
  return t.rows.map(row => {
    let h = 0;
    const cells = row.cells.map(c => {
      const cw = c.w * k, inner = Math.max(cw - 2 * PAD, 10); let ch = 0, prev = null; const lines = [];
      for (const b of c.blocks) if (b.type === "p") {
        const ls = layoutPara(env, b, inner, { PAGE: 0, NUMPAGES: 0 }), sp = b.pr.sp || {};
        const g = spacingGap(prev, { sid: b.sid, ctx: b.pr.ctx, before: (+sp.before || 0) * TW });
        ch += g; ls.forEach(l => { lines.push({ ...l, dy: ch }); ch += l.h; }); prev = { sid: b.sid, ctx: b.pr.ctx, after: (+sp.after || 0) * TW };
      }
      if (prev) ch += prev.after; h = Math.max(h, ch);
      return { ...c, x: c.x0 * k, w: cw, lines, PAD };
    });
    return { cells, h: Math.max(h, 12) };
  });
}

/* ---------- 8. Pagination ---------- */
function partBlocks(env, rid) { if (!rid || !env.partXml[rid]) return []; const c = env.parts[rid] ??= buildBlocks(env, env.partXml[rid].c.filter(x => typeof x !== "string"), createCounter(env.num)); return c; }
function partHeight(env, blocks, W, fields) { let y = 0, prev = null; for (const b of blocks) if (b.type === "p") { const sp = b.pr.sp || {}; y += spacingGap(prev, { sid: b.sid, ctx: b.pr.ctx, before: (+sp.before || 0) * TW }); for (const l of layoutPara(env, b, W, fields)) y += l.h; prev = { sid: b.sid, ctx: b.pr.ctx, after: (+sp.after || 0) * TW }; } return y + (prev ? prev.after : 0); }
function variantOf(sec, first, pageNo, evenOdd) { return sec.titlePg && first ? "first" : evenOdd && pageNo % 2 === 0 ? "even" : "default"; }
function paginate(env, blocks) {
  const pages = []; let pg = null, prevSec = null, lastPageNo = 0;
  const newPage = (sec, soft) => {
    const first = !pg || pg.sec !== sec, pageNo = first && sec.start !== undefined ? sec.start : lastPageNo + 1; lastPageNo = pageNo;
    const variant = variantOf(sec, first, pageNo, env.evenOdd), W = sec.W - sec.left - sec.right;
    const hd = partBlocks(env, sec.refs["h" + variant]), ft = partBlocks(env, sec.refs["f" + variant]), fields = { PAGE: pageNo, NUMPAGES: 0 };
    const hh = partHeight(env, hd, W, fields), fh = partHeight(env, ft, W, fields);
    pg = { sec, no: pageNo, W, top: Math.max(sec.top, sec.header + hh), bottom: sec.H - Math.max(sec.bottom, sec.footer + fh), items: [], vis: false, prev: null, variant, hd, ft, idx: pages.length };
    pg.y = pg.top; pages.push(pg);
  };
  const place = (it, h, gap = 0) => { pg.y += gap; it.y = pg.y; pg.items.push(it); pg.y += h; };
  for (let bi = 0; bi < blocks.length; bi++) {
    const b = blocks[bi]; if (b.type === "sect") continue;
    if (!pg || (b.sec !== pg.sec && b.sec.brk !== "continuous")) newPage(b.sec);
    if (b.type === "t") {
      const rows = layoutTable(env, b, pg.W, null);
      if (pg.prev) pg.y += pg.prev.after;
      for (const r of rows) { if (pg.y + r.h > pg.bottom && pg.items.length) newPage(b.sec); place({ t: "row", r, x0: pg.sec.left }, r.h); pg.vis = true; }
      pg.prev = null; continue;
    }
    const lines = layoutPara(env, b, pg.W, { PAGE: pg.no, NUMPAGES: 0 }), sp = b.pr.sp || {};
    const cur = { sid: b.sid, ctx: b.pr.ctx, before: (+sp.before || 0) * TW, after: (+sp.after || 0) * TW };
    if (b.pr.pbb && pg.items.length) newPage(b.sec);
    const gapOf = () => pg.items.length === 0 && pg.idx > 0 ? 0 : spacingGap(pg.prev, cur);
    if (b.pr.keepNext) { const nx = blocks.slice(bi + 1).find(q => q.type !== "sect"); if (nx && nx.type === "p") { const nl = layoutPara(env, nx, pg.W, { PAGE: pg.no, NUMPAGES: 0 }); const need = gapOf() + lines.reduce((a, l) => a + l.h, 0) + (nl[0] ? nl[0].h : 0); if (pg.y + need > pg.bottom && pg.items.length) newPage(b.sec); } }
    let j = 0;
    while (j < lines.length) {
      const ln = lines[j];
      if (ln.pb && (pg.items.length || j > 0)) newPage(b.sec);
      else if (ln.soft) { if (pg.vis) newPage(b.sec); else { pg.items = []; pg.y = pg.top; pg.prev = null; } }
      let n = 1; while (j + n < lines.length && !lines[j + n].pb && !lines[j + n].soft) n++;
      const seg = lines.slice(j, j + n).map(l => l.h), gap = j === 0 ? gapOf() : 0;
      const k = linesThatFit(seg, pg.bottom - pg.y - gap, { widow: b.pr.widow !== false, keepLines: !!b.pr.keepLines, atStart: j === 0 });
      if (k === 0) { if (pg.items.length) { newPage(b.sec); continue; } place({ t: "line", l: lines[j], b }, lines[j].h, gap); j++; continue; }
      let first = true;
      for (let q = 0; q < k; q++) { const l = lines[j + q]; place({ t: "line", l, b }, l.h, first ? gap : 0); first = false; }
      pg.items.slice(-k).forEach(it => { if (it.l.items.some(q => q.k !== "sp")) pg.vis = true; });
      j += k;
      if (k < n && j < lines.length) newPage(b.sec);
    }
    pg.prev = cur;
    if (b.pbAfter) { newPage(b.sec); }
  }
  if (!pages.length) newPage(blocks.find(b => b.type === "sect")?.sect || defaultSect());
  return pages;
}
const defaultSect = () => ({ W: 595.28, H: 841.89, refs: {}, top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36, brk: "nextPage" });

/* ---------- 9. Drawing ---------- */
function drawLine(env, page, H, l, x0, top) {
  const base = top + l.h - l.desc, PL = env.PL;
  l.items.forEach((it, i) => {
    if (it.k === "img") { const im = env.img(it); page.drawImage(im, { x: x0 + it.x, y: H - base, width: it.w, height: it.h }); return; }
    if (it.k !== "w" && !(it.k === "sp" && it.r && (it.r.u || it.r.strike))) return;
    const x = x0 + it.x, y = H - base + (it.rise || 0), color = hexRgb(PL, it.r && it.r.color);
    if (it.k === "w") { const next = l.items[i + 1], s = env.clean(it.fk, it.s) + (next && next.k === "sp" ? " " : ""); page.drawText(s, { x, y, size: it.sz, font: env.fonts[it.fk], color }); }
    if (it.r && it.r.u) page.drawLine({ start: { x, y: y - it.sz * .1 }, end: { x: x + it.w, y: y - it.sz * .1 }, thickness: it.sz * .05, color });
    if (it.r && it.r.strike) page.drawLine({ start: { x, y: y + it.sz * .3 }, end: { x: x + it.w, y: y + it.sz * .3 }, thickness: it.sz * .05, color });
  });
}
function drawPart(env, page, pg, blocks, y0, fields) {
  let y = y0, prev = null, W = pg.W;
  for (const b of blocks) if (b.type === "p") {
    const sp = b.pr.sp || {}; y += spacingGap(prev, { sid: b.sid, ctx: b.pr.ctx, before: (+sp.before || 0) * TW });
    for (const l of layoutPara(env, b, W, fields)) { drawLine(env, page, pg.sec.H, l, pg.sec.left, y); y += l.h; }
    prev = { sid: b.sid, ctx: b.pr.ctx, after: (+sp.after || 0) * TW };
  }
}
function drawRow(env, page, H, it) {
  const PL = env.PL; let y = it.y;
  for (const c of it.r.cells) {
    const x = it.x0 + c.x;
    if (c.fill && c.fill !== "auto") page.drawRectangle({ x, y: H - y - it.r.h, width: c.w, height: it.r.h, color: hexRgb(PL, c.fill) });
    const seg = (a, b, c2, d) => page.drawLine({ start: { x: a, y: H - b }, end: { x: c2, y: H - d }, thickness: .5, color: PL.rgb(0, 0, 0) });
    if (c.borders.top) seg(x, y, x + c.w, y); if (c.borders.bottom) seg(x, y + it.r.h, x + c.w, y + it.r.h);
    if (c.borders.left) seg(x, y, x, y + it.r.h); if (c.borders.right) seg(x + c.w, y, x + c.w, y + it.r.h);
    for (const l of c.lines) drawLine(env, page, H, l, x + c.PAD, y + l.dy);
  }
}

/* ---------- 10. Entry point ---------- */
export async function docxToPdf(buf, { PDFLib, JSZip, fontkit, loadFont, status = () => {} }) {
  const zip = await JSZip.loadAsync(buf), rd = async p => { const f = zip.file(p); return f ? f.async("string") : null; };
  const docXml = await rd("word/document.xml");
  if (!docXml) throw new Error("This is not a valid .docx file.");
  const env = { PL: PDFLib, S: loadStyles(parseXml(await rd("word/styles.xml") || "<w:styles/>")), warn: new Set(), need: new Set(), wc: new Map(), parts: {}, partXml: {}, rels: {}, media: {}, theme: { minor: "Calibri", major: "Calibri" }, fonts: {}, evenOdd: false };
  const numXml = await rd("word/numbering.xml"); env.num = numXml ? loadNumbering(parseXml(numXml)) : { abs: {}, nums: {} };
  const th = await rd("word/theme/theme1.xml"); if (th) { env.theme.minor = (/<a:minorFont>\s*<a:latin typeface="([^"]+)"/.exec(th) || [])[1] || "Calibri"; env.theme.major = (/<a:majorFont>\s*<a:latin typeface="([^"]+)"/.exec(th) || [])[1] || "Calibri"; }
  env.evenOdd = /<w:evenAndOddHeaders\b/.test(await rd("word/settings.xml") || "");
  for (const rel of kids(parseXml(await rd("word/_rels/document.xml.rels") || "<Relationships/>"), "Relationship")) {
    const t = rel.a.Target.replace(/^\/?(word\/)?/, "word/"); env.rels[rel.a.Id] = t;
    if (/header|footer/.test(rel.a.Type)) { const x = await rd(t); if (x) env.partXml[rel.a.Id] = parseXml(x); }
  }
  for (const name of Object.keys(zip.files)) if (/^word\/media\//.test(name)) env.media[name] = await zip.file(name).async("uint8array");
  for (const rid in env.partXml) { const rp = env.rels[rid].replace(/([^/]+)$/, "_rels/$1.rels"), rx = await rd(rp); if (rx) for (const rel of kids(parseXml(rx), "Relationship")) env.rels[rel.a.Id] = env.rels[rel.a.Id] || rel.a.Target.replace(/^\/?(word\/)?/, "word/"); }
  const body = kid(parseXml(docXml), "w:body");
  const blocks = buildBlocks(env, body.c.filter(x => typeof x !== "string" && x.n !== "w:sectPr"), createCounter(env.num));
  const finalSect = buildSect(kid(body, "w:sectPr") || { n: "w:sectPr", a: {}, c: [] }, env);
  // assign each block its section (sections end at their sectPr); inherit header/footer refs
  let pending = []; const secs = [];
  for (const b of blocks) { if (b.type === "sect") { secs.push(b.sect); pending.forEach(x => x.sec = b.sect); pending = []; } else pending.push(b); }
  pending.forEach(x => x.sec = finalSect); secs.push(finalSect);
  secs.forEach((s, i) => { if (i) s.refs = { ...secs[i - 1].refs, ...s.refs }; });
  // preload every font the document uses and embed as subsets
  blocks.forEach(b => { if (b.type === "p" && b.label) fontKey(env, b.label.r); });
  for (const rid in env.partXml) partBlocks(env, rid);
  status("Loading fonts…");
  const pdf = await PDFLib.PDFDocument.create(); pdf.registerFontkit(fontkit);
  pdf.setProducer("converter.io"); pdf.setCreator("converter.io");
  const names = [...env.need]; for (const n of ["Tinos-Regular"]) if (!names.length) names.push(n);
  env.charsets = {};
  for (const n of names) { env.fonts[n] = await pdf.embedFont(await loadFont(n), { subset: false }); env.charsets[n] = new Set(env.fonts[n].getCharacterSet()); }
  const bad = new Set();
  env.clean = (fk, s) => { const cs = env.charsets[fk]; let o = ""; for (const ch of s) { if (cs.has(ch.codePointAt(0))) o += ch; else { o += "?"; bad.add(ch); } } return o; };
  env.img = it => it.embedded;
  const walk = list => list.flatMap(b => b.type === "t" ? b.rows.flatMap(r => r.cells.flatMap(c => walk(c.blocks))) : b.type === "p" ? [b] : []);
  for (const b of walk([...blocks, ...Object.values(env.parts).flat()])) for (const it of b.items) if (it.k === "img") it.embedded = it.mime === "png" ? await pdf.embedPng(it.data) : await pdf.embedJpg(it.data);
  status("Laying out pages…");
  const pages = paginate(env, blocks.filter(b => b.sec));
  const N = pages.length;
  status("Writing PDF…");
  for (const pg of pages) {
    const page = pdf.addPage([pg.sec.W, pg.sec.H]), H = pg.sec.H, fields = { PAGE: pg.no, NUMPAGES: N };
    drawPart(env, page, pg, pg.hd, pg.sec.header, fields);
    if (pg.ft.length) drawPart(env, page, pg, pg.ft, pg.sec.H - pg.sec.footer - partHeight(env, pg.ft, pg.W, fields), fields);
    for (const it of pg.items) { if (it.t === "line") drawLine(env, page, H, it.l, pg.sec.left, it.y); else drawRow(env, page, H, it); }
  }
  if (bad.size) env.warn.add(`${bad.size} character type(s) such as "${[...bad].slice(0, 3).join(" ")}" are not in the built-in fonts and were replaced with "?". Sinhala and Tamil text is not supported in Word to PDF yet.`);
  return { bytes: await pdf.save({ useObjectStreams: true }), pages: N, warnings: [...env.warn] };
}
