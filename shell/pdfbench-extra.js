/* YourOfficeSuite PDF Bench, part 2: Redact, Compare, PDF to Text, Crop, Flatten, Black & white,
   and the built-in Edit and annotate tool. Everything runs on this computer with the libraries in shell/vendor. */
const K = window.YOSPDF.kit;
const { pdfjs, readFile, loadLib, loadView, saveOne, saveOut, replaceOption, renderPage, place, visualSize, canvasBlob, blobBytes, pad, btn, field, select, radios, dropZone, runner, workspace, guard, takeInitial, signaturePad } = K;
const { PDFDocument, PDFName, PDFDict, PDFArray, PDFNumber, PDFRef, PDFString, PDFHexString, PDFTextField, PDFCheckBox, PDFDropdown, PDFOptionList, PDFRadioGroup,
  StandardFonts, BlendMode, LineCapStyle, rgb, degrees, pushGraphicsState, popGraphicsState, concatTransformationMatrix, drawObject } = window.PDFLib;
const Y = window.YOS;
const { api, el, icon, toast, errText, fmtSize, baseName, stemOf } = Y;
const { Util, OPS, ImageKind } = pdfjs;
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const rid = () => Math.random().toString(36).slice(2, 9);
const plural = (n, word, many) => `${n} ${n === 1 ? word : many || `${word}s`}`;

// ================================================================ operations (same logic as the online tools)

// ---------- Redact ----------
const EMAIL = /[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+/g;
const PHONE = /\+?\d[\d ().-]{7,}\d/g;
const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// Finds words, email addresses and phone numbers on every page. Boxes are fractions of the page as seen (0 to 1, from the top left).
async function findText(bytes, name, opts, tick) {
  const view = await loadView(bytes, name); const res = []; let textPages = 0;
  const mctx = document.createElement('canvas').getContext('2d'); const pats = [];
  for (const w of opts.words.map((x) => x.trim()).filter(Boolean)) pats.push(new RegExp(escRe(w), 'gi'));
  if (opts.emails) pats.push(EMAIL);
  if (opts.phones) pats.push(PHONE);
  const pages = view.numPages;
  try {
    for (let n = 1; n <= pages; n++) {
      const page = await view.getPage(n); const vp = page.getViewport({ scale: 1 });
      const tc = await page.getTextContent(); const styles = tc.styles || {};
      const items = tc.items.filter((it) => typeof it.str === 'string' && it.str.trim());
      if (items.length) textPages++;
      for (const it of items) {
        const str = it.str; if (!str || !it.width) continue;
        const tx = Util.transform(vp.transform, it.transform);
        const fh = Math.hypot(tx[2], tx[3]) || Math.abs(it.height) || 10; // font height on screen
        const len = Math.hypot(tx[0], tx[1]) || 1; const ux = tx[0] / len; const uy = tx[1] / len; // text direction
        const nx = -uy; const ny = ux;
        const full = it.width * vp.scale;
        // where a match starts and ends inside the item, measured with a similar font (letters differ in width)
        const fam = (styles[it.fontName] && styles[it.fontName].fontFamily) || 'sans-serif';
        mctx.font = `100px ${fam}`; const total = mctx.measureText(str).width || str.length;
        const at = (k) => (k <= 0 ? 0 : k >= str.length ? full : (mctx.measureText(str.slice(0, k)).width / total) * full);
        for (const re of pats) {
          re.lastIndex = 0; let m;
          while ((m = re.exec(str))) {
            if (!m[0].length) { re.lastIndex++; continue; }
            const padA = 0.12 * fh; const a = Math.max(0, at(m.index) - padA); const b = Math.min(full, at(m.index + m[0].length) + padA);
            const pts = [];
            for (const t of [a, b]) for (const up of [-0.28 * fh, 1.0 * fh]) pts.push([tx[4] + ux * t - nx * up, tx[5] + uy * t - ny * up]);
            const xs = pts.map((p) => p[0]); const ys = pts.map((p) => p[1]);
            const x0 = Math.max(0, Math.min(...xs) - 1); const x1 = Math.min(vp.width, Math.max(...xs) + 1);
            const y0 = Math.max(0, Math.min(...ys) - 1); const y1 = Math.min(vp.height, Math.max(...ys) + 1);
            if (x1 > x0 && y1 > y0) res.push({ id: rid(), page: n - 1, x: x0 / vp.width, y: y0 / vp.height, w: (x1 - x0) / vp.width, h: (y1 - y0) / vp.height, auto: true });
          }
        }
      }
      page.cleanup(); if (tick) tick(n / pages);
    }
  } finally { await view.destroy(); }
  return { boxes: res, textPages, pages };
}

// Burns the boxes in. Pages with a box are rebuilt as images, so what is under a box is no longer in the file.
async function opRedact(bytes, name, { boxes, color }, tick) {
  const src = await loadLib(bytes, name); const view = await loadView(bytes, name); const out = await PDFDocument.create();
  const n = src.getPageCount(); const hit = new Set(boxes.map((b) => b.page));
  const keep = Array.from({ length: n }, (_, i) => i).filter((i) => !hit.has(i)); const copied = new Map();
  if (keep.length) { const pages = await out.copyPages(src, keep); keep.forEach((i, k) => copied.set(i, pages[k])); }
  try {
    for (let i = 0; i < n; i++) {
      if (!hit.has(i)) { out.addPage(copied.get(i)); if (tick) tick((i + 1) / n); continue; }
      const r = await renderPage(view, i + 1, { scale: 200 / 72, maxDim: 3400, forms: true }); // about 200 DPI
      const g = r.canvas.getContext('2d'); g.fillStyle = color === 'white' ? '#ffffff' : '#000000';
      const W = r.canvas.width; const H = r.canvas.height;
      for (const b of boxes) if (b.page === i) g.fillRect(Math.floor(b.x * W), Math.floor(b.y * H), Math.ceil(b.w * W) + 1, Math.ceil(b.h * H) + 1);
      const img = await out.embedJpg(await blobBytes(await canvasBlob(r.canvas, 'image/jpeg', 0.9)));
      out.addPage([r.ptW, r.ptH]).drawImage(img, { x: 0, y: 0, width: r.ptW, height: r.ptH });
      r.canvas.width = r.canvas.height = 1; if (tick) tick((i + 1) / n);
    }
  } finally { await view.destroy(); }
  out.setProducer('YourOfficeSuite'); out.setCreator('YourOfficeSuite'); // the new file starts with no title, author or other hidden details
  return { bytes: await out.save({ useObjectStreams: true }), imagePages: hit.size };
}

// ---------- Compare ----------
async function readWords(bytes, name, tick) {
  const view = await loadView(bytes, name); const words = []; const mctx = document.createElement('canvas').getContext('2d'); const pages = view.numPages;
  try {
    for (let n = 1; n <= pages; n++) {
      const page = await view.getPage(n); const vp = page.getViewport({ scale: 1 });
      const tc = await page.getTextContent(); const styles = tc.styles || {};
      for (const it of tc.items) {
        const str = it.str; if (!str || !str.trim() || !it.width) continue;
        const tx = Util.transform(vp.transform, it.transform);
        const fh = Math.hypot(tx[2], tx[3]) || Math.abs(it.height) || 10;
        const len = Math.hypot(tx[0], tx[1]) || 1; const ux = tx[0] / len; const uy = tx[1] / len; const nx = -uy; const ny = ux;
        const full = it.width * vp.scale;
        mctx.font = `100px ${(styles[it.fontName] && styles[it.fontName].fontFamily) || 'sans-serif'}`; const total = mctx.measureText(str).width || str.length;
        const at = (k) => (k <= 0 ? 0 : k >= str.length ? full : (mctx.measureText(str.slice(0, k)).width / total) * full);
        const re = /\S+/g; let m;
        while ((m = re.exec(str))) {
          const a = at(m.index); const b = at(m.index + m[0].length); const xs = []; const ys = [];
          for (const t of [a, b]) for (const up of [-0.25 * fh, 0.95 * fh]) { xs.push(tx[4] + ux * t - nx * up); ys.push(tx[5] + uy * t - ny * up); }
          const x0 = Math.max(0, Math.min(...xs)); const x1 = Math.min(vp.width, Math.max(...xs)); const y0 = Math.max(0, Math.min(...ys)); const y1 = Math.min(vp.height, Math.max(...ys));
          words.push({ t: m[0].normalize('NFKC'), p: n - 1, x: x0 / vp.width, y: y0 / vp.height, w: Math.max(0, x1 - x0) / vp.width, h: Math.max(0, y1 - y0) / vp.height });
        }
      }
      page.cleanup(); if (tick) tick(n / pages);
    }
    return { words, pages };
  } finally { await view.destroy(); }
}

// Myers diff on word ids, with a limit on how many edits it will search.
function myers(a, b, maxD) {
  const n = a.length; const m = b.length; const off = maxD + 2; const v = new Int32Array(2 * maxD + 5); v[off + 1] = 0; const trace = [];
  for (let d = 0; d <= maxD; d++) {
    trace.push(v.slice(off - d - 1, off + d + 2)); // keys -d-1 .. d+1
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[off + k - 1] < v[off + k + 1]) ? v[off + k + 1] : v[off + k - 1] + 1; let y = x - k;
      while (x < n && y < m && a[x] === b[y]) { x++; y++; }
      v[off + k] = x;
      if (x >= n && y >= m) {
        const ops = []; let cx = n; let cy = m;
        for (let dd = d; dd >= 0; dd--) {
          const t = trace[dd]; const get = (key) => t[key + dd + 1]; const kk = cx - cy;
          const pk = kk === -dd || (kk !== dd && get(kk - 1) < get(kk + 1)) ? kk + 1 : kk - 1;
          const px = dd === 0 ? 0 : get(pk); const py = px - pk;
          while (cx > px && cy > py) { ops.push({ k: 'eq', a: cx - 1, b: cy - 1 }); cx--; cy--; }
          if (dd > 0) { if (cx === px) ops.push({ k: 'ins', a: cx, b: py }); else ops.push({ k: 'del', a: px, b: cy }); cx = px; cy = py; }
        }
        return ops.reverse();
      }
    }
  }
  return null;
}
function diffIds(a, b, aStart, bStart, maxD) {
  // skip the shared start and end first: most documents share most of their words
  let s = 0; while (s < a.length && s < b.length && a[s] === b[s]) s++;
  let e = 0; while (e < a.length - s && e < b.length - s && a[a.length - 1 - e] === b[b.length - 1 - e]) e++;
  const mid = myers(a.subarray(s, a.length - e), b.subarray(s, b.length - e), maxD); if (!mid) return null;
  const ops = [];
  for (let i = 0; i < s; i++) ops.push({ k: 'eq', a: aStart + i, b: bStart + i });
  for (const o of mid) ops.push({ k: o.k, a: aStart + s + o.a, b: bStart + s + o.b });
  for (let i = e; i > 0; i--) ops.push({ k: 'eq', a: aStart + a.length - i, b: bStart + b.length - i });
  return ops;
}
function compareWords(oldDoc, newDoc) {
  const ids = new Map(); const id = (t) => { let v = ids.get(t); if (v == null) { v = ids.size; ids.set(t, v); } return v; };
  const A = Int32Array.from(oldDoc.words, (w) => id(w.t)); const B = Int32Array.from(newDoc.words, (w) => id(w.t));
  let exact = true; let ops = diffIds(A, B, 0, 0, 4000);
  if (!ops) {
    // too many changes for one pass: compare page by page instead
    exact = false; ops = []; const pages = Math.max(oldDoc.pages, newDoc.pages); let ai = 0; let bi = 0;
    for (let p = 0; p < pages; p++) {
      let ae = ai; while (ae < A.length && oldDoc.words[ae].p <= p) ae++;
      let be = bi; while (be < B.length && newDoc.words[be].p <= p) be++;
      const part = diffIds(A.subarray(ai, ae), B.subarray(bi, be), ai, bi, 1500);
      if (part) ops.push(...part);
      else { for (let i = ai; i < ae; i++) ops.push({ k: 'del', a: i, b: bi }); for (let j = bi; j < be; j++) ops.push({ k: 'ins', a: ae, b: j }); }
      ai = ae; bi = be;
    }
  }
  const hunks = []; let cur = null; let added = 0; let removed = 0;
  for (const o of ops) {
    if (o.k === 'eq') { if (cur) { cur.after = o.b; hunks.push(cur); cur = null; } continue; }
    if (!cur) cur = { del: [], ins: [], before: o.b - 1, after: -1 };
    if (o.k === 'del') { cur.del.push(o.a); removed++; } else { cur.ins.push(o.b); added++; }
  }
  if (cur) hunks.push(cur);
  return { ops, hunks, added, removed, exact };
}
// joins the boxes of neighbouring words on the same line, so highlights look like one band
function mergeBoxes(words, idx) {
  const out = [];
  for (const i of [...idx].sort((x, y) => x - y)) {
    const w = words[i]; if (!w || !w.w) continue; const last = out[out.length - 1];
    if (last && last.p === w.p && Math.abs(last.y - w.y) < w.h * 0.5 && w.x >= last.x && w.x - (last.x + last.w) < w.h * 1.6) {
      const x1 = Math.max(last.x + last.w, w.x + w.w); const y0 = Math.min(last.y, w.y); const y1 = Math.max(last.y + last.h, w.y + w.h);
      last.w = x1 - last.x; last.y = y0; last.h = y1 - y0;
    } else out.push({ p: w.p, x: w.x, y: w.y, w: w.w, h: w.h });
  }
  return out;
}
// Marked-up PDF: each original page (removed words in red) followed by its new page (added words in green).
async function opMarkup(oldF, newF, oldDoc, newDoc, diff, tick) {
  const A = await loadLib(oldF.bytes, oldF.name); const B = await loadLib(newF.bytes, newF.name);
  const out = await PDFDocument.create(); const font = await out.embedFont(StandardFonts.Helvetica);
  const oldP = await out.copyPages(A, A.getPageIndices()); const newP = await out.copyPages(B, B.getPageIndices());
  const del = mergeBoxes(oldDoc.words, diff.hunks.flatMap((h) => h.del)); const ins = mergeBoxes(newDoc.words, diff.hunks.flatMap((h) => h.ins));
  const mark = (page, boxes, p, color, label) => {
    const v = visualSize(page);
    for (const b of boxes) if (b.p === p) {
      const vw = b.w * v.w; const vh = b.h * v.h; const q = place(page, b.x * v.w, b.y * v.h, vw, vh);
      page.drawRectangle({ x: q.x, y: q.y, width: vw, height: vh, rotate: q.rotate, color: rgb(...color), opacity: 0.38, blendMode: BlendMode.Multiply });
    }
    const size = 8; const tw = font.widthOfTextAtSize(label, size) + 10; const q = place(page, v.w - tw - 6, 6, tw, 14);
    page.drawRectangle({ x: q.x, y: q.y, width: tw, height: 14, rotate: q.rotate, color: rgb(1, 1, 1), opacity: 0.92, borderColor: rgb(...color), borderWidth: 0.8 });
    const t = place(page, v.w - tw - 1, 6, 0, 10);
    page.drawText(label, { x: t.x, y: t.y, size, font, color: rgb(0.15, 0.15, 0.2), rotate: t.rotate });
  };
  const n = Math.max(oldP.length, newP.length);
  for (let i = 0; i < n; i++) {
    if (oldP[i]) mark(out.addPage(oldP[i]), del, i, [0.95, 0.25, 0.3], `Original, page ${i + 1}`);
    if (newP[i]) mark(out.addPage(newP[i]), ins, i, [0.15, 0.78, 0.35], `New, page ${i + 1}`);
    if (tick) tick((i + 1) / n);
  }
  out.setTitle('PDF comparison'); out.setProducer('YourOfficeSuite'); out.setCreator('YourOfficeSuite');
  return out.save({ useObjectStreams: true });
}

// ---------- PDF to Text ----------
async function opToText(bytes, name, tick) {
  const view = await loadView(bytes, name); const parts = []; let words = 0; const pages = view.numPages;
  try {
    for (let n = 1; n <= pages; n++) {
      const page = await view.getPage(n); const tc = await page.getTextContent(); let s = ''; let lastY = null;
      for (const it of tc.items) {
        if (typeof it.str !== 'string') continue;
        const y = it.transform ? it.transform[5] : null;
        if (lastY != null && y != null && Math.abs(y - lastY) > 2 && !s.endsWith('\n')) s += '\n';
        s += it.str; if (it.hasEOL) s += '\n'; if (y != null) lastY = y;
      }
      s = s.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
      words += (s.match(/\S+/g) || []).length;
      parts.push(pages > 1 ? `--- Page ${n} ---\n${s}` : s);
      page.cleanup(); if (tick) tick(n / pages);
    }
  } finally { await view.destroy(); }
  if (!words) throw new Error('No text found. This PDF may be a scan, which is an image of text.');
  return { bytes: new TextEncoder().encode('﻿' + parts.join('\n\n') + '\n'), words, pages: parts.length };
}
// Saves every image inside the PDF at full size: JPG for photos, PNG when it has transparency. Skips duplicates and tiny images.
async function opExtractImages(bytes, name, tick) {
  const view = await loadView(bytes, name); const out = []; const seen = new Set(); const pages = view.numPages;
  const kinds = new Set([OPS.paintImageXObject, OPS.paintInlineImageXObject, OPS.paintImageXObjectRepeat]);
  try {
    for (let n = 1; n <= pages; n++) {
      const page = await view.getPage(n); const list = await page.getOperatorList(); let k = 0;
      for (let i = 0; i < list.fnArray.length; i++) {
        if (!kinds.has(list.fnArray[i])) continue;
        const arg = list.argsArray[i][0]; let img = null;
        if (typeof arg === 'string') {
          if (seen.has(arg)) continue; seen.add(arg);
          const store = arg.startsWith('g_') ? page.commonObjs : page.objs;
          img = await new Promise((res) => { try { store.get(arg, res); } catch (_) { res(null); } });
        } else img = arg;
        if (!img || !img.width || !img.height || img.width < 16 || img.height < 16) continue;
        const c = document.createElement('canvas'); c.width = img.width; c.height = img.height; const g = c.getContext('2d'); let alpha = false;
        if (img.bitmap) {
          g.drawImage(img.bitmap, 0, 0);
          const px = g.getImageData(0, 0, img.width, img.height).data; for (let q = 3; q < px.length; q += 4) if (px[q] < 255) { alpha = true; break; }
        } else if (img.data) {
          const idt = g.createImageData(img.width, img.height); const d = idt.data; const src = img.data;
          if (img.kind === ImageKind.RGBA_32BPP) { d.set(src.subarray(0, d.length)); alpha = true; }
          else if (img.kind === ImageKind.RGB_24BPP) { for (let p = 0, q = 0; q < d.length; p += 3, q += 4) { d[q] = src[p]; d[q + 1] = src[p + 1]; d[q + 2] = src[p + 2]; d[q + 3] = 255; } }
          else if (img.kind === ImageKind.GRAYSCALE_1BPP) {
            const rowBytes = (img.width + 7) >> 3;
            for (let y = 0; y < img.height; y++) for (let x = 0; x < img.width; x++) { const bit = (src[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1; const v = bit ? 255 : 0; const q = (y * img.width + x) * 4; d[q] = d[q + 1] = d[q + 2] = v; d[q + 3] = 255; }
          } else continue;
          g.putImageData(idt, 0, 0);
        } else continue;
        k++;
        const png = alpha && img.width * img.height < 4e6;
        const data = await blobBytes(await canvasBlob(c, png ? 'image/png' : 'image/jpeg', 0.92)); c.width = c.height = 1;
        // the same picture used on several pages is saved once
        let hsh = 2166136261; for (let q = 0; q < data.length; q += 7) hsh = Math.imul(hsh ^ data[q], 16777619);
        const key = `${img.width}x${img.height}:${data.length}:${hsh}`; if (seen.has(key)) { k--; continue; } seen.add(key);
        out.push({ name: `${stemOf(name)}-page${pad(n, 3)}-${k}.${png ? 'png' : 'jpg'}`, bytes: data });
      }
      page.cleanup(); if (tick) tick(n / pages);
    }
  } finally { await view.destroy(); }
  if (!out.length) throw new Error('No images found in this PDF.');
  return out;
}

// ---------- Crop and page size ----------
// Finds the white margins of a page, in points, from a small render.
async function autoMargins(view, pageIndex) {
  const r = await renderPage(view, pageIndex + 1, { scale: 1, maxDim: 900, forms: true });
  const w = r.canvas.width; const h = r.canvas.height; const d = r.canvas.getContext('2d').getImageData(0, 0, w, h).data;
  let x0 = w; let y0 = h; let x1 = -1; let y1 = -1;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const q = (y * w + x) * 4; if (d[q] < 245 || d[q + 1] < 245 || d[q + 2] < 245) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; } }
  r.canvas.width = r.canvas.height = 1;
  if (x1 < 0) return { top: 0, right: 0, bottom: 0, left: 0 };
  const k = r.ptW / w; const gap = 8; // keep a little white space
  return { top: Math.max(0, y0 * k - gap), right: Math.max(0, (w - 1 - x1) * k - gap), bottom: Math.max(0, (h - 1 - y1) * k - gap), left: Math.max(0, x0 * k - gap) };
}
// margins are in points on the page as you see it; 'auto' gives every page its own white margins
async function opCrop(bytes, name, margins, pages, tick) {
  const doc = await loadLib(bytes, name); const all = doc.getPages(); const view = margins === 'auto' ? await loadView(bytes, name) : null;
  const list = pages === 'all' ? all.map((_, i) => i) : pages; let done = 0;
  try {
    for (const i of list) {
      const page = all[i]; if (!page) continue; const v = visualSize(page);
      const m = margins === 'auto' ? await autoMargins(view, i) : margins;
      if (tick) tick(++done / list.length);
      if (m.left + m.right >= v.w - 4 || m.top + m.bottom >= v.h - 4) {
        if (margins === 'auto') continue; // blank page: leave it as it is
        throw new Error(`The margins are bigger than page ${i + 1}. Use smaller numbers.`);
      }
      const a = place(page, m.left, m.top, 0, 0); const b = place(page, v.w - m.right, v.h - m.bottom, 0, 0); // visual corners in PDF space
      const x0 = Math.min(a.x, b.x); const y0 = Math.min(a.y, b.y); const x1 = Math.max(a.x, b.x); const y1 = Math.max(a.y, b.y);
      page.setCropBox(x0, y0, x1 - x0, y1 - y0);
      try { page.setTrimBox(x0, y0, x1 - x0, y1 - y0); page.setBleedBox(x0, y0, x1 - x0, y1 - y0); } catch (_) { /* optional boxes */ }
    }
  } finally { if (view) await view.destroy(); }
  return doc.save({ useObjectStreams: true });
}
const PAPER = { a4: [595.28, 841.89], letter: [612, 792], legal: [612, 1008], a5: [419.53, 595.28] };
// Each page is scaled to fit the paper size and centred. Wide pages stay wide. Rotated pages come out upright.
async function opResize(bytes, name, size, tick) {
  const src = await loadLib(bytes, name);
  try { const f = src.getForm(); if (f.getFields().length) f.flatten(); } catch (_) { /* no form */ }
  const out = await PDFDocument.create(); const pages = src.getPages(); let i = 0;
  for (const p of pages) {
    const crop = p.getCropBox(); const emb = await out.embedPage(p, { left: crop.x, bottom: crop.y, right: crop.x + crop.width, top: crop.y + crop.height });
    const rot = ((p.getRotation().angle % 360) + 360) % 360; const v = visualSize(p);
    let [W, H] = PAPER[size]; if (v.w > v.h) [W, H] = [H, W];
    const k = Math.min(W / v.w, H / v.h); const ew = crop.width * k; const eh = crop.height * k;
    const th = (-rot * Math.PI) / 180; const cs = Math.cos(th); const sn = Math.sin(th);
    const pts = [[0, 0], [ew, 0], [0, eh], [ew, eh]].map(([x, y]) => [x * cs - y * sn, x * sn + y * cs]);
    const minX = Math.min(...pts.map((q) => q[0])); const minY = Math.min(...pts.map((q) => q[1]));
    const ox = (W - v.w * k) / 2; const oy = (H - v.h * k) / 2;
    out.addPage([W, H]).drawPage(emb, { x: ox - minX, y: oy - minY, xScale: k, yScale: k, rotate: degrees(-rot) });
    if (tick) tick(++i / pages.length);
  }
  return out.save({ useObjectStreams: true });
}

// ---------- Flatten ----------
const KEEP_ANNOT = new Set(['Link', 'Popup', 'Widget']);
// Forms and comments only: fields are flattened, each comment or stamp is drawn into the page and then removed.
async function opFlatten(bytes, name) {
  const doc = await loadLib(bytes, name); let fields = 0; let notes = 0;
  try { const form = doc.getForm(); fields = form.getFields().length; if (fields) form.flatten(); } catch (_) { fields = 0; }
  for (const page of doc.getPages()) {
    const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray); if (!annots) continue;
    const keep = [];
    for (let i = 0; i < annots.size(); i++) {
      const ref = annots.get(i); const a = annots.lookupMaybe(i, PDFDict); if (!a) continue;
      const subN = a.lookupMaybe(PDFName.of('Subtype'), PDFName); const sub = subN ? subN.decodeText() : '';
      if (sub === 'Popup') continue; // popups go with their comment
      if (KEEP_ANNOT.has(sub)) { keep.push(ref); continue; }
      const fl = a.lookupMaybe(PDFName.of('F'), PDFNumber); const flags = fl ? fl.asNumber() : 0;
      const ap = a.lookupMaybe(PDFName.of('AP'), PDFDict); let n = ap ? ap.get(PDFName.of('N')) : null;
      if (n && !(n instanceof PDFRef) && ap) { const st = ap.lookup(PDFName.of('N')); if (st instanceof PDFDict) { const as = a.lookupMaybe(PDFName.of('AS'), PDFName); n = as ? st.get(as) : null; } }
      if ((flags & 2) || !(n instanceof PDFRef)) { notes++; continue; } // hidden or nothing to draw: drop it
      const xo = doc.context.lookup(n);
      const rArr = a.lookup(PDFName.of('Rect'), PDFArray); const r = [0, 1, 2, 3].map((j) => rArr.lookup(j, PDFNumber).asNumber());
      const bb = xo.dict.lookup(PDFName.of('BBox'), PDFArray); const b = [0, 1, 2, 3].map((j) => bb.lookup(j, PDFNumber).asNumber());
      const mArr = xo.dict.lookupMaybe(PDFName.of('Matrix'), PDFArray); const M = mArr ? [0, 1, 2, 3, 4, 5].map((j) => mArr.lookup(j, PDFNumber).asNumber()) : [1, 0, 0, 1, 0, 0];
      const corners = [[b[0], b[1]], [b[2], b[1]], [b[0], b[3]], [b[2], b[3]]].map(([x, y]) => [M[0] * x + M[2] * y + M[4], M[1] * x + M[3] * y + M[5]]);
      const tx0 = Math.min(...corners.map((c) => c[0])); const ty0 = Math.min(...corners.map((c) => c[1])); const tx1 = Math.max(...corners.map((c) => c[0])); const ty1 = Math.max(...corners.map((c) => c[1]));
      const rx0 = Math.min(r[0], r[2]); const ry0 = Math.min(r[1], r[3]); const rx1 = Math.max(r[0], r[2]); const ry1 = Math.max(r[1], r[3]);
      const sx = tx1 > tx0 ? (rx1 - rx0) / (tx1 - tx0) : 1; const sy = ty1 > ty0 ? (ry1 - ry0) / (ty1 - ty0) : 1;
      const tag = page.node.newXObject('FlatAnnot', n);
      page.pushOperators(pushGraphicsState(), concatTransformationMatrix(sx, 0, 0, sy, rx0 - tx0 * sx, ry0 - ty0 * sy), drawObject(tag), popGraphicsState());
      notes++;
    }
    if (keep.length) page.node.set(PDFName.of('Annots'), doc.context.obj(keep)); else page.node.delete(PDFName.of('Annots'));
  }
  if (!fields && !notes) throw new Error('This PDF has no form fields or comments to flatten.');
  return { bytes: await doc.save({ useObjectStreams: true }), fields, notes };
}
// Whole PDF as images: every page becomes a PNG, so nothing can be selected or edited.
async function opFlattenImages(bytes, name, dpi, tick) {
  const view = await loadView(bytes, name); const out = await PDFDocument.create(); const pages = view.numPages;
  try {
    for (let n = 1; n <= pages; n++) {
      const r = await renderPage(view, n, { scale: dpi / 72, maxDim: 3400, forms: true });
      const img = await out.embedPng(await blobBytes(await canvasBlob(r.canvas, 'image/png')));
      out.addPage([r.ptW, r.ptH]).drawImage(img, { x: 0, y: 0, width: r.ptW, height: r.ptH });
      r.canvas.width = r.canvas.height = 1; if (tick) tick(n / pages);
    }
  } finally { await view.destroy(); }
  out.setProducer('YourOfficeSuite'); out.setCreator('YourOfficeSuite');
  return out.save({ useObjectStreams: true });
}

// ---------- Black & white ----------
async function opGrayscale(bytes, name, dpi, tick) {
  const view = await loadView(bytes, name); const out = await PDFDocument.create(); const pages = view.numPages;
  try {
    for (let n = 1; n <= pages; n++) {
      const r = await renderPage(view, n, { scale: dpi / 72, maxDim: 3400, forms: true });
      const g = r.canvas.getContext('2d'); const idt = g.getImageData(0, 0, r.canvas.width, r.canvas.height); const d = idt.data;
      for (let q = 0; q < d.length; q += 4) { const y = Math.round(0.299 * d[q] + 0.587 * d[q + 1] + 0.114 * d[q + 2]); d[q] = d[q + 1] = d[q + 2] = y; }
      g.putImageData(idt, 0, 0);
      const img = await out.embedJpg(await blobBytes(await canvasBlob(r.canvas, 'image/jpeg', 0.85)));
      out.addPage([r.ptW, r.ptH]).drawImage(img, { x: 0, y: 0, width: r.ptW, height: r.ptH });
      r.canvas.width = r.canvas.height = 1; if (tick) tick(n / pages);
    }
  } finally { await view.destroy(); }
  return out.save({ useObjectStreams: true });
}

// ---------- Edit and annotate ----------
const BASELINE = 0.88; const LINE = 1.2; // text sits this far below the top of its box, in font sizes
const MARKS = { tick: 'M4 12.5 9.5 18 20 6', cross: 'M6 6 18 18 M18 6 6 18' }; // drawn as lines in a 24 x 24 box, so they show in every PDF reader
const WIN_ANSI = /^[\x20-\x7E\xA0-\xFF€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ\t\r\n]*$/;
const latinOnly = (t) => WIN_ANSI.test(t);
const hexRgb = (hex) => { const n = parseInt(String(hex || '#111111').replace('#', '').padEnd(6, '0').slice(0, 6), 16); return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); };
const prettyName = (name) => { const last = name.split('.').pop() || name; return last.replace(/\[\d+\]/g, '').replace(/[_]+/g, ' ').replace(/\s+/g, ' ').trim() || name; };

async function formInfo(bytes, name) {
  const doc = await loadLib(bytes, name); let form;
  try { form = doc.getForm(); } catch (_) { return []; }
  const fields = form.getFields(); if (!fields.length) return [];
  const pages = doc.getPages(); const where = new Map(); // which page each widget sits on
  pages.forEach((p, i) => { const annots = p.node.Annots(); if (annots) for (let k = 0; k < annots.size(); k++) where.set(annots.lookup(k), i); });
  const out = [];
  for (const f of fields) {
    let type; let value; let options; let multiline = false; let maxLength;
    try {
      if (f instanceof PDFTextField) { type = 'text'; value = f.getText() || ''; multiline = f.isMultiline(); maxLength = f.getMaxLength(); }
      else if (f instanceof PDFCheckBox) { type = 'check'; value = f.isChecked(); }
      else if (f instanceof PDFDropdown || f instanceof PDFOptionList) { type = 'choice'; options = f.getOptions(); value = f.getSelected()[0] || ''; }
      else if (f instanceof PDFRadioGroup) { type = 'radio'; options = f.getOptions(); value = f.getSelected() || ''; }
      else continue; // buttons and signature fields are left alone
    } catch (_) { continue; }
    let label = prettyName(f.getName());
    try { const tu = f.acroField.dict.lookup(PDFName.of('TU')); if (tu instanceof PDFString || tu instanceof PDFHexString) { const t = tu.decodeText().trim(); if (t && t.length <= 80) label = t; } } catch (_) { /* keep the short name */ }
    const rects = [];
    try {
      for (const w of f.acroField.getWidgets()) {
        let pi = where.get(w.dict);
        if (pi == null) { const pref = w.P(); if (pref) pi = pages.findIndex((p) => p.ref === pref); }
        if (pi == null || pi < 0) continue;
        const page = pages[pi];
        if (((page.getRotation().angle % 360) + 360) % 360 !== 0) continue; // outlines only on upright pages
        const c = page.getCropBox(); const r = w.getRectangle();
        rects.push({ page: pi, x: (r.x - c.x) / c.width, y: (c.y + c.height - (r.y + r.height)) / c.height, w: r.width / c.width, h: r.height / c.height });
      }
    } catch (_) { /* no outline for this field */ }
    let readOnly = false; try { readOnly = !!f.isReadOnly(); } catch (_) { readOnly = false; }
    out.push({ key: f.getName(), label, type, value, options, multiline, maxLength, readOnly, rects });
  }
  return out;
}
// A line the standard PDF font cannot show (for example Hindi or Chinese) is drawn as an image, so it still appears.
async function textImage(doc, text, size, color) {
  const k = 4; const px = Math.max(8, Math.round(size * k)); const font = `${px}px Helvetica, Arial, "Noto Sans", sans-serif`;
  const c = document.createElement('canvas'); let g = c.getContext('2d'); g.font = font;
  const w = Math.ceil(g.measureText(text).width) + 4; const h = Math.ceil(px * LINE);
  c.width = Math.max(2, w); c.height = h; g = c.getContext('2d'); g.font = font; g.fillStyle = color; g.textBaseline = 'alphabetic';
  g.fillText(text, 2, px * BASELINE);
  return { img: await doc.embedPng(await blobBytes(await canvasBlob(c, 'image/png'))), w: c.width / k, h: c.height / k };
}
// items: text, highlights, ticks, crosses and signatures (x, y, w, h are fractions of the page as seen; size in points). fields: form values.
async function opEdit(bytes, { items, fields, lock }, name) {
  const doc = await loadLib(bytes, name); const helv = await doc.embedFont(StandardFonts.Helvetica);
  const keys = Object.keys(fields);
  if (keys.length || lock) {
    const form = doc.getForm(); const bad = [];
    for (const f of form.getFields()) {
      const key = f.getName(); if (!(key in fields)) continue; const v = fields[key];
      try {
        if (f instanceof PDFTextField) {
          const t = String(v == null ? '' : v); if (t && !latinOnly(t)) { bad.push(prettyName(key)); continue; }
          const max = f.getMaxLength(); f.setText(max ? t.slice(0, max) : t);
          // multi-line boxes set to "auto size" come out huge, so give them a normal size
          if (f.isMultiline() && /(^|\s)0(\.0+)?\s+Tf/.test(f.acroField.getDefaultAppearance() || ' 0 Tf')) f.setFontSize(10);
        } else if (f instanceof PDFCheckBox) { if (v) f.check(); else f.uncheck(); }
        else if (f instanceof PDFDropdown || f instanceof PDFOptionList) { if (v) f.select(String(v)); else f.clear(); }
        else if (f instanceof PDFRadioGroup) { if (v) f.select(String(v)); else f.clear(); }
      } catch (_) { /* a field that cannot be changed stays as it was */ }
    }
    if (bad.length) throw new Error(`Form fields can hold Latin letters, numbers and common symbols. Check: ${bad.slice(0, 3).join(', ')}${bad.length > 3 ? '…' : ''}. For other scripts, use Text on the page.`);
    if (lock) { try { form.updateFieldAppearances(helv); } catch (_) { /* flatten will try again */ } form.flatten(); }
  }
  const pages = doc.getPages(); const sigCache = new Map();
  for (const it of items) {
    const page = pages[it.page]; if (!page) continue;
    const v = visualSize(page); const vx = it.x * v.w; const vy = it.y * v.h;
    if (it.kind === 'hl') {
      const vw = it.w * v.w; const vh = it.h * v.h; const p = place(page, vx, vy, vw, vh);
      page.drawRectangle({ x: p.x, y: p.y, width: vw, height: vh, rotate: p.rotate, color: rgb(1, 0.86, 0.2), opacity: 0.45, blendMode: BlendMode.Multiply });
    } else if (it.kind === 'text') {
      const size = it.size || 12; const color = it.color || '#111111'; const lines = String(it.text || '').split('\n');
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i]; if (!line.trim()) continue; const top = vy + i * size * LINE;
        if (latinOnly(line)) { const p = place(page, vx, top, 0, size * BASELINE); page.drawText(line, { x: p.x, y: p.y, size, font: helv, color: hexRgb(color), rotate: p.rotate }); }
        else { const t = await textImage(doc, line, size, color); const p = place(page, vx, top, t.w, t.h); page.drawImage(t.img, { x: p.x, y: p.y, width: t.w, height: t.h, rotate: p.rotate }); }
      }
    } else if (it.kind === 'tick' || it.kind === 'cross') {
      const size = it.size || 16; const p = place(page, vx, vy, 0, 0);
      page.drawSvgPath(MARKS[it.kind], { x: p.x, y: p.y, scale: size / 24, rotate: p.rotate, borderColor: hexRgb(it.color || '#111111'), borderWidth: 3, borderLineCap: LineCapStyle.Round });
    } else if (it.kind === 'sig' && it.png) {
      let img = sigCache.get(it.png); if (!img) { img = await doc.embedPng(it.png); sigCache.set(it.png, img); }
      const vw = it.w * v.w; const vh = it.h * v.h; const p = place(page, vx, vy, vw, vh);
      page.drawImage(img, { x: p.x, y: p.y, width: vw, height: vh, rotate: p.rotate });
    }
  }
  return doc.save({ useObjectStreams: true });
}

// ================================================================ shared UI pieces
const note = (text) => el('div', 'hint', text);

// One page of a PDF with previous / next and zoom. Things placed in `layer` use percentages, so they stay put when zooming.
function pageStage(view, { forms, onShow } = {}) {
  const wrap = el('div', 'stage-wrap'); const stage = el('div', 'pstage'); const layer = el('div', 'player'); wrap.append(stage);
  const pager = el('div', 'pager'); const prev = el('button', 'btn icon ghost'); prev.append(icon('i-back')); prev.title = 'Previous page'; prev.dataset.act = 'prev';
  const next = el('button', 'btn icon ghost'); const nx = icon('i-back'); nx.style.transform = 'rotate(180deg)'; next.append(nx); next.title = 'Next page'; next.dataset.act = 'next';
  const lab = el('span'); const zo = el('button', 'btn ghost zoomb', '−'); zo.title = 'Zoom out'; const zi = el('button', 'btn ghost zoomb', '+'); zi.title = 'Zoom in'; const zl = el('span', 'zooml');
  pager.append(prev, lab, next, el('span', 'sep'), zo, zl, zi);
  const st = { wrap, pager, layer, page: 1, pages: view.numPages, ptW: 1, ptH: 1, zoom: 1, canvas: null };
  let busy = false;
  const size = () => {
    if (!st.canvas) return; const w = Math.round(Math.min(st.ptW * 1.2, 760) * st.zoom);
    st.canvas.style.width = `${w}px`; st.canvas.style.height = `${Math.round(w * st.ptH / st.ptW)}px`; layer.style.setProperty('--pt', `${w / st.ptW}px`); zl.textContent = `${Math.round(st.zoom * 100)}%`;
  };
  st.show = async () => {
    if (busy) return; busy = true;
    try {
      const r = await renderPage(view, st.page, { scale: 2, maxDim: 1900, forms });
      st.ptW = r.ptW; st.ptH = r.ptH; st.canvas = r.canvas; stage.textContent = ''; stage.append(r.canvas, layer); size();
      lab.textContent = `Page ${st.page} of ${st.pages}`; prev.disabled = st.page === 1; next.disabled = st.page === st.pages; stage.dataset.page = String(st.page);
      if (onShow) await onShow(st);
    } finally { busy = false; }
  };
  st.go = async (n) => { st.page = clamp(n, 1, st.pages); await st.show(); };
  prev.addEventListener('click', () => { if (!busy && st.page > 1) st.go(st.page - 1); }); next.addEventListener('click', () => { if (!busy && st.page < st.pages) st.go(st.page + 1); });
  zo.addEventListener('click', () => { st.zoom = clamp(Math.round((st.zoom - 0.2) * 10) / 10, 0.6, 2); size(); }); zi.addEventListener('click', () => { st.zoom = clamp(Math.round((st.zoom + 0.2) * 10) / 10, 0.6, 2); size(); });
  return st;
}
// A box on the page that can be dragged, resized from its corner and removed. b holds fractions of the page.
function fracBox(layer, b, cls, { onRemove, onChange, resizable = true, ratio } = {}) {
  const d = el('div', `pbox ${cls}`);
  const put = () => { d.style.left = `${b.x * 100}%`; d.style.top = `${b.y * 100}%`; if (b.w != null) { d.style.width = `${b.w * 100}%`; d.style.height = `${b.h * 100}%`; } }; put();
  const x = el('button', 'bx', '×'); x.title = 'Remove'; x.addEventListener('pointerdown', (e) => e.stopPropagation());
  x.addEventListener('click', (e) => { e.stopPropagation(); d.remove(); if (onRemove) onRemove(); }); d.append(x);
  let h = null; if (resizable) { h = el('span', 'h'); d.append(h); }
  let act = null;
  d.addEventListener('pointerdown', (e) => {
    if (e.target.closest('[contenteditable],input,select,button')) return;
    e.preventDefault(); e.stopPropagation(); const r = layer.getBoundingClientRect();
    act = { kind: e.target === h ? 'size' : 'move', sx: e.clientX, sy: e.clientY, b: { ...b }, W: r.width, H: r.height }; d.setPointerCapture(e.pointerId);
  });
  d.addEventListener('pointermove', (e) => {
    if (!act) return; const dx = (e.clientX - act.sx) / act.W; const dy = (e.clientY - act.sy) / act.H;
    if (act.kind === 'move') { b.x = clamp(act.b.x + dx, 0, 1 - (b.w || 0.01)); b.y = clamp(act.b.y + dy, 0, 1 - (b.h || 0.01)); }
    else { b.w = clamp(act.b.w + dx, 0.012, 1 - b.x); b.h = ratio ? b.w * ratio * act.W / act.H : clamp(act.b.h + dy, 0.006, 1 - b.y); }
    put(); if (onChange) onChange();
  });
  const end = () => { act = null; }; d.addEventListener('pointerup', end); d.addEventListener('pointercancel', end);
  d.put = put; layer.append(d); return d;
}
// Lets the user drag out a new box on an empty part of the page.
function rubber(layer, cls, isOn, done) {
  let s = null; let ghost = null;
  const pos = (e) => { const r = layer.getBoundingClientRect(); return { x: clamp((e.clientX - r.left) / r.width, 0, 1), y: clamp((e.clientY - r.top) / r.height, 0, 1) }; };
  layer.addEventListener('pointerdown', (e) => {
    if (e.target !== layer || !isOn() || e.button !== 0) return; e.preventDefault();
    s = pos(e); ghost = el('div', `pbox ${cls} ghost`); layer.append(ghost); layer.setPointerCapture(e.pointerId);
  });
  layer.addEventListener('pointermove', (e) => {
    if (!s) return; const p = pos(e);
    ghost.style.left = `${Math.min(s.x, p.x) * 100}%`; ghost.style.top = `${Math.min(s.y, p.y) * 100}%`; ghost.style.width = `${Math.abs(p.x - s.x) * 100}%`; ghost.style.height = `${Math.abs(p.y - s.y) * 100}%`;
  });
  const up = (e) => {
    if (!s) return; const p = pos(e); const b = { x: Math.min(s.x, p.x), y: Math.min(s.y, p.y), w: Math.abs(p.x - s.x), h: Math.abs(p.y - s.y) };
    ghost.remove(); ghost = null; s = null; if (b.w > 0.008 && b.h > 0.005) done(b);
  };
  layer.addEventListener('pointerup', up); layer.addEventListener('pointercancel', () => { if (ghost) ghost.remove(); ghost = null; s = null; });
}
// One PDF in. build(ctx) adds the preview, the options and the run button.
function onePdf(tool, build) {
  const { main, side } = workspace(tool);
  const start = () => { main.textContent = ''; side.hidden = true; side.textContent = ''; main.append(dropZone({ kind: 'pdf', title: tool.pick || 'Choose a PDF', onPaths: (p) => load(p[0]) })); };
  let open = null;
  async function load(path) {
    const f = await guard(() => readFile(path)); if (!f) return;
    const view = await guard(() => loadView(f.bytes, f.name)); if (!view) return;
    if (open) { try { await open.destroy(); } catch (_) { /* already closed */ } } open = view;
    main.textContent = ''; side.textContent = ''; side.hidden = false;
    const bar = el('div', 'bar2'); const other = btn('btn ghost', 'Choose another file'); other.addEventListener('click', start);
    bar.append(el('b', '', f.name), el('span', 'chip', `${plural(view.numPages, 'page')} · ${fmtSize(f.size)}`), el('div', 'grow'), other); main.append(bar);
    await build({ f, view, main, side, bar, other });
  }
  start();
  const init = takeInitial(); if (init) load(init[0]);
}
const preview = async (ctx, opts) => { const st = pageStage(ctx.view, opts); ctx.main.append(st.wrap, st.pager); await st.show(); return st; };

// ================================================================ tools

function mountRedact(tool) {
  onePdf(tool, async ({ f, view, main, side }) => {
    let boxes = []; const count = el('div', 'hint'); count.id = 'redCount';
    const color = select([['black', 'Black'], ['white', 'White']], 'black'); color.id = 'redColor';
    const refresh = () => {
      const pages = new Set(boxes.map((b) => b.page)).size;
      count.textContent = boxes.length ? `${plural(boxes.length, 'box', 'boxes')} on ${plural(pages, 'page')}.` : 'No boxes yet. Drag on the page to draw one.'; count.dataset.n = String(boxes.length);
    };
    const addEl = (b) => fracBox(st.layer, b, `rbox${color.value === 'white' ? ' white' : ''}`, { onRemove: () => { boxes = boxes.filter((x) => x !== b); refresh(); } });
    const paint = () => { st.layer.textContent = ''; for (const b of boxes) if (b.page === st.page - 1) addEl(b); refresh(); };
    const st = pageStage(view, { forms: true, onShow: paint }); st.layer.classList.add('draw'); main.append(st.wrap, st.pager);
    rubber(st.layer, 'rbox', () => true, (r) => { const b = { id: rid(), page: st.page - 1, ...r }; boxes.push(b); addEl(b); refresh(); });
    color.addEventListener('change', paint);

    const words = el('input', 'field'); words.id = 'redWords'; words.placeholder = 'Names or numbers, with commas between';
    const chk = (id, text) => { const l = el('label', 'chk'); const c = el('input'); c.type = 'checkbox'; c.id = id; l.append(c, document.createTextNode(text)); return [l, c]; };
    const [le, ce] = chk('redEmails', 'All email addresses'); const [lp, cp] = chk('redPhones', 'All phone numbers');
    const find = btn('btn', 'Find and black out', 'i-search'); find.dataset.act = 'find'; const msg = el('div', 'hint'); msg.id = 'redFound';
    find.addEventListener('click', async () => {
      const opts = { words: words.value.split(','), emails: ce.checked, phones: cp.checked };
      if (!opts.words.some((w) => w.trim()) && !opts.emails && !opts.phones) { msg.textContent = 'Type a word or tick a box first.'; return; }
      find.disabled = true; msg.textContent = 'Searching…';
      const r = await guard(() => findText(f.bytes, f.name, opts)); find.disabled = false; if (!r) { msg.textContent = ''; return; }
      boxes = boxes.filter((b) => !b.auto).concat(r.boxes);
      const pages = new Set(r.boxes.map((b) => b.page));
      msg.textContent = !r.textPages ? 'This PDF looks like a scan, so search cannot read it. Draw the boxes by hand.'
        : r.boxes.length ? `Found ${plural(r.boxes.length, 'match', 'matches')} on ${plural(pages.size, 'page')}. Check them before saving.` : 'No matches found.';
      msg.dataset.n = String(r.boxes.length);
      if (r.boxes.length && !pages.has(st.page - 1)) await st.go(Math.min(...pages) + 1); else paint();
    });
    const clear = btn('btn ghost', 'Remove all'); clear.dataset.act = 'clear'; clear.addEventListener('click', () => { boxes = []; paint(); });
    side.append(el('h3', '', 'Find and black out'), words, le, lp, find, msg, el('h3', '', 'Boxes'), note('Drag on the page to draw a box. Drag a box to move it, pull its corner to resize it.'), field('Box colour', color), count, clear);
    runner(side, 'Save redacted PDF', async (r) => {
      if (!boxes.length) throw new Error('Draw a box over what you want to hide first.');
      const out = await opRedact(f.bytes, f.name, { boxes, color: color.value }, (v) => r.tick(v * 0.95));
      const saved = await saveOne(out.bytes, f.dir, `${stemOf(f.name)}-redacted.pdf`);
      if (saved) r.done(`Saved ${baseName(saved)}. What was under the boxes is removed for good.`, [saved]);
    });
    side.append(note('Pages with a box are saved as images, so their text can no longer be selected. Other pages stay as they are. Your original file is not changed.'));
    await st.show();
  });
}

function mountCompare(tool) {
  const { main, side } = workspace(tool); side.hidden = false;
  const F = { a: null, b: null }; let views = { a: null, b: null }; let last = null;
  const slots = el('div', 'cmp-slots'); const result = el('div'); result.id = 'cmpResult'; let run; let save = null; let saveMsg = null;
  const slotEls = {};
  const closeViews = async () => { for (const k of ['a', 'b']) if (views[k]) { try { await views[k].destroy(); } catch (_) { /* closed */ } views[k] = null; } };
  const sync = () => {
    for (const k of ['a', 'b']) { const s = slotEls[k]; s.name.textContent = F[k] ? F[k].name : 'No file chosen'; s.box.classList.toggle('has', !!F[k]); }
    if (run) run.enable(!!(F.a && F.b));
  };
  const reset = () => { last = null; result.textContent = ''; if (save) { save.hidden = true; saveMsg.textContent = ''; } if (run) run.clear(); };
  const setFile = async (k, path) => { const f = await guard(() => readFile(path)); if (!f) return; F[k] = f; await closeViews(); reset(); sync(); };
  const slot = (k, label) => {
    const box = el('div', 'cmp-slot'); box.dataset.slot = k; const name = el('span', 'nm', 'No file chosen'); const b = btn('btn', 'Choose a PDF'); b.dataset.act = `choose-${k}`;
    b.addEventListener('click', async () => { const p = await api.invoke('dlg:open', { kind: 'pdf', multi: false, title: label }); if (p.length) setFile(k, p[0]); });
    box.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); box.classList.add('over'); }); box.addEventListener('dragleave', () => box.classList.remove('over'));
    box.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); box.classList.remove('over'); const p = [...e.dataTransfer.files].map((x) => api.pathForFile(x)).filter((x) => /\.pdf$/i.test(x)); if (p.length) setFile(k, p[0]); else toast('Choose a PDF file', true); });
    box.append(el('b', '', label), name, b); slotEls[k] = { box, name }; return box;
  };
  const swap = el('button', 'btn icon ghost', '⇄'); swap.title = 'Swap the two files'; swap.dataset.act = 'swap';
  swap.addEventListener('click', async () => { [F.a, F.b] = [F.b, F.a]; await closeViews(); reset(); sync(); });
  slots.append(slot('a', 'Original'), swap, slot('b', 'New version')); main.append(slots, result);

  function showResult() {
    const { A, B, diff } = last; result.textContent = '';
    const pagesChanged = new Set(diff.hunks.map((h) => (h.ins.length ? B.words[h.ins[0]].p : A.words[h.del[0]].p)));
    const sum = el('div', 'cmp-sum'); sum.id = 'cmpSum'; sum.dataset.added = String(diff.added); sum.dataset.removed = String(diff.removed); sum.dataset.changes = String(diff.hunks.length);
    sum.append(el('span', 'chip add', `+${plural(diff.added, 'word')} added`), el('span', 'chip del', `−${diff.removed} removed`), el('span', 'chip', `${plural(diff.hunks.length, 'change')} on ${plural(pagesChanged.size, 'page')}`));
    result.append(sum);
    if (!diff.hunks.length) { result.append(el('div', 'cmp-same', 'The text of the two files is the same.')); return; }
    if (!diff.exact) result.append(note('These files differ a lot, so they were compared page by page.'));
    const tabs = el('div', 'seg'); const tList = el('button', 'on', 'List of changes'); tList.dataset.v = 'list'; const tPages = el('button', '', 'Pages'); tPages.dataset.v = 'pages'; tabs.append(tList, tPages); tabs.id = 'cmpTabs';
    const list = el('div', 'cmp-list'); list.id = 'cmpList'; const pagesV = el('div', 'cmp-pages'); pagesV.hidden = true; result.append(tabs, list, pagesV);
    const ctx = (words, from, to) => words.slice(Math.max(0, from), Math.max(0, to)).map((w) => w.t).join(' ');
    const shown = diff.hunks.slice(0, 400);
    for (const h of shown) {
      const p = h.ins.length ? B.words[h.ins[0]].p : A.words[h.del[0]].p; const row = el('button', 'chg'); row.dataset.page = String(p + 1);
      const tx = el('div', 'tx'); const before = ctx(B.words, h.before - 5, h.before + 1); const after = h.after >= 0 ? ctx(B.words, h.after, h.after + 6) : '';
      if (before) tx.append(document.createTextNode(`${before} `));
      if (h.del.length) tx.append(el('del', '', h.del.map((i) => A.words[i].t).join(' ')), document.createTextNode(' '));
      if (h.ins.length) tx.append(el('ins', '', h.ins.map((i) => B.words[i].t).join(' ')), document.createTextNode(' '));
      if (after) tx.append(document.createTextNode(after));
      row.append(el('span', 'chip', `Page ${p + 1}`), tx); row.addEventListener('click', () => { setView('pages'); showPages(p + 1); }); list.append(row);
    }
    if (diff.hunks.length > shown.length) list.append(note(`Showing the first ${shown.length} changes. The marked-up PDF has all of them.`));
    const delB = mergeBoxes(A.words, diff.hunks.flatMap((h) => h.del)); const insB = mergeBoxes(B.words, diff.hunks.flatMap((h) => h.ins));
    const total = Math.max(A.pages, B.pages); let pageNo = 1;
    const cols = el('div', 'cmp-cols'); const pager = el('div', 'pager'); const prev = el('button', 'btn icon ghost'); prev.append(icon('i-back')); const next = el('button', 'btn icon ghost'); const nx = icon('i-back'); nx.style.transform = 'rotate(180deg)'; next.append(nx); const lab = el('span');
    pager.append(prev, lab, next); pagesV.append(cols, pager);
    async function showPages(n) {
      pageNo = clamp(n, 1, total); cols.textContent = ''; lab.textContent = `Page ${pageNo} of ${total}`; prev.disabled = pageNo === 1; next.disabled = pageNo === total;
      for (const [k, title, boxes, cls] of [['a', 'Original', delB, 'del'], ['b', 'New version', insB, 'ins']]) {
        const col = el('div', 'cmp-col'); col.append(el('b', '', title)); cols.append(col);
        if (!views[k]) views[k] = await loadView(F[k].bytes, F[k].name);
        if (pageNo > views[k].numPages) { col.append(note('No page here.')); continue; }
        const r = await renderPage(views[k], pageNo, { scale: 1.5, maxDim: 1300 }); const stg = el('div', 'pstage'); const layer = el('div', 'player'); r.canvas.style.width = '100%'; r.canvas.style.height = 'auto';
        for (const b of boxes) if (b.p === pageNo - 1) { const m = el('div', `cmark ${cls}`); m.style.left = `${b.x * 100}%`; m.style.top = `${b.y * 100}%`; m.style.width = `${b.w * 100}%`; m.style.height = `${b.h * 100}%`; layer.append(m); }
        stg.append(r.canvas, layer); col.append(stg);
      }
    }
    prev.addEventListener('click', () => showPages(pageNo - 1)); next.addEventListener('click', () => showPages(pageNo + 1));
    let drawn = false;
    const setView = (v) => { tList.classList.toggle('on', v === 'list'); tPages.classList.toggle('on', v === 'pages'); list.hidden = v !== 'list'; pagesV.hidden = v !== 'pages'; if (v === 'pages' && !drawn) { drawn = true; guard(() => showPages(pageNo)); } };
    tList.addEventListener('click', () => setView('list')); tPages.addEventListener('click', () => setView('pages'));
  }

  side.append(el('h3', '', 'Compare two versions'), note('Choose the original and the new version. Every word that was added or removed is listed.'));
  run = runner(side, 'Compare', async (r) => {
    if (!F.a || !F.b) throw new Error('Choose both files first.');
    await closeViews(); save.hidden = true; saveMsg.textContent = '';
    const A = await readWords(F.a.bytes, F.a.name, (v) => r.tick(v * 0.45)); const B = await readWords(F.b.bytes, F.b.name, (v) => r.tick(0.45 + v * 0.45));
    if (!A.words.length && !B.words.length) throw new Error('No text found in these PDFs. They may be scans, which are images of text.');
    const diff = compareWords(A, B); last = { A, B, diff }; showResult(); save.hidden = !diff.hunks.length;
    r.done(diff.hunks.length ? `Found ${plural(diff.hunks.length, 'change')}.` : 'No changes found.');
  });
  save = btn('btn', 'Save marked-up PDF', 'i-download'); save.dataset.act = 'markup'; save.hidden = true; saveMsg = el('div');
  save.addEventListener('click', async () => {
    if (!last) return; save.disabled = true; saveMsg.textContent = '';
    const out = await guard(() => opMarkup(F.a, F.b, last.A, last.B, last.diff));
    if (out) { const saved = await guard(() => saveOne(out, F.a.dir, `${stemOf(F.a.name)}-compared.pdf`)); if (saved) { const box = el('div', 'result', `Saved ${baseName(saved)}.`); const s = btn('btn', 'Show in folder'); s.addEventListener('click', () => api.invoke('fs:reveal', saved)); const acts = el('div', 'acts'); acts.append(s); box.append(acts); saveMsg.append(box); saveMsg.dataset.saved = saved; } }
    save.disabled = false;
  });
  side.append(save, saveMsg, note('Text is compared word by word. Changes to pictures, colours or layout are not listed.'));
  sync();
  const init = takeInitial(); if (init) { setFile('a', init[0]); if (init[1]) setFile('b', init[1]); }
}

function mountToText(tool) {
  onePdf(tool, async (ctx) => {
    const { f, side } = ctx; await preview(ctx);
    const what = radios('totext', [['text', 'Text', 'All the text in one .txt file, page by page'], ['images', 'Images', 'Every picture inside the PDF, saved at full size in a new folder']], 'text');
    side.append(el('h3', '', 'What to take out'), what);
    runner(side, 'Save', async (r) => {
      if (what.value() === 'text') {
        const out = await opToText(f.bytes, f.name, (v) => r.tick(v * 0.95));
        const saved = await saveOne(out.bytes, f.dir, `${stemOf(f.name)}.txt`); if (saved) r.done(`Saved ${baseName(saved)}: ${plural(out.words, 'word')} from ${plural(out.pages, 'page')}.`, [saved]);
      } else {
        const files = await opExtractImages(f.bytes, f.name, (v) => r.tick(v * 0.9));
        const dir = await api.invoke('fs:uniquePath', f.dir, `${stemOf(f.name)}-images`); await api.invoke('fs:makeFolder', dir); const paths = [];
        for (const x of files) { const t = await api.invoke('fs:uniquePath', dir, x.name); await api.invoke('fs:write', { file: t, data: x.bytes, force: true }); paths.push(t); }
        Y.refreshHome(); r.done(`Saved ${plural(paths.length, 'image')} to the folder ${baseName(dir)}, next to the PDF.`, [paths[0]]);
      }
    });
    side.append(note('Text is saved as plain lines, so columns and tables lose their layout. A scanned PDF is a picture of text, so it has no text to copy.'));
  });
}

const MM = 72 / 25.4;
function mountCrop(tool) {
  onePdf(tool, async ({ f, view, main, side }) => {
    let auto = false; let mode = 'crop';
    const masks = ['t', 'b', 'l', 'r'].map((k) => el('div', `cmask ${k}`)); const inputs = {};
    const typed = () => ({ top: (Number(inputs.top.value) || 0) * MM, right: (Number(inputs.right.value) || 0) * MM, bottom: (Number(inputs.bottom.value) || 0) * MM, left: (Number(inputs.left.value) || 0) * MM });
    const mask = () => {
      const on = mode === 'crop'; const m = typed(); const [t, b, l, r] = masks; for (const x of masks) x.style.display = on ? '' : 'none';
      const ft = clamp(m.top / st.ptH, 0, 1); const fb = clamp(m.bottom / st.ptH, 0, 1 - ft); const fl = clamp(m.left / st.ptW, 0, 1); const fr = clamp(m.right / st.ptW, 0, 1 - fl);
      t.style.cssText = `left:0;top:0;width:100%;height:${ft * 100}%`; b.style.cssText = `left:0;bottom:0;width:100%;height:${fb * 100}%`;
      l.style.cssText = `left:0;top:${ft * 100}%;width:${fl * 100}%;height:${(1 - ft - fb) * 100}%`; r.style.cssText = `right:0;top:${ft * 100}%;width:${fr * 100}%;height:${(1 - ft - fb) * 100}%`;
      if (!on) for (const x of masks) x.style.display = 'none';
    };
    const setAuto = async () => { const m = await autoMargins(view, st.page - 1); for (const k of ['top', 'right', 'bottom', 'left']) inputs[k].value = String(Math.round(m[k] / MM)); mask(); };
    const st = pageStage(view, { forms: true, onShow: async () => { st.layer.textContent = ''; st.layer.append(...masks); if (auto) await setAuto(); else mask(); } }); main.append(st.wrap, st.pager);

    const tabs = el('div', 'seg'); tabs.id = 'cropTabs'; const tc = el('button', 'on', 'Crop margins'); tc.dataset.m = 'crop'; const ts = el('button', '', 'Change page size'); ts.dataset.m = 'size'; tabs.append(tc, ts);
    const pc = el('div', 'pane'); const ps = el('div', 'pane'); ps.hidden = true; side.append(tabs, pc, ps);
    const setMode = (m) => { mode = m; tc.classList.toggle('on', m === 'crop'); ts.classList.toggle('on', m === 'size'); pc.hidden = m !== 'crop'; ps.hidden = m !== 'size'; mask(); };
    tc.addEventListener('click', () => setMode('crop')); ts.addEventListener('click', () => setMode('size'));

    const grid = el('div', 'mgrid');
    for (const [k, label] of [['top', 'Top'], ['bottom', 'Bottom'], ['left', 'Left'], ['right', 'Right']]) {
      const i = el('input', 'field'); i.type = 'number'; i.min = '0'; i.step = '1'; i.value = '0'; i.id = `crop-${k}`; inputs[k] = i;
      i.addEventListener('input', () => { auto = false; autoNote.hidden = true; mask(); }); grid.append(field(`${label} (mm)`, i));
    }
    const autoBtn = btn('btn', 'Find the white margins for me', 'i-search'); autoBtn.dataset.act = 'auto'; const autoNote = note('Each page gets its own margins when "All pages" is chosen.'); autoNote.hidden = true;
    autoBtn.addEventListener('click', async () => { auto = true; autoNote.hidden = false; await guard(setAuto); });
    const scope = radios('cropscope', [['all', 'All pages'], ['one', 'Only the page shown']], 'all');
    pc.append(el('h3', '', 'Margins to cut off'), grid, autoBtn, autoNote, el('h3', '', 'Pages'), scope);
    runner(pc, 'Save cropped PDF', async (r) => {
      const one = scope.value() === 'one'; const m = auto && !one ? 'auto' : auto ? await autoMargins(view, st.page - 1) : typed();
      if (m !== 'auto' && !(m.top || m.right || m.bottom || m.left)) throw new Error('Type a margin, or let the app find the white margins.');
      const out = await opCrop(f.bytes, f.name, m, one ? [st.page - 1] : 'all', (v) => r.tick(v * 0.95));
      const saved = await saveOne(out, f.dir, `${stemOf(f.name)}-cropped.pdf`); if (saved) r.done(`Saved ${baseName(saved)}.`, [saved]);
    });
    pc.append(note('Cropping hides the edges. The content outside is still in the file, so to remove private details use Redact PDF.'));

    const size = select([['a4', 'A4'], ['letter', 'Letter'], ['a5', 'A5'], ['legal', 'Legal']], 'a4'); size.id = 'cropSize';
    ps.append(el('h3', '', 'New page size'), field('Paper size', size), note('Each page is scaled to fit and centred. Wide pages stay wide.'));
    runner(ps, 'Save resized PDF', async (r) => {
      const out = await opResize(f.bytes, f.name, size.value, (v) => r.tick(v * 0.95));
      const saved = await saveOne(out, f.dir, `${stemOf(f.name)}-${size.value}.pdf`); if (saved) r.done(`Saved ${baseName(saved)}.`, [saved]);
    });
    ps.append(note('Form fields become part of the page. Links and comments are not kept in the resized file.'));
    await st.show();
  });
}

function mountFlatten(tool) {
  onePdf(tool, async (ctx) => {
    const { f, side } = ctx; await preview(ctx, { forms: true });
    const how = radios('flatmode', [['image', 'Whole PDF as images', 'Nothing can be selected, copied, searched or edited'], ['forms', 'Forms and comments only', 'Text stays selectable and links keep working']], 'image');
    const q = select([['150', 'Standard (150 DPI)'], ['200', 'Print (200 DPI)']], '150'); q.id = 'flatDpi'; const qf = field('Quality', q);
    const n1 = note('Every page becomes a PNG image, with forms, comments and stamps included. Text can no longer be selected, copied or searched, and links stop working. The file is usually bigger than the original.');
    const n2 = note('Form fields and comments become part of the page, so they can no longer be changed.'); n2.hidden = true;
    how.addEventListener('change', () => { const img = how.value() === 'image'; qf.hidden = !img; n1.hidden = !img; n2.hidden = img; });
    side.append(el('h3', '', 'How to flatten'), how, qf, n1, n2);
    runner(side, 'Save flattened PDF', async (r) => {
      let bytes; let text;
      if (how.value() === 'image') { bytes = await opFlattenImages(f.bytes, f.name, Number(q.value), (v) => r.tick(v * 0.95)); text = 'Every page is now an image.'; }
      else { const out = await opFlatten(f.bytes, f.name); bytes = out.bytes; text = `Flattened ${plural(out.fields, 'form field')} and ${plural(out.notes, 'comment')}.`; }
      const saved = await saveOne(bytes, f.dir, `${stemOf(f.name)}-flattened.pdf`); if (saved) r.done(`Saved ${baseName(saved)}. ${text}`, [saved]);
    });
  });
}

function mountGray(tool) {
  onePdf(tool, async (ctx) => {
    const { f, side } = ctx; await preview(ctx, { forms: true });
    const q = radios('graydpi', [['200', 'Print quality', '200 DPI, sharp on paper'], ['120', 'Smaller file', '120 DPI, good for the screen']], '200');
    side.append(el('h3', '', 'Quality'), q, note('Each page is saved as a black and white image, so text can no longer be selected.'));
    runner(side, 'Save black & white PDF', async (r) => {
      const out = await opGrayscale(f.bytes, f.name, Number(q.value()), (v) => r.tick(v * 0.95));
      const saved = await saveOne(out, f.dir, `${stemOf(f.name)}-grayscale.pdf`); if (saved) r.done(`Saved ${baseName(saved)}.`, [saved]);
    });
  });
}

function mountEditor(tool) {
  onePdf(tool, async ({ f, view, main, side, bar, other }) => {
    const full = btn('btn', 'Open in the full editor'); full.id = 'editFull'; full.title = 'Change the text and pictures that are already in the PDF'; full.addEventListener('click', () => Y.openFile(f.path)); bar.insertBefore(full, other);
    let items = []; let mode = ''; let sig = null; const vals = {}; let fields = [];
    const tools = el('div', 'seg etools'); tools.id = 'editTools';
    const TOOLS = [['text', 'Text'], ['hl', 'Highlight'], ['tick', 'Tick'], ['cross', 'Cross'], ['sig', 'Signature']];
    const HELP = { '': 'Pick a tool, then click on the page.', text: 'Click where the text should start, then type.', hl: 'Drag across what you want to highlight.', tick: 'Click to place a tick.', cross: 'Click to place a cross.', sig: 'Make your signature below, then add it to the page.' };
    const help = note(HELP['']); help.id = 'editHelp';
    const tsize = select([['9', '9 pt'], ['10', '10 pt'], ['12', '12 pt'], ['14', '14 pt'], ['18', '18 pt'], ['24', '24 pt']], '12'); tsize.id = 'editSize';
    const tcol = select([['#111111', 'Black'], ['#1b3a8c', 'Blue'], ['#c62828', 'Red']], '#111111'); tcol.id = 'editColor';
    const msize = select([['12', 'Small'], ['16', 'Medium'], ['22', 'Large']], '16'); msize.id = 'editMark';
    const pText = el('div', 'pane'); pText.append(field('Text size', tsize), field('Colour', tcol));
    const pMark = el('div', 'pane'); pMark.append(field('Mark size', msize), field('Colour', (() => { const s = select([['#111111', 'Black'], ['#1b3a8c', 'Blue'], ['#c62828', 'Red'], ['#1b8a3a', 'Green']], '#111111'); s.id = 'editMarkColor'; return s; })()));
    const padUi = signaturePad({ pad: 'editSigPad', type: 'editSigType' }); const pSig = el('div', 'pane');
    const sigAdd = btn('btn', 'Add to this page'); sigAdd.dataset.act = 'sig-add';
    pSig.append(padUi.tabs, padUi.panes.draw, padUi.panes.type, padUi.panes.image, field('Ink', padUi.ink), padUi.clear, sigAdd);
    const setMode = (m) => {
      mode = mode === m ? '' : m; for (const b of tools.children) b.classList.toggle('on', b.dataset.t === mode);
      help.textContent = HELP[mode]; pText.hidden = mode !== 'text'; pMark.hidden = !(mode === 'tick' || mode === 'cross'); pSig.hidden = mode !== 'sig';
      st.layer.classList.toggle('draw', mode !== '' && mode !== 'sig'); if (mode === 'sig') padUi.setMode('draw');
    };
    for (const [t, label] of TOOLS) { const b = el('button', '', label); b.dataset.t = t; b.addEventListener('click', () => setMode(t)); tools.append(b); }

    const drop = (it) => { items = items.filter((x) => x !== it); };
    function addEl(it, focus) {
      if (it.kind === 'hl') return fracBox(st.layer, it, 'ehl', { onRemove: () => drop(it) });
      if (it.kind === 'sig') { const d = fracBox(st.layer, it, 'esig', { onRemove: () => drop(it), ratio: it.ratio }); const im = el('img'); im.src = it.url; im.alt = 'Your signature'; d.prepend(im); return d; }
      if (it.kind === 'tick' || it.kind === 'cross') {
        const d = fracBox(st.layer, it, 'emark', { onRemove: () => drop(it), resizable: false }); d.style.width = d.style.height = `calc(var(--pt) * ${it.size})`;
        const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('viewBox', '0 0 24 24'); const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        p.setAttribute('d', MARKS[it.kind]); p.setAttribute('fill', 'none'); p.setAttribute('stroke', it.color); p.setAttribute('stroke-width', '3'); p.setAttribute('stroke-linecap', 'round'); svg.append(p); d.prepend(svg); return d;
      }
      const d = fracBox(st.layer, it, 'etext', { onRemove: () => drop(it), resizable: false }); const grip = el('span', 'grip', '⠿'); grip.title = 'Drag to move';
      const tx = el('div', 'etx'); tx.contentEditable = 'plaintext-only'; tx.spellcheck = false; tx.textContent = it.text || ''; tx.style.fontSize = `calc(var(--pt) * ${it.size})`; tx.style.color = it.color;
      tx.addEventListener('input', () => { it.text = tx.innerText.replace(/\r/g, '').replace(/\n$/, ''); });
      tx.addEventListener('blur', () => { if (!String(it.text || '').trim()) { d.remove(); drop(it); } });
      d.prepend(grip, tx); if (focus) setTimeout(() => tx.focus(), 0); return d;
    }
    function outlines() {
      for (const fl of fields) for (const r of fl.rects) {
        if (r.page !== st.page - 1) continue;
        const o = el('button', `fout${fl.readOnly ? ' ro' : ''}`); o.style.left = `${r.x * 100}%`; o.style.top = `${r.y * 100}%`; o.style.width = `${r.w * 100}%`; o.style.height = `${r.h * 100}%`; o.title = fl.label; o.dataset.key = fl.key;
        const v = fl.key in vals ? vals[fl.key] : null; if (v != null) o.textContent = fl.type === 'check' ? (v ? '✓' : '') : String(v);
        if (v != null) o.classList.add('set');
        o.addEventListener('pointerdown', (e) => e.stopPropagation());
        o.addEventListener('click', (e) => { e.stopPropagation(); const inp = formBox.querySelector(`[data-key="${CSS.escape(fl.key)}"]`); if (inp) { inp.scrollIntoView({ block: 'center' }); inp.focus(); if (fl.type === 'check' && !fl.readOnly) inp.click(); } });
        st.layer.append(o);
      }
    }
    const paint = () => { st.layer.textContent = ''; outlines(); for (const it of items) if (it.page === st.page - 1) addEl(it); };
    const st = pageStage(view, { forms: true, onShow: paint }); main.append(tools, st.wrap, st.pager);
    rubber(st.layer, 'ehl', () => mode === 'hl', (r) => { const it = { id: rid(), page: st.page - 1, kind: 'hl', ...r }; items.push(it); addEl(it); });
    st.layer.addEventListener('click', (e) => {
      if (e.target !== st.layer || !['text', 'tick', 'cross'].includes(mode)) return;
      const r = st.layer.getBoundingClientRect(); const x = (e.clientX - r.left) / r.width; const y = (e.clientY - r.top) / r.height;
      if (mode === 'text') { const size = Number(tsize.value); const it = { id: rid(), page: st.page - 1, kind: 'text', x, y: clamp(y - (size * 0.6) / st.ptH, 0, 1), text: '', size, color: tcol.value }; items.push(it); addEl(it, true); }
      else { const size = Number(msize.value); const it = { id: rid(), page: st.page - 1, kind: mode, x: clamp(x - size / 2 / st.ptW, 0, 1), y: clamp(y - size / 2 / st.ptH, 0, 1), size, color: document.getElementById('editMarkColor').value }; items.push(it); addEl(it); }
    });
    sigAdd.addEventListener('click', async () => {
      const s = await guard(padUi.makeSignature); if (!s) return; sig = s;
      const w = Math.min(0.32, 180 / st.ptW); const h = w * s.ratio * st.ptW / st.ptH;
      const it = { id: rid(), page: st.page - 1, kind: 'sig', x: clamp(0.58, 0, 1 - w), y: clamp(0.78, 0, 1 - h), w, h, png: s.png, url: s.url, ratio: s.ratio }; items.push(it); addEl(it);
      help.textContent = 'Drag the signature into place. Pull the dot to resize it.';
    });

    // form fields
    const formBox = el('div', 'eform'); formBox.id = 'editForm'; const lockL = el('label', 'chk'); const lock = el('input'); lock.type = 'checkbox'; lock.id = 'editLock';
    const lt = el('div'); lt.append(document.createTextNode('Lock the form after filling'), el('span', '', 'The answers become part of the page and can no longer be changed.')); lockL.append(lock, lt);
    fields = (await guard(() => formInfo(f.bytes, f.name))) || [];
    const setVal = (fl, v) => { vals[fl.key] = v; for (const o of st.layer.querySelectorAll('.fout')) if (o.dataset.key === fl.key) { o.textContent = fl.type === 'check' ? (v ? '✓' : '') : String(v); o.classList.add('set'); } };
    for (const fl of fields) {
      let inp;
      if (fl.type === 'check') { const l = el('label', 'chk'); inp = el('input'); inp.type = 'checkbox'; inp.checked = !!fl.value; inp.addEventListener('change', () => setVal(fl, inp.checked)); l.append(inp, document.createTextNode(fl.label)); formBox.append(l); }
      else if (fl.type === 'text') { inp = el(fl.multiline ? 'textarea' : 'input', 'field'); inp.value = String(fl.value || ''); if (fl.maxLength) inp.maxLength = fl.maxLength; inp.addEventListener('input', () => setVal(fl, inp.value)); formBox.append(field(fl.label, inp)); }
      else { inp = select([['', '(none)'], ...(fl.options || []).map((o) => [o, o])], String(fl.value || '')); inp.addEventListener('change', () => setVal(fl, inp.value)); formBox.append(field(fl.label, inp)); }
      inp.dataset.key = fl.key; inp.disabled = fl.readOnly;
    }
    side.append(el('h3', '', 'Add to the page'), help, pText, pMark, pSig);
    if (fields.length) side.append(el('h3', '', `Form fields (${fields.length})`), note('Fill them in here, or click a field on the page.'), formBox, lockL);
    const rep = replaceOption(side);
    runner(side, 'Save PDF', async (r) => {
      const use = items.filter((it) => it.kind !== 'text' || String(it.text || '').trim());
      if (!use.length && !Object.keys(vals).length && !lock.checked) throw new Error('Add something to the page or fill in a field first.');
      const out = await opEdit(f.bytes, { items: use, fields: vals, lock: lock.checked }, f.name); r.tick(0.9);
      const saved = await saveOut(out, f, `${stemOf(f.name)}-edited.pdf`, rep); if (saved) r.done(`Saved ${baseName(saved)}.`, [saved]);
    });
    side.append(note('To change text or pictures that are already in the PDF, use "Open in the full editor".'));
    setMode(''); pText.hidden = true; pMark.hidden = true; pSig.hidden = true;
    await st.show(); void sig;
  });
}

// ================================================================ register
window.YOSPDF.registerTools([
  { id: 'edit', name: 'Edit and annotate', desc: 'Text, highlights, ticks, signatures, forms', long: 'Add text, highlights, ticks, crosses and your signature, and fill in forms. Everything stays on this computer.', pick: 'Choose a PDF to edit', mount: mountEditor },
  { id: 'metadata', desc: 'Remove hidden details and AI labels', long: 'See, remove or add the hidden details inside PDF, Word, Excel, PowerPoint, JPG and PNG files, such as author, company, camera and location. Also removes AI labels (Content Credentials) from JPG and PNG pictures, without changing the picture.' },
  { id: 'redact', after: 'edit', icon: 'i-redact', name: 'Redact PDF', desc: 'Black out private details for good', long: 'Black out names, account numbers and signatures so they are removed for good, not just covered.', pick: 'Choose a PDF to redact', mount: mountRedact },
  { id: 'compare', after: 'redact', icon: 'i-compare', name: 'Compare PDF', desc: 'See what changed between two versions', long: 'Shows every word added or removed between two versions of a PDF.', mount: mountCompare },
  { id: 'pdf2text', after: 'pdf2img', icon: 'i-text', name: 'PDF to Text', desc: 'Copy out the text or the images', long: 'Save all the text of a PDF as a .txt file, or save every image inside it.', mount: mountToText },
  { id: 'crop', after: 'pagenum', icon: 'i-crop', name: 'Crop PDF', desc: 'Trim margins or change the page size', long: 'Cut off the margins, or fit every page to A4, Letter, A5 or Legal.', mount: mountCrop },
  { id: 'flatten', after: 'crop', icon: 'i-layers', name: 'Flatten PDF', desc: 'Lock forms, comments and text in place', long: 'Turn the whole PDF into images so nothing can be edited, or make only forms and comments part of the page.', mount: mountFlatten },
  { id: 'grayscale', after: 'flatten', icon: 'i-contrast', name: 'Black & white PDF', desc: 'Turn colour pages grey to save ink', long: 'Turn a colour PDF into shades of grey, ready for printing.', mount: mountGray },
]);
window.YOSPDF.extra = { findText, opRedact, readWords, compareWords, opMarkup, opToText, opExtractImages, opCrop, opResize, autoMargins, opFlatten, opFlattenImages, opGrayscale, formInfo, opEdit };
