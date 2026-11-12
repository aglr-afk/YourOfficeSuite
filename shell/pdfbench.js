/* YourOfficeSuite PDF Bench: everyday PDF tools that run fully on this computer. */
import * as pdfjs from './vendor/pdfjs/pdf.min.mjs';

pdfjs.GlobalWorkerOptions.workerSrc = new URL('./vendor/pdfjs/pdf.worker.min.mjs', import.meta.url).href;
const PDFJS_OPTS = { cMapUrl: new URL('./vendor/pdfjs/cmaps/', import.meta.url).href, cMapPacked: true,
  standardFontDataUrl: new URL('./vendor/pdfjs/standard_fonts/', import.meta.url).href, isEvalSupported: false };

const { PDFDocument, PDFName, PDFRawStream, PDFArray, PDFNumber, PDFDict, degrees } = window.PDFLib;
const Y = window.YOS;
const { api, el, icon, toast, errText, fmtSize, baseName, dirName, stemOf, extOf } = Y;
const $ = (id) => document.getElementById(id);
const IMG_EXT = ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif'];
const OFFICE_EXT = ['docx', 'xlsx', 'pptx', 'doc', 'xls', 'ppt', 'odt', 'ods', 'odp', 'rtf', 'txt', 'csv'];

// ---------------------------------------------------------------- file helpers
async function readFile(path) {
  const r = await api.invoke('fs:read', path);
  if (!r.size) throw new Error(`${r.name} is empty`);
  return { path, name: r.name, dir: dirName(path), bytes: new Uint8Array(r.data), size: r.size };
}
async function loadLib(bytes, name) {
  try { return await PDFDocument.load(bytes); } catch (e) {
    if (/encrypt/i.test(String(e && e.message))) throw new Error(`${name || 'This PDF'} is password-protected. Use Unlock in PDF Bench first.`);
    throw new Error(`${name || 'This PDF'} could not be read. It may be damaged.`);
  }
}
async function loadView(bytes, name) {
  try { return await pdfjs.getDocument({ data: bytes.slice(), ...PDFJS_OPTS }).promise; } catch (e) {
    if (e && e.name === 'PasswordException') throw new Error(`${name || 'This PDF'} is password-protected. Use Unlock in PDF Bench first.`);
    throw new Error(`${name || 'This PDF'} could not be read. It may be damaged.`);
  }
}
async function saveOne(bytes, dir, name) {
  const ext = extOf(name);
  const target = await api.invoke('dlg:saveAs', { defaultPath: dir ? await api.invoke('path:join', dir, name) : name, ext });
  if (!target) return null;
  await api.invoke('fs:write', { file: target, data: bytes, force: true });
  return target;
}
async function saveMany(files, title) {
  const dir = await api.invoke('dlg:pickFolder', title || 'Choose where to save the files');
  if (!dir) return null;
  const out = [];
  for (const f of files) {
    const target = await api.invoke('fs:uniquePath', dir, f.name);
    await api.invoke('fs:write', { file: target, data: f.bytes, force: true });
    out.push(target);
  }
  return out;
}
const pad = (n, w) => String(n).padStart(w, '0');
const canvasBlob = (c, type, q) => new Promise((res, rej) => c.toBlob((b) => (b ? res(b) : rej(new Error('Could not encode the image'))), type, q));
const blobBytes = async (b) => new Uint8Array(await b.arrayBuffer());

function parseRanges(text, max) {
  const out = new Set();
  for (const part of String(text).split(/[,;\s]+/).filter(Boolean)) {
    const m = part.match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!m) throw new Error(`"${part}" is not a page number or range`);
    const a = Number(m[1]); const b = m[2] ? Number(m[2]) : a;
    if (a < 1 || b < a || b > max) throw new Error(`Pages go from 1 to ${max}`);
    for (let i = a; i <= b; i++) out.add(i - 1);
  }
  return [...out].sort((x, y) => x - y);
}

// Where to draw something so it looks upright at a top-left based position on the page
// as the reader sees it, whatever rotation the page carries.
function visualSize(page) {
  const c = page.getCropBox(); const r = ((page.getRotation().angle % 360) + 360) % 360;
  return r % 180 ? { w: c.height, h: c.width } : { w: c.width, h: c.height };
}
function place(page, vx, vy, vw, vh) {
  const c = page.getCropBox(); const r = ((page.getRotation().angle % 360) + 360) % 360;
  if (r === 90) return { x: c.x + vy + vh, y: c.y + vx, rotate: degrees(90) };
  if (r === 180) return { x: c.x + c.width - vx, y: c.y + vy + vh, rotate: degrees(180) };
  if (r === 270) return { x: c.x + c.width - (vy + vh), y: c.y + c.height - vx, rotate: degrees(270) };
  return { x: c.x + vx, y: c.y + c.height - vy - vh, rotate: degrees(0) };
}
function textPng(text, { px = 64, color = '#000', font = '600 64px Inter, Arial, sans-serif' } = {}) {
  const c = document.createElement('canvas'); let g = c.getContext('2d'); g.font = font.replace(/\d+px/, px + 'px');
  const m = g.measureText(text); const w = Math.ceil(m.width) + 8; const h = Math.ceil(px * 1.3);
  c.width = Math.max(2, w); c.height = h; g = c.getContext('2d'); g.font = font.replace(/\d+px/, px + 'px'); g.fillStyle = color; g.textBaseline = 'middle';
  g.fillText(text, 4, h / 2);
  return { canvas: c, w: c.width, h };
}
async function renderPage(view, n, { scale, maxDim } = {}) {
  const page = await view.getPage(n);
  let vp = page.getViewport({ scale: scale || 1 });
  if (maxDim && Math.max(vp.width, vp.height) > maxDim) vp = page.getViewport({ scale: (scale || 1) * maxDim / Math.max(vp.width, vp.height) });
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(vp.width)); c.height = Math.max(1, Math.round(vp.height));
  const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: g, viewport: vp }).promise;
  const base = page.getViewport({ scale: 1 });
  page.cleanup();
  return { canvas: c, ptW: base.width, ptH: base.height, scale: vp.scale };
}

// ---------------------------------------------------------------- operations
async function opMerge(files) {
  const out = await PDFDocument.create();
  for (const f of files) {
    const src = await loadLib(f.bytes, f.name);
    for (const p of await out.copyPages(src, src.getPageIndices())) out.addPage(p);
  }
  return out.save({ useObjectStreams: true });
}
async function opPick(bytes, order, name) {
  // order: [{ src: zero-based page, rot: extra clockwise degrees }]
  const src = await loadLib(bytes, name); const out = await PDFDocument.create();
  const pages = await out.copyPages(src, order.map((o) => o.src));
  pages.forEach((p, i) => { const r = (p.getRotation().angle + (order[i].rot || 0)) % 360; p.setRotation(degrees(r)); out.addPage(p); });
  return out.save({ useObjectStreams: true });
}
async function inflate(u8) {
  const s = new Blob([u8]).stream().pipeThrough(new DecompressionStream('deflate'));
  return new Uint8Array(await new Response(s).arrayBuffer());
}
function unPredict(data, cols, comps) {
  const row = cols * comps; const rows = Math.floor(data.length / (row + 1)); const out = new Uint8Array(rows * row);
  for (let y = 0; y < rows; y++) {
    const ft = data[y * (row + 1)]; const inOff = y * (row + 1) + 1; const o = y * row;
    for (let x = 0; x < row; x++) {
      const raw = data[inOff + x]; const a = x >= comps ? out[o + x - comps] : 0; const b = y ? out[o - row + x] : 0; const c = y && x >= comps ? out[o - row + x - comps] : 0;
      let v = raw;
      if (ft === 1) v = raw + a; else if (ft === 2) v = raw + b; else if (ft === 3) v = raw + ((a + b) >> 1);
      else if (ft === 4) { const p = a + b - c; const pa = Math.abs(p - a); const pb = Math.abs(p - b); const pc = Math.abs(p - c); v = raw + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c); }
      out[o + x] = v & 255;
    }
  }
  return out;
}
function ascii85(src) {
  const out = []; let tuple = 0; let count = 0;
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (c === 126) break;                       // ~> ends the data
    if (c <= 32) continue;
    if (c === 122 && count === 0) { out.push(0, 0, 0, 0); continue; }
    if (c < 33 || c > 117) throw new Error('bad ascii85');
    tuple = tuple * 85 + (c - 33); count++;
    if (count === 5) { out.push((tuple >>> 24) & 255, (tuple >>> 16) & 255, (tuple >>> 8) & 255, tuple & 255); tuple = 0; count = 0; }
  }
  if (count > 1) { for (let k = count; k < 5; k++) tuple = tuple * 85 + 84; for (let k = 0; k < count - 1; k++) out.push((tuple >>> (24 - 8 * k)) & 255); }
  return new Uint8Array(out);
}
const N = (s) => PDFName.of(s);
const numOf = (d, k) => { const v = d.lookup(N(k)); return v instanceof PDFNumber ? v.asNumber() : 0; };
function channels(dict) {
  const cs = dict.lookup(N('ColorSpace'));
  if (cs instanceof PDFName) return cs === N('DeviceRGB') ? 3 : cs === N('DeviceGray') ? 1 : 0;
  if (cs instanceof PDFArray && cs.size() >= 2 && cs.lookup(0) === N('ICCBased')) {
    const st = cs.lookup(1); const n = st && st.dict ? numOf(st.dict, 'N') : 0; return n === 3 || n === 1 ? n : 0;
  }
  return 0;
}
async function recompressImages(doc, { maxDim, quality }, tick) {
  const items = doc.context.enumerateIndirectObjects().filter(([, o]) => o instanceof PDFRawStream && o.dict.lookup(N('Subtype')) === N('Image'));
  let changed = 0; let i = 0;
  for (const [ref, obj] of items) {
    if (tick) tick(++i / Math.max(1, items.length));
    try {
      const d = obj.dict; const w = numOf(d, 'Width'); const h = numOf(d, 'Height');
      if (!w || !h || w * h < 40000 || w * h > 60e6) continue;
      if (d.has(N('ImageMask')) || d.has(N('Decode')) || d.has(N('Mask'))) continue;
      const comps = channels(d); if (!comps) continue;
      let f = d.lookup(N('Filter')); let data = obj.contents;
      if (f instanceof PDFArray) {
        if (f.size() === 2 && f.lookup(0) === N('ASCII85Decode')) { data = ascii85(data); f = f.lookup(1); } else f = f.size() === 1 ? f.lookup(0) : null;
      }
      let bitmap;
      if (f === N('DCTDecode')) bitmap = await createImageBitmap(new Blob([data], { type: 'image/jpeg' }));
      else if (f === N('FlateDecode')) {
        if (numOf(d, 'BitsPerComponent') !== 8) continue;
        let raw = await inflate(data);
        const parms = d.lookup(N('DecodeParms'));
        if (parms instanceof PDFDict) {
          const pred = numOf(parms, 'Predictor') || 1;
          if (pred >= 10) raw = unPredict(raw, numOf(parms, 'Columns') || w, numOf(parms, 'Colors') || comps); else if (pred !== 1) continue;
        }
        if (raw.length < w * h * comps) continue;
        const px = new Uint8ClampedArray(w * h * 4);
        for (let p = 0, q = 0; p < w * h; p++, q += 4) {
          if (comps === 3) { px[q] = raw[p * 3]; px[q + 1] = raw[p * 3 + 1]; px[q + 2] = raw[p * 3 + 2]; } else { px[q] = px[q + 1] = px[q + 2] = raw[p]; }
          px[q + 3] = 255;
        }
        bitmap = await createImageBitmap(new ImageData(px, w, h));
      } else continue;
      const k = Math.min(1, maxDim / Math.max(w, h)); const tw = Math.max(1, Math.round(w * k)); const th = Math.max(1, Math.round(h * k));
      const c = new OffscreenCanvas(tw, th); const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, tw, th); g.drawImage(bitmap, 0, 0, tw, th); bitmap.close();
      const jpg = await blobBytes(await c.convertToBlob({ type: 'image/jpeg', quality }));
      if (jpg.length >= obj.contents.length * 0.85) continue;
      d.set(N('Width'), PDFNumber.of(tw)); d.set(N('Height'), PDFNumber.of(th)); d.set(N('ColorSpace'), N('DeviceRGB'));
      d.set(N('BitsPerComponent'), PDFNumber.of(8)); d.set(N('Filter'), N('DCTDecode')); d.delete(N('DecodeParms'));
      doc.context.assign(ref, PDFRawStream.of(d, jpg)); changed++;
    } catch (_) { /* leave this image as it is */ }
  }
  return changed;
}
async function opCompress(bytes, level, name, tick) {
  if (level === 'strong') {
    const view = await loadView(bytes, name); const out = await PDFDocument.create();
    for (let n = 1; n <= view.numPages; n++) {
      const r = await renderPage(view, n, { scale: 110 / 72, maxDim: 2400 });
      const img = await out.embedJpg(await blobBytes(await canvasBlob(r.canvas, 'image/jpeg', 0.6)));
      out.addPage([r.ptW, r.ptH]).drawImage(img, { x: 0, y: 0, width: r.ptW, height: r.ptH });
      if (tick) tick(n / view.numPages);
    }
    await view.destroy();
    return out.save({ useObjectStreams: true });
  }
  const doc = await loadLib(bytes, name);
  await recompressImages(doc, { maxDim: 1600, quality: 0.68 }, tick);
  return doc.save({ useObjectStreams: true });
}
async function opToImages(bytes, { format, dpi }, name, tick) {
  const view = await loadView(bytes, name); const out = []; const w = String(view.numPages).length < 3 ? 3 : String(view.numPages).length;
  for (let n = 1; n <= view.numPages; n++) {
    const r = await renderPage(view, n, { scale: dpi / 72, maxDim: 6000 });
    const blob = await canvasBlob(r.canvas, format === 'png' ? 'image/png' : 'image/jpeg', 0.9);
    out.push({ name: `${stemOf(name)}-page-${pad(n, w)}.${format}`, bytes: await blobBytes(blob) });
    if (tick) tick(n / view.numPages);
  }
  await view.destroy();
  return out;
}
async function opFromImages(files, { size }, tick) {
  const out = await PDFDocument.create(); let i = 0;
  for (const f of files) {
    const ext = extOf(f.name); let img;
    if (ext === 'jpg' || ext === 'jpeg') { try { img = await out.embedJpg(f.bytes); } catch (_) { img = null; } }
    if (!img && ext === 'png') { try { img = await out.embedPng(f.bytes); } catch (_) { img = null; } }
    if (!img) {
      const bmp = await createImageBitmap(new Blob([f.bytes])); const c = document.createElement('canvas'); c.width = bmp.width; c.height = bmp.height;
      c.getContext('2d').drawImage(bmp, 0, 0); img = await out.embedPng(await blobBytes(await canvasBlob(c, 'image/png')));
    }
    if (size === 'a4') {
      const land = img.width > img.height; const pw = land ? 841.89 : 595.28; const ph = land ? 595.28 : 841.89; const m = 28;
      const k = Math.min((pw - 2 * m) / img.width, (ph - 2 * m) / img.height); const w = img.width * k; const h = img.height * k;
      out.addPage([pw, ph]).drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
    } else {
      const k = 72 / 96; out.addPage([img.width * k, img.height * k]).drawImage(img, { x: 0, y: 0, width: img.width * k, height: img.height * k });
    }
    if (tick) tick(++i / files.length);
  }
  return out.save({ useObjectStreams: true });
}
async function opWatermark(bytes, { text, opacity, sizeK }, name) {
  const doc = await loadLib(bytes, name); const t = textPng(text, { px: 96, color: '#5b5870' });
  const img = await doc.embedPng(await blobBytes(await canvasBlob(t.canvas, 'image/png')));
  for (const page of doc.getPages()) {
    const v = visualSize(page); const w = Math.min(v.w, v.h) * sizeK; const h = w * t.h / t.w;
    // centre of the page, tilted 35 degrees as the reader sees it
    const a = 35 * Math.PI / 180; const cx = v.w / 2; const cy = v.h / 2;
    const vx = cx - (w / 2) * Math.cos(a) + (h / 2) * Math.sin(a); const vyBottom = cy + (w / 2) * Math.sin(a) + (h / 2) * Math.cos(a);
    const p = place(page, vx, vyBottom, 0, 0);
    page.drawImage(img, { x: p.x, y: p.y, width: w, height: h, rotate: degrees(p.rotate.angle + 35), opacity });
  }
  return doc.save({ useObjectStreams: true });
}
async function opPageNumbers(bytes, { pos, start, style }, name) {
  const doc = await loadLib(bytes, name); const pages = doc.getPages(); const total = pages.length;
  for (let i = 0; i < total; i++) {
    const page = pages[i]; const label = style === 'of' ? `${i + start} / ${total + start - 1}` : String(i + start);
    const t = textPng(label, { px: 48, color: '#2a2640', font: '500 48px Inter, Arial, sans-serif' });
    const img = await doc.embedPng(await blobBytes(await canvasBlob(t.canvas, 'image/png')));
    const v = visualSize(page); const h = 13; const w = h * t.w / t.h; const m = 26;
    const vx = pos.endsWith('right') ? v.w - m - w : pos.endsWith('left') ? m : (v.w - w) / 2; const vy = pos.startsWith('top') ? m - 6 : v.h - m - h + 6;
    const p = place(page, vx, vy, w, h);
    page.drawImage(img, { x: p.x, y: p.y, width: w, height: h, rotate: p.rotate });
  }
  return doc.save({ useObjectStreams: true });
}
async function opSign(bytes, { pageIndex, png, vx, vy, vw, vh }, name) {
  const doc = await loadLib(bytes, name); const page = doc.getPage(pageIndex); const img = await doc.embedPng(png);
  const p = place(page, vx, vy, vw, vh);
  page.drawImage(img, { x: p.x, y: p.y, width: vw, height: vh, rotate: p.rotate });
  return doc.save({ useObjectStreams: true });
}

// ---------------------------------------------------------------- UI pieces
function btn(cls, label, ic) { const b = el('button', cls, label); if (ic) b.prepend(icon(ic)); return b; }
function field(label, input) { const l = el('label', 'f', label); l.append(input); return l; }
function select(options, value) { const s = el('select', 'field'); for (const [v, t] of options) { const o = el('option', '', t); o.value = v; s.append(o); } if (value) s.value = value; return s; }
function radios(name, options, value) {
  const wrap = el('div'); wrap.style.display = 'flex'; wrap.style.flexDirection = 'column'; wrap.style.gap = '8px';
  for (const [v, title, desc] of options) {
    const l = el('label', 'opt'); const i = el('input'); i.type = 'radio'; i.name = name; i.value = v; i.checked = v === value;
    const tx = el('div'); tx.append(el('b', '', title)); if (desc) tx.append(el('span', '', desc)); l.append(i, tx); wrap.append(l);
  }
  wrap.value = () => wrap.querySelector('input:checked').value;
  return wrap;
}
function dropZone({ kind, multi, title, hint, onPaths }) {
  const exts = kind === 'pdf' ? ['pdf'] : kind === 'image' ? IMG_EXT : kind === 'meta' ? window.YOSMETA.TYPES : OFFICE_EXT;
  const z = el('div', 'drop'); const ic = el('div', 'big-ico'); ic.append(icon('i-open'));
  const b = btn('btn primary big', multi ? 'Choose files' : 'Choose a file'); b.dataset.act = 'choose';
  z.append(ic, el('b', '', title), el('span', '', hint || 'or drop it here'), b);
  b.addEventListener('click', async () => { const p = await api.invoke('dlg:open', { kind, multi: !!multi, title }); if (p.length) onPaths(p); });
  z.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); z.classList.add('over'); });
  z.addEventListener('dragleave', () => z.classList.remove('over'));
  z.addEventListener('drop', (e) => {
    e.preventDefault(); e.stopPropagation(); z.classList.remove('over');
    const paths = [...e.dataTransfer.files].map((f) => api.pathForFile(f)).filter((p) => p && exts.includes(extOf(baseName(p))));
    if (!paths.length) { toast('That file type does not fit this tool', true); return; }
    onPaths(multi ? paths : paths.slice(0, 1));
  });
  return z;
}
function runner(side, label, fn) {
  const prog = el('div', 'progress'); const bar = el('i'); prog.append(bar); prog.hidden = true;
  const b = btn('btn primary big', label); b.dataset.act = 'run';
  const res = el('div'); side.append(b, prog, res);
  const api2 = {
    button: b, tick: (v) => { bar.style.width = `${Math.round(v * 100)}%`; }, enable: (on) => { b.disabled = !on; },
    done: (text, paths) => {
      res.textContent = ''; const box = el('div', 'result'); box.append(el('div', '', text));
      if (paths && paths.length) {
        const acts = el('div', 'acts');
        if (paths.length === 1 && extOf(baseName(paths[0])) === 'pdf') { const o = btn('btn', 'Open'); o.addEventListener('click', () => Y.openFile(paths[0])); acts.append(o); }
        const s = btn('btn', 'Show in folder'); s.addEventListener('click', () => api.invoke('fs:reveal', paths[0])); acts.append(s); box.append(acts);
      }
      res.append(box);
    },
    fail: (text) => { res.textContent = ''; res.append(el('div', 'result err', text)); },
    clear: () => { res.textContent = ''; },
  };
  b.addEventListener('click', async () => {
    b.disabled = true; prog.hidden = false; api2.tick(0.03); api2.clear();
    try { await fn(api2); } catch (e) { api2.fail(errText(e)); }
    prog.hidden = true; b.disabled = false;
  });
  return api2;
}
function fileList(host, files, { onChange, sub }) {
  const render = () => {
    host.textContent = '';
    files.forEach((f, i) => {
      const it = el('div', 'fitem'); const g = el('div', 'grow'); g.append(el('b', '', f.name), el('span', '', sub ? sub(f) : fmtSize(f.size)));
      const st = el('span', 'state ' + (f.state === 'done' ? 'ok' : f.state === 'failed' ? 'err' : ''), f.stateText || ''); f.stateEl = st;
      const up = el('button', 'mini'); up.append(icon('i-up')); up.title = 'Move up'; up.disabled = i === 0;
      const dn = el('button', 'mini'); dn.append(icon('i-down')); dn.title = 'Move down'; dn.disabled = i === files.length - 1;
      const rm = el('button', 'mini'); rm.append(icon('i-x')); rm.title = 'Remove';
      up.addEventListener('click', () => { files.splice(i - 1, 0, files.splice(i, 1)[0]); render(); onChange(); });
      dn.addEventListener('click', () => { files.splice(i + 1, 0, files.splice(i, 1)[0]); render(); onChange(); });
      rm.addEventListener('click', () => { files.splice(i, 1); render(); onChange(); });
      it.append(Y.badge(f.kind || 'pdf', f.icon || 'i-pdf'), g, st, up, dn, rm); host.append(it);
    });
  };
  render(); return render;
}
// Thumbnail grid of one PDF. mode: 'select' (tick pages) or 'edit' (rotate, delete, reorder).
function pageGrid(host, view, mode, onChange) {
  const pages = Array.from({ length: view.numPages }, (_, i) => ({ src: i, rot: 0, sel: false }));
  const cache = new Map(); let dragFrom = -1;
  const io = new IntersectionObserver((entries) => {
    for (const en of entries) if (en.isIntersecting) { io.unobserve(en.target); draw(en.target); }
  }, { root: host.closest('.page'), rootMargin: '300px' });
  // a page turned on its side is scaled down so it still fits its tile
  const turn = (rot, c) => `rotate(${rot}deg)${Number(rot) % 180 ? ` scale(${Math.min(1, c.height ? Math.min(c.width, 130) / c.height : 1).toFixed(3)})` : ''}`;
  async function draw(cell) {
    const src = Number(cell.dataset.src); const holder = cell.querySelector('.cv');
    try {
      if (!cache.has(src)) cache.set(src, renderPage(view, src + 1, { maxDim: 300 }).then((r) => r.canvas));
      const base = await cache.get(src); const c = document.createElement('canvas'); c.width = base.width; c.height = base.height; c.getContext('2d').drawImage(base, 0, 0);
      c.style.transform = turn(cell.dataset.rot, c); holder.textContent = ''; holder.append(c);
    } catch (_) { holder.textContent = 'Preview unavailable'; }
  }
  function render() {
    io.disconnect(); host.textContent = '';
    pages.forEach((p, i) => {
      const cell = el('div', 'pg' + (p.sel ? ' sel' : '')); cell.dataset.src = p.src; cell.dataset.rot = p.rot; cell.dataset.idx = i;
      const cv = el('div', 'cv'); const tick = el('span', 'tick'); tick.append(icon('i-check'));
      cell.append(tick, cv, el('div', 'no', mode === 'edit' && p.src !== i ? `${i + 1}  (was ${p.src + 1})` : String(i + 1)));
      if (mode === 'select') cv.addEventListener('click', () => { p.sel = !p.sel; cell.classList.toggle('sel', p.sel); onChange(); });
      if (mode === 'edit') {
        const acts = el('div', 'acts');
        const mk = (ic, title, fn, dis) => { const b = el('button', 'mini'); b.append(icon(ic)); b.title = title; b.disabled = !!dis; b.dataset.act = title.split(' ')[0].toLowerCase(); b.addEventListener('click', fn); acts.append(b); };
        mk('i-back', 'Earlier in the file', () => { pages.splice(i - 1, 0, pages.splice(i, 1)[0]); render(); onChange(); }, i === 0);
        mk('i-rotate', 'Rotate right', () => { p.rot = (p.rot + 90) % 360; cell.dataset.rot = p.rot; const c = cv.querySelector('canvas'); if (c) c.style.transform = turn(p.rot, c); onChange(); });
        mk('i-trash', 'Delete this page', () => { pages.splice(i, 1); render(); onChange(); }, pages.length === 1);
        cell.append(acts); cell.draggable = true;
        cell.addEventListener('dragstart', (e) => { dragFrom = i; cell.classList.add('drag'); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', String(i)); });
        cell.addEventListener('dragend', () => { cell.classList.remove('drag'); dragFrom = -1; });
        cell.addEventListener('dragover', (e) => { if (dragFrom >= 0) { e.preventDefault(); e.stopPropagation(); cell.classList.add('over'); } });
        cell.addEventListener('dragleave', () => cell.classList.remove('over'));
        cell.addEventListener('drop', (e) => { if (dragFrom < 0) return; e.preventDefault(); e.stopPropagation(); const from = dragFrom; dragFrom = -1; if (from !== i) { pages.splice(i, 0, pages.splice(from, 1)[0]); render(); onChange(); } });
      }
      host.append(cell); io.observe(cell);
    });
  }
  render();
  return { pages, render, selected: () => pages.filter((p) => p.sel).map((p) => p.src),
    select: (list) => { const s = new Set(list); for (const p of pages) p.sel = s.has(p.src); render(); onChange(); },
    changed: () => pages.length !== view.numPages || pages.some((p, i) => p.src !== i || p.rot) };
}

// A file handed over from elsewhere in the app (right-click menu) for the tool being opened.
let pendingPaths = null;
const takeInitial = () => { const p = pendingPaths; pendingPaths = null; return p; };
function replaceOption(side) {
  const l = el('label', 'chk'); const c = el('input'); c.type = 'checkbox'; c.dataset.act = 'replace';
  const t = el('div'); t.append(document.createTextNode('Replace the original file'), el('span', '', 'No copy is made, so the old version is overwritten.')); l.append(c, t); side.append(l); return c;
}
async function confirmReplace(names) {
  const r = await api.invoke('dlg:message', { type: 'warning', buttons: ['Replace', 'Cancel'], defaultId: 1, cancelId: 1, title: 'Replace the original',
    message: names.length === 1 ? `Replace "${names[0]}" with the new version?` : `Replace ${names.length} original files with the new versions?`, detail: 'The original will be overwritten.' });
  return r === 0;
}
// Saves a result either over the original (when the box is ticked) or as a new file the user names.
async function saveOut(bytes, f, newName, chk) {
  if (!chk || !chk.checked) return saveOne(bytes, f.dir, newName);
  if (Y.isOpen(f.path)) throw new Error('This file is open in a tab. Close it first, then replace it.');
  if (!(await confirmReplace([f.name]))) return null;
  await api.invoke('fs:write', { file: f.path, data: bytes, force: true }); f.bytes = bytes; f.size = bytes.length; Y.refreshHome();
  return f.path;
}

// ---------------------------------------------------------------- tool workspaces
function workspace(tool) {
  const root = $('pdfTool'); root.textContent = '';
  const head = el('div', 'tool-head'); const back = el('button', 'btn icon ghost'); back.append(icon('i-back')); back.title = 'All tools'; back.id = 'pdfBack';
  back.addEventListener('click', showTools);
  const tx = el('div'); tx.append(el('h1', '', tool.name), el('p', '', tool.long || tool.desc)); head.append(back, tx);
  const body = el('div', 'tool-body'); const main = el('div', 'tool-main'); const side = el('div', 'tool-side'); side.hidden = true;
  body.append(main, side); root.append(head, body);
  $('pdfHome').hidden = true; root.hidden = false; root.dataset.tool = tool.id;
  return { main, side };
}
function showTools() { $('pdfTool').hidden = true; $('pdfTool').textContent = ''; $('pdfHome').hidden = false; }
const guard = async (fn) => { try { return await fn(); } catch (e) { toast(errText(e), true); return null; } };

// One PDF in, shown as page thumbnails. build(ctx) adds the options and the run button.
function singlePdfTool(tool, mode, build) {
  const { main, side } = workspace(tool);
  const start = () => { main.textContent = ''; side.hidden = true; side.textContent = ''; main.append(dropZone({ kind: 'pdf', title: tool.pick || 'Choose a PDF', onPaths: (p) => load(p[0]) })); };
  async function load(path) {
    const f = await guard(() => readFile(path)); if (!f) return;
    const view = await guard(() => loadView(f.bytes, f.name)); if (!view) return;
    main.textContent = ''; side.textContent = ''; side.hidden = false;
    const bar = el('div', 'bar2'); const other = btn('btn ghost', 'Choose another file'); other.addEventListener('click', start);
    bar.append(el('b', '', f.name), el('span', 'chip', `${view.numPages} page${view.numPages > 1 ? 's' : ''} · ${fmtSize(f.size)}`), el('div', 'grow'), other); main.append(bar);
    let grid = null;
    if (mode) { const host = el('div', 'pgrid'); main.append(host); grid = pageGrid(host, view, mode, () => ctx.onChange && ctx.onChange()); }
    const ctx = { f, view, grid, main, side, bar, onChange: null };
    build(ctx);
  }
  start();
  const init = takeInitial(); if (init) load(init[0]);
}

const TOOLS = [
  { id: 'merge', icon: 'i-merge', name: 'Merge PDFs', desc: 'Join several PDFs into one', long: 'Add the files, put them in order, and join them into a single PDF.', mount: mountMerge },
  { id: 'split', icon: 'i-split', name: 'Split or extract', desc: 'Pull out the pages you need', long: 'Tick the pages to keep, or save every page as its own file.', mount: mountSplit },
  { id: 'organize', icon: 'i-organize', name: 'Organise pages', desc: 'Reorder, rotate and delete pages', long: 'Drag pages into a new order, rotate them, or remove the ones you do not need.', mount: mountOrganize },
  { id: 'compress', icon: 'i-compress', name: 'Compress', desc: 'Make a PDF smaller to send', long: 'Shrink a PDF so it is easier to email or upload.', mount: mountCompress },
  { id: 'sign', icon: 'i-sign', name: 'Sign', desc: 'Add your signature to a PDF', long: 'Draw, type or upload your signature, then drop it where it belongs.', mount: mountSign },
  { id: 'edit', icon: 'i-edit', name: 'Edit and annotate', desc: 'Text, highlights, notes, forms', long: 'Opens the PDF in the editor to add text, highlight, comment and fill in forms.', mount: mountEdit },
  { id: 'office2pdf', icon: 'i-convert', name: 'Word, Excel, PowerPoint to PDF', desc: 'Turn documents into PDFs', long: 'Convert documents, spreadsheets and presentations to PDF. Each PDF is saved next to its original.', mount: mountOffice },
  { id: 'pdf2word', icon: 'i-doc', name: 'PDF to Word', desc: 'Get an editable document', long: 'Turn a PDF into a Word document you can edit. Works best on PDFs made from text, not scans.', mount: mountToWord },
  { id: 'pdf2img', icon: 'i-image', name: 'PDF to images', desc: 'Save pages as JPG or PNG', long: 'Save every page of a PDF as a picture.', mount: mountToImages },
  { id: 'img2pdf', icon: 'i-image', name: 'Images to PDF', desc: 'Photos and scans into one PDF', long: 'Combine pictures into a single PDF, one per page.', mount: mountFromImages },
  { id: 'watermark', icon: 'i-watermark', name: 'Watermark', desc: 'Stamp text across every page', long: 'Add a text stamp such as DRAFT or CONFIDENTIAL across every page.', mount: mountWatermark },
  { id: 'pagenum', icon: 'i-hash', name: 'Page numbers', desc: 'Number the pages', long: 'Add page numbers to every page.', mount: mountPageNumbers },
  { id: 'protect', icon: 'i-lock', name: 'Protect with a password', desc: 'Lock a PDF so only you can open it', long: 'Add a password that is needed to open the PDF.', mount: mountProtect },
  { id: 'unlock', icon: 'i-unlock', name: 'Unlock', desc: 'Remove a password you know', long: 'Remove the password from a PDF. You need to know the current password.', mount: mountUnlock },
  { id: 'metadata', icon: 'i-tag', name: 'File details (metadata)', desc: 'Remove or add hidden details', long: 'See, remove or add the hidden details inside PDF, Word, Excel, PowerPoint, JPG and PNG files, such as author, company, camera and location.', mount: mountMetadata },
];

function mountMerge(tool) {
  const { main, side } = workspace(tool); const files = []; const list = el('div', 'flist'); let render = () => {}; let run;
  const sync = () => { side.hidden = !files.length; zone.style.display = files.length ? 'none' : ''; more.style.display = files.length ? '' : 'none'; if (run) run.enable(files.length >= 2); hint.textContent = files.length < 2 ? 'Add at least two PDFs.' : `${files.length} files, ${files.reduce((a, f) => a + f.pages, 0)} pages in total.`; };
  const add = async (paths) => {
    for (const p of paths) { const f = await guard(async () => { const x = await readFile(p); x.pages = (await loadLib(x.bytes, x.name)).getPageCount(); return x; }); if (f) files.push(f); }
    render(); sync();
  };
  const zone = dropZone({ kind: 'pdf', multi: true, title: 'Choose the PDFs to join', hint: 'or drop them here', onPaths: add });
  const more = btn('btn addmore', 'Add more files', 'i-plus'); more.addEventListener('click', async () => { const p = await api.invoke('dlg:open', { kind: 'pdf', multi: true }); if (p.length) add(p); });
  main.append(zone, list, more);
  const hint = el('div', 'hint'); side.append(el('h3', '', 'Order'), el('div', 'hint', 'Files are joined from top to bottom. Use the arrows to change the order.'), hint);
  render = fileList(list, files, { onChange: sync, sub: (f) => `${f.pages} page${f.pages > 1 ? 's' : ''} · ${fmtSize(f.size)}` });
  run = runner(side, 'Merge PDFs', async (r) => {
    const out = await opMerge(files); r.tick(0.9);
    const saved = await saveOne(out, files[0].dir, `${stemOf(files[0].name)}-merged.pdf`);
    if (saved) r.done(`Merged ${files.length} files into ${baseName(saved)} (${fmtSize(out.length)}).`, [saved]);
  });
  sync();
  const init = takeInitial(); if (init) add(init);
}

function mountSplit(tool) {
  singlePdfTool(tool, 'select', (c) => {
    const mode = radios('splitmode', [['extract', 'Extract ticked pages', 'One new PDF with just those pages'], ['each', 'One PDF per page', 'Every page becomes its own file']], 'extract');
    const range = el('input', 'field'); range.placeholder = 'For example 1-3, 6'; range.id = 'splitRange';
    const hint = el('div', 'hint');
    const all = btn('btn ghost', 'Tick all'); const none = btn('btn ghost', 'Clear'); c.bar.insertBefore(none, c.bar.lastChild); c.bar.insertBefore(all, none);
    all.addEventListener('click', () => c.grid.select(c.grid.pages.map((p) => p.src))); none.addEventListener('click', () => c.grid.select([]));
    range.addEventListener('change', () => { try { c.grid.select(parseRanges(range.value, c.view.numPages)); } catch (e) { toast(errText(e), true); } });
    c.side.append(el('h3', '', 'What to save'), mode, field('Pages to extract', range), hint);
    const run = runner(c.side, 'Save', async (r) => {
      const stem = stemOf(c.f.name);
      if (mode.value() === 'each') {
        const n = c.view.numPages; const w = Math.max(3, String(n).length); const outs = [];
        for (let i = 0; i < n; i++) { outs.push({ name: `${stem}-page-${pad(i + 1, w)}.pdf`, bytes: await opPick(c.f.bytes, [{ src: i }], c.f.name) }); r.tick((i + 1) / n); }
        const saved = await saveMany(outs, 'Choose where to save the pages'); if (saved) r.done(`Saved ${saved.length} files.`, saved);
      } else {
        const sel = c.grid.selected(); if (!sel.length) throw new Error('Tick at least one page first.');
        const out = await opPick(c.f.bytes, sel.map((s) => ({ src: s })), c.f.name);
        const saved = await saveOne(out, c.f.dir, `${stem}-pages.pdf`); if (saved) r.done(`Saved ${sel.length} page${sel.length > 1 ? 's' : ''} to ${baseName(saved)}.`, [saved]);
      }
    });
    c.onChange = () => { const n = c.grid.selected().length; hint.textContent = n ? `${n} page${n > 1 ? 's' : ''} ticked.` : 'Click pages to tick them.'; };
    c.onChange(); void run;
  });
}

function mountOrganize(tool) {
  singlePdfTool(tool, 'edit', (c) => {
    const hint = el('div', 'hint'); c.side.append(el('h3', '', 'Your changes'), hint); const rep = replaceOption(c.side);
    runner(c.side, 'Save PDF', async (r) => {
      const out = await opPick(c.f.bytes, c.grid.pages.map((p) => ({ src: p.src, rot: p.rot })), c.f.name); r.tick(0.9);
      const saved = await saveOut(out, c.f, `${stemOf(c.f.name)}-organised.pdf`, rep); if (saved) r.done(`Saved ${c.grid.pages.length} pages to ${baseName(saved)}.`, [saved]);
    });
    c.onChange = () => { hint.textContent = c.grid.changed() ? `${c.grid.pages.length} of ${c.view.numPages} pages kept. Drag a page to move it.` : 'Drag pages to reorder them, or use the buttons under each page.'; };
    c.onChange();
  });
}

function mountCompress(tool) {
  singlePdfTool(tool, null, (c) => {
    const level = radios('level', [['balanced', 'Balanced', 'Shrinks pictures. Text stays sharp and selectable.'], ['strong', 'Smallest file', 'Turns pages into pictures. Text can no longer be selected or searched.']], 'balanced');
    c.side.append(el('h3', '', 'Compression'), level); const rep = replaceOption(c.side);
    runner(c.side, 'Compress', async (r) => {
      const before = c.f.size; const out = await opCompress(c.f.bytes, level.value(), c.f.name, (v) => r.tick(v * 0.95));
      if (out.length >= c.f.size * 0.97) { r.done(`This PDF is already compact (${fmtSize(c.f.size)}). Nothing was saved.`); return; }
      const saved = await saveOut(out, c.f, `${stemOf(c.f.name)}-compressed.pdf`, rep);
      if (saved) r.done(`${fmtSize(before)} → ${fmtSize(out.length)}, ${Math.round((1 - out.length / before) * 100)}% smaller.`, [saved]);
    });
  });
}

function mountToImages(tool) {
  singlePdfTool(tool, null, (c) => {
    const fmt = select([['jpg', 'JPG (smaller files)'], ['png', 'PNG (sharper)']], 'jpg'); const dpi = select([['110', 'Standard'], ['200', 'High'], ['300', 'Print quality']], '110');
    c.side.append(el('h3', '', 'Pictures'), field('Format', fmt), field('Quality', dpi));
    runner(c.side, 'Save pictures', async (r) => {
      const outs = await opToImages(c.f.bytes, { format: fmt.value, dpi: Number(dpi.value) }, c.f.name, (v) => r.tick(v * 0.95));
      const saved = await saveMany(outs, 'Choose where to save the pictures'); if (saved) r.done(`Saved ${saved.length} picture${saved.length > 1 ? 's' : ''}.`, saved);
    });
  });
}

function mountWatermark(tool) {
  singlePdfTool(tool, null, (c) => {
    const text = el('input', 'field'); text.value = 'CONFIDENTIAL'; text.maxLength = 60; text.id = 'wmText';
    const size = select([['0.5', 'Small'], ['0.75', 'Medium'], ['1', 'Large']], '0.75'); const op = select([['0.12', 'Faint'], ['0.22', 'Normal'], ['0.4', 'Strong']], '0.22');
    c.side.append(el('h3', '', 'Stamp'), field('Text', text), field('Size', size), field('Strength', op)); const rep = replaceOption(c.side);
    runner(c.side, 'Add watermark', async (r) => {
      if (!text.value.trim()) throw new Error('Type the watermark text first.');
      const out = await opWatermark(c.f.bytes, { text: text.value.trim(), opacity: Number(op.value), sizeK: Number(size.value) }, c.f.name); r.tick(0.9);
      const saved = await saveOut(out, c.f, `${stemOf(c.f.name)}-watermarked.pdf`, rep); if (saved) r.done(`Saved ${baseName(saved)}.`, [saved]);
    });
  });
}

function mountPageNumbers(tool) {
  singlePdfTool(tool, null, (c) => {
    const pos = select([['bottom-center', 'Bottom centre'], ['bottom-right', 'Bottom right'], ['bottom-left', 'Bottom left'], ['top-right', 'Top right'], ['top-center', 'Top centre']], 'bottom-center');
    const style = select([['n', '1, 2, 3'], ['of', '1 / 10, 2 / 10']], 'n'); const start = el('input', 'field'); start.type = 'number'; start.min = '1'; start.value = '1';
    c.side.append(el('h3', '', 'Numbers'), field('Position', pos), field('Style', style), field('Start at', start)); const rep = replaceOption(c.side);
    runner(c.side, 'Add page numbers', async (r) => {
      const out = await opPageNumbers(c.f.bytes, { pos: pos.value, style: style.value, start: Math.max(1, Number(start.value) || 1) }, c.f.name); r.tick(0.9);
      const saved = await saveOut(out, c.f, `${stemOf(c.f.name)}-numbered.pdf`, rep); if (saved) r.done(`Saved ${baseName(saved)}.`, [saved]);
    });
  });
}

function mountFromImages(tool) {
  const { main, side } = workspace(tool); const files = []; const list = el('div', 'flist'); let render = () => {}; let run;
  const sync = () => { side.hidden = !files.length; zone.style.display = files.length ? 'none' : ''; more.style.display = files.length ? '' : 'none'; if (run) run.enable(files.length >= 1); };
  const add = async (paths) => { for (const p of paths) { const f = await guard(() => readFile(p)); if (f) { f.kind = 'pptx'; f.icon = 'i-image'; files.push(f); } } render(); sync(); };
  const zone = dropZone({ kind: 'image', multi: true, title: 'Choose pictures', hint: 'JPG, PNG, WebP, BMP or GIF', onPaths: add });
  const more = btn('btn addmore', 'Add more pictures', 'i-plus'); more.addEventListener('click', async () => { const p = await api.invoke('dlg:open', { kind: 'image', multi: true }); if (p.length) add(p); });
  main.append(zone, list, more);
  const size = select([['fit', 'Same size as each picture'], ['a4', 'A4 pages']], 'a4'); side.append(el('h3', '', 'Pages'), field('Page size', size), el('div', 'hint', 'One picture per page, from top to bottom.'));
  render = fileList(list, files, { onChange: sync });
  run = runner(side, 'Create PDF', async (r) => {
    const out = await opFromImages(files, { size: size.value }, (v) => r.tick(v * 0.9));
    const saved = await saveOne(out, files[0].dir, `${stemOf(files[0].name)}.pdf`); if (saved) r.done(`Saved ${files.length} picture${files.length > 1 ? 's' : ''} to ${baseName(saved)}.`, [saved]);
  });
  sync();
  const init = takeInitial(); if (init) add(init);
}

// Converts through the editor engine, one file after another, saving next to each original.
function convertTool(tool, { kind, title, hint, target, button }) {
  const { main, side } = workspace(tool); const files = []; const list = el('div', 'flist'); let render = () => {}; let run;
  const sync = () => { side.hidden = !files.length; zone.style.display = files.length ? 'none' : ''; more.style.display = files.length ? '' : 'none'; if (run) run.enable(files.length >= 1); };
  const add = async (paths) => {
    for (const p of paths) {
      const f = await guard(() => readFile(p)); if (!f) continue;
      const k = { docx: 'docx', doc: 'docx', odt: 'docx', rtf: 'docx', txt: 'docx', xlsx: 'xlsx', xls: 'xlsx', ods: 'xlsx', csv: 'xlsx', pptx: 'pptx', ppt: 'pptx', odp: 'pptx', pdf: 'pdf' }[extOf(f.name)] || 'docx';
      f.kind = k; f.icon = { docx: 'i-doc', xlsx: 'i-sheet', pptx: 'i-slides', pdf: 'i-pdf' }[k]; files.push(f);
    }
    render(); sync();
  };
  const zone = dropZone({ kind, multi: true, title, hint, onPaths: add });
  const more = btn('btn addmore', 'Add more files', 'i-plus'); more.addEventListener('click', async () => { const p = await api.invoke('dlg:open', { kind, multi: true }); if (p.length) add(p); });
  main.append(zone, list, more);
  side.append(el('h3', '', 'Where it goes'), el('div', 'hint', `Each ${target === 'pdf' ? 'PDF' : 'Word document'} is saved in the same folder as its original, with the same name.`));
  render = fileList(list, files, { onChange: sync });
  run = runner(side, button, async (r) => {
    const saved = []; let failed = 0;
    for (let i = 0; i < files.length; i++) {
      const f = files[i]; f.state = 'working'; f.stateText = 'Converting…'; render();
      try {
        const out = await Y.convertWithEngine(f.bytes, f.name, target);
        if (!out || out.length < 100) throw new Error('empty result');
        const path = await api.invoke('fs:uniquePath', f.dir, `${stemOf(f.name)}.${target}`);
        await api.invoke('fs:write', { file: path, data: out, force: true });
        saved.push(path); f.state = 'done'; f.stateText = 'Done';
      } catch (e) { failed++; f.state = 'failed'; f.stateText = 'Could not convert'; }
      render(); r.tick((i + 1) / files.length);
    }
    if (!saved.length) throw new Error('None of the files could be converted.');
    r.done(`Converted ${saved.length} file${saved.length > 1 ? 's' : ''}${failed ? `, ${failed} failed` : ''}.`, saved);
  });
  sync();
  const init = takeInitial(); if (init) add(init);
}
function mountOffice(tool) { convertTool(tool, { kind: 'convertible', title: 'Choose documents to convert', hint: 'Word, Excel or PowerPoint files', target: 'pdf', button: 'Convert to PDF' }); }
function mountToWord(tool) { convertTool(tool, { kind: 'pdf', title: 'Choose PDFs to convert', hint: 'or drop them here', target: 'docx', button: 'Convert to Word' }); }

function mountMetadata(tool) {
  const M = window.YOSMETA; const { main, side } = workspace(tool); const files = []; const list = el('div', 'flist'); const inside = el('div'); let render = () => {}; let run;
  const KIND = { pdf: ['pdf', 'i-pdf'], docx: ['docx', 'i-doc'], xlsx: ['xlsx', 'i-sheet'], pptx: ['pptx', 'i-slides'], jpg: ['pptx', 'i-image'], png: ['pptx', 'i-image'] };
  const fields = {}; const fieldBox = el('div'); fieldBox.style.display = 'flex'; fieldBox.style.flexDirection = 'column'; fieldBox.style.gap = '10px';
  for (const [k, label, ph] of [['title', 'Title', ''], ['author', 'Author', ''], ['subject', 'Subject', ''], ['keywords', 'Keywords', 'Separate with commas'], ['comments', 'Comments', ''], ['company', 'Company', 'Word, Excel and PowerPoint only'], ['copyright', 'Copyright', 'Pictures only']]) {
    const i = el('input', 'field'); i.placeholder = ph; i.maxLength = 500; i.dataset.meta = k; fields[k] = i; fieldBox.append(field(label, i));
  }
  const mode = radios('metamode', [['remove', 'Remove all details', 'Saves a clean copy with the hidden details taken out'], ['edit', 'Add or change details', 'Saves a copy with the details you type below']], 'remove');
  const hint = el('div', 'hint');
  const showInside = () => {
    inside.textContent = '';
    for (const f of files) {
      const box = el('div', 'meta-box'); const h = el('h4', '', f.name); h.append(el('span', '', f.metaErr ? '' : `${f.meta.length} detail${f.meta.length === 1 ? '' : 's'} found`)); box.append(h);
      if (f.metaErr) box.append(el('div', 'none', f.metaErr)); else if (!f.meta.length) box.append(el('div', 'none', 'No hidden details in this file.'));
      else { const dl = el('dl'); for (const [k, v] of f.meta) dl.append(el('dt', '', k), el('dd', '', v)); box.append(dl); }
      inside.append(box);
    }
  };
  const sync = () => {
    side.hidden = !files.length; zone.style.display = files.length ? 'none' : ''; more.style.display = files.length ? '' : 'none'; if (run) run.enable(files.length >= 1);
    const editing = mode.value() === 'edit'; fieldBox.style.display = editing ? 'flex' : 'none'; if (run) run.button.textContent = editing ? 'Save with these details' : 'Remove details';
    hint.textContent = editing ? 'Boxes you leave empty are not changed. In pictures, the old details are replaced by the ones you type.' : 'This clears the file properties. Names inside comments and tracked changes, and pictures placed inside a document, are not touched.';
    showInside();
  };
  const add = async (paths) => {
    for (const p of paths) {
      const f = await guard(() => readFile(p)); if (!f) continue;
      const t = M.typeOf(f.name); if (!KIND[t]) { toast(`${f.name} is not a supported file type`, true); continue; }
      [f.kind, f.icon] = KIND[t];
      try { f.meta = await M.read(f.bytes, f.name); } catch (e) { f.meta = []; f.metaErr = errText(e); }
      files.push(f);
    }
    if (files.length === 1 && !files[0].metaErr) for (const [k, v] of files[0].meta) { const key = { Title: 'title', Author: 'author', Subject: 'subject', Keywords: 'keywords', Comments: 'comments', Company: 'company', Copyright: 'copyright' }[k]; if (key && !fields[key].value) fields[key].value = v; }
    render(); sync();
  };
  const zone = dropZone({ kind: 'meta', multi: true, title: 'Choose files to check', hint: 'PDF, Word, Excel, PowerPoint, JPG or PNG', onPaths: add });
  const more = btn('btn addmore', 'Add more files', 'i-plus'); more.addEventListener('click', async () => { const p = await api.invoke('dlg:open', { kind: 'meta', multi: true }); if (p.length) add(p); });
  main.append(zone, list, more, inside);
  side.append(el('h3', '', 'What to do'), mode, fieldBox, hint); const rep = replaceOption(side);
  mode.addEventListener('change', sync);
  render = fileList(list, files, { onChange: sync, sub: (f) => `${fmtSize(f.size)}` });
  run = runner(side, 'Remove details', async (r) => {
    const editing = mode.value() === 'edit'; const vals = {}; for (const k of Object.keys(fields)) vals[k] = fields[k].value;
    if (editing && !Object.values(vals).some((v) => v.trim())) throw new Error('Fill in at least one box first.');
    const saved = []; let failed = 0; const over = rep.checked;
    if (over) { const open = files.find((f) => Y.isOpen(f.path)); if (open) throw new Error(`${open.name} is open in a tab. Close it first, then replace it.`); if (!(await confirmReplace(files.map((f) => f.name)))) return; }
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      try {
        const out = editing ? await M.edit(f.bytes, f.name, vals) : await M.remove(f.bytes, f.name);
        if (!out || out.length < 20) throw new Error('empty result');
        const path = over ? f.path : await api.invoke('fs:uniquePath', f.dir, `${stemOf(f.name)}-${editing ? 'details' : 'clean'}.${extOf(f.name)}`);
        await api.invoke('fs:write', { file: path, data: out, force: true });
        saved.push(path); f.state = 'done'; f.stateText = 'Done';
        if (over) { f.bytes = out; f.size = out.length; try { f.meta = await M.read(out, f.name); } catch (_) { f.meta = []; } }
      } catch (e) { failed++; f.state = 'failed'; f.stateText = errText(e).slice(0, 60); }
      r.tick((i + 1) / files.length);
    }
    render();
    if (!saved.length) throw new Error('None of the files could be processed.');
    if (over) { showInside(); Y.refreshHome(); }
    r.done(`${editing ? 'Saved details in' : 'Removed details from'} ${saved.length} file${saved.length > 1 ? 's' : ''}${failed ? `, ${failed} failed` : ''}. ${over ? 'The originals were replaced.' : `Each copy is next to its original, ending in "-${editing ? 'details' : 'clean'}".`}`, saved);
  });
  sync();
  const init = takeInitial(); if (init) add(init);
}

function pwField(id, placeholder) { const i = el('input', 'field'); i.type = 'password'; i.id = id; i.placeholder = placeholder || ''; i.autocomplete = 'new-password'; i.maxLength = 120; return i; }
function showToggle(inputs) {
  const l = el('label', 'chk'); const c = el('input'); c.type = 'checkbox'; l.append(c, document.createTextNode('Show password'));
  c.addEventListener('change', () => { for (const i of inputs) i.type = c.checked ? 'text' : 'password'; }); return l;
}
function mountProtect(tool) {
  singlePdfTool(tool, null, (c) => {
    const p1 = pwField('pw1', 'At least 4 characters'); const p2 = pwField('pw2', 'Type it again');
    c.side.append(el('h3', '', 'Password'), field('New password', p1), field('Confirm password', p2), showToggle([p1, p2]),
      el('div', 'hint', 'Anyone opening the file will need this password. If you forget it, the file cannot be opened, so keep it somewhere safe.'));
    runner(c.side, 'Protect PDF', async (r) => {
      if (p1.value.length < 4) throw new Error('Use a password of at least 4 characters.');
      if (p1.value !== p2.value) throw new Error('The two passwords do not match.');
      const doc = await loadLib(c.f.bytes, c.f.name);
      doc.encrypt({ userPassword: p1.value, ownerPassword: p1.value, permissions: { printing: 'highResolution', modifying: true, copying: true, annotating: true, fillingForms: true, contentAccessibility: true, documentAssembly: true } });
      const out = await doc.save({ useObjectStreams: false }); r.tick(0.9);
      const saved = await saveOne(out, c.f.dir, `${stemOf(c.f.name)}-protected.pdf`); if (saved) r.done(`Saved ${baseName(saved)}. It now needs the password to open.`, [saved]);
    });
  });
}
function mountUnlock(tool) {
  const { main, side } = workspace(tool);
  const start = () => { main.textContent = ''; side.hidden = true; side.textContent = ''; main.append(dropZone({ kind: 'pdf', title: 'Choose the locked PDF', onPaths: (p) => load(p[0]) })); };
  const openWith = async (bytes, password) => { try { return { doc: await PDFDocument.load(bytes, password == null ? undefined : { password }) }; } catch (e) { return { err: String((e && e.message) || e) }; } };
  async function load(path) {
    const f = await guard(() => readFile(path)); if (!f) return;
    main.textContent = ''; side.textContent = ''; side.hidden = false;
    const bar = el('div', 'bar2'); const other = btn('btn ghost', 'Choose another file'); other.addEventListener('click', start); const state = el('span', 'chip', 'Checking…');
    bar.append(el('b', '', f.name), state, el('div', 'grow'), other); main.append(bar);
    const plain = await openWith(f.bytes); const noUser = plain.doc ? null : await openWith(f.bytes, '');
    const pw = pwField('pwUnlock', 'Current password');
    if (plain.doc) { state.textContent = 'Not password-protected'; side.append(el('h3', '', 'Nothing to unlock'), el('div', 'hint', 'This PDF opens without a password.')); return; }
    if (!/encrypt|password/i.test(plain.err)) { state.textContent = 'Could not be read'; side.append(el('h3', '', 'Cannot open this file'), el('div', 'hint', 'This PDF could not be read. It may be damaged.')); return; }
    if (noUser && noUser.doc) { state.textContent = 'Restricted'; side.append(el('h3', '', 'Remove restrictions'), el('div', 'hint', 'This PDF opens without a password but has limits on printing or copying. Unlocking removes them.')); }
    else { state.textContent = 'Password-protected'; side.append(el('h3', '', 'Password'), field('Current password', pw), showToggle([pw]), el('div', 'hint', 'Type the password this PDF opens with. The unlocked copy will open without one.')); }
    runner(side, 'Unlock PDF', async (r) => {
      const res = noUser && noUser.doc ? noUser : await openWith(f.bytes, pw.value);
      if (!res.doc) throw new Error(/incorrect|password/i.test(res.err) ? 'That password is not right.' : 'This PDF could not be unlocked.');
      const out = await res.doc.save({ useObjectStreams: true }); r.tick(0.9);
      const saved = await saveOne(out, f.dir, `${stemOf(f.name)}-unlocked.pdf`); if (saved) r.done(`Saved ${baseName(saved)}. It opens without a password.`, [saved]);
    });
  }
  start();
  const init = takeInitial(); if (init) load(init[0]);
}

async function mountEdit() {
  const init = takeInitial(); if (init) { Y.openFile(init[0]); return; }
  const p = await api.invoke('dlg:open', { kind: 'pdf', multi: false, title: 'Choose a PDF to edit' });
  if (p.length) Y.openFile(p[0]);
}

function mountSign(tool) {
  const { main, side } = workspace(tool);
  const start = () => { main.textContent = ''; side.hidden = true; side.textContent = ''; main.append(dropZone({ kind: 'pdf', title: 'Choose the PDF to sign', onPaths: (p) => load(p[0]) })); };
  async function load(path) {
    const f = await guard(() => readFile(path)); if (!f) return;
    const view = await guard(() => loadView(f.bytes, f.name)); if (!view) return;
    main.textContent = ''; side.textContent = ''; side.hidden = false;
    let pageNo = 1; let sig = null; let box = null; let stageScale = 1;
    const bar = el('div', 'bar2'); const other = btn('btn ghost', 'Choose another file'); other.addEventListener('click', start);
    bar.append(el('b', '', f.name), el('span', 'chip', `${view.numPages} page${view.numPages > 1 ? 's' : ''}`), el('div', 'grow'), other);
    const wrap = el('div', 'stage-wrap'); const stage = el('div', 'pstage'); stage.id = 'signStage'; wrap.append(stage);
    const pager = el('div', 'pager'); const prev = el('button', 'btn icon ghost'); prev.append(icon('i-back')); const next = el('button', 'btn icon ghost'); const nx = icon('i-back'); nx.style.transform = 'rotate(180deg)'; next.append(nx);
    const lab = el('span'); pager.append(prev, lab, next); main.append(bar, wrap, pager);
    async function show() {
      const r = await renderPage(view, pageNo, { scale: 1.6, maxDim: 1000 }); stageScale = r.scale;
      r.canvas.style.width = `${Math.round(r.canvas.width / 1.25)}px`; stageScale = r.scale / 1.25;
      stage.textContent = ''; stage.append(r.canvas); lab.textContent = `Page ${pageNo} of ${view.numPages}`; prev.disabled = pageNo === 1; next.disabled = pageNo === view.numPages;
      if (box) stage.append(box);
    }
    prev.addEventListener('click', () => { pageNo--; show(); }); next.addEventListener('click', () => { pageNo++; show(); });

    // signature maker
    const tabs = el('div', 'seg sig-tabs'); const modes = [['draw', 'Draw'], ['type', 'Type'], ['image', 'Picture']]; let mode = 'draw';
    const padC = el('canvas', 'sigpad'); padC.id = 'sigPad'; const typed = el('input', 'field'); typed.placeholder = 'Type your name'; typed.id = 'sigType'; typed.maxLength = 40;
    const pick = btn('btn', 'Choose a picture'); const ink = select([['#14121c', 'Black ink'], ['#1b3a8c', 'Blue ink']], '#14121c'); const clear = btn('btn ghost', 'Clear');
    const panes = { draw: el('div'), type: el('div'), image: el('div') };
    panes.draw.append(padC); panes.type.append(typed); panes.image.append(pick, el('div', 'hint', 'A photo of your signature on white paper works well.'));
    let imgCanvas = null; let hasInk = false; let g = null;
    const setMode = (m) => { mode = m; for (const b of tabs.children) b.classList.toggle('on', b.dataset.m === m); for (const k of Object.keys(panes)) panes[k].hidden = k !== m; if (m === 'draw') fit(); };
    for (const [m, t] of modes) { const b = el('button', '', t); b.dataset.m = m; b.addEventListener('click', () => setMode(m)); tabs.append(b); }
    function fit() { const r = padC.getBoundingClientRect(); if (!r.width) return; if (padC.width !== Math.round(r.width * 2)) { padC.width = Math.round(r.width * 2); padC.height = Math.round(r.height * 2); hasInk = false; } g = padC.getContext('2d'); g.lineCap = 'round'; g.lineJoin = 'round'; g.lineWidth = 4.2; g.strokeStyle = ink.value; }
    let drawing = false; let last = null;
    const pt = (e) => { const r = padC.getBoundingClientRect(); return { x: (e.clientX - r.left) * padC.width / r.width, y: (e.clientY - r.top) * padC.height / r.height }; };
    padC.addEventListener('pointerdown', (e) => { fit(); drawing = true; last = pt(e); padC.setPointerCapture(e.pointerId); g.beginPath(); g.arc(last.x, last.y, 1.6, 0, 7); g.fillStyle = ink.value; g.fill(); hasInk = true; });
    padC.addEventListener('pointermove', (e) => { if (!drawing) return; const p = pt(e); g.strokeStyle = ink.value; g.beginPath(); g.moveTo(last.x, last.y); g.lineTo(p.x, p.y); g.stroke(); last = p; });
    const stop = () => { drawing = false; }; padC.addEventListener('pointerup', stop); padC.addEventListener('pointercancel', stop);
    clear.addEventListener('click', () => { if (g) g.clearRect(0, 0, padC.width, padC.height); hasInk = false; typed.value = ''; imgCanvas = null; });
    pick.addEventListener('click', async () => {
      const p = await api.invoke('dlg:open', { kind: 'image', multi: false, title: 'Choose a picture of your signature' }); if (!p.length) return;
      const im = await guard(() => readFile(p[0])); if (!im) return;
      const bmp = await createImageBitmap(new Blob([im.bytes])); const k = Math.min(1, 900 / Math.max(bmp.width, bmp.height));
      const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k); const x = c.getContext('2d'); x.drawImage(bmp, 0, 0, c.width, c.height);
      const d = x.getImageData(0, 0, c.width, c.height); // paper turns transparent, ink stays
      for (let i = 0; i < d.data.length; i += 4) { const lum = (d.data[i] * 3 + d.data[i + 1] * 6 + d.data[i + 2]) / 10; d.data[i + 3] = lum > 200 ? 0 : lum > 140 ? Math.round(255 * (200 - lum) / 60) : d.data[i + 3]; }
      x.putImageData(d, 0, 0); imgCanvas = c; toast('Picture ready. Now add it to the page.');
    });
    function trimmed(src) {
      const x = src.getContext('2d'); const d = x.getImageData(0, 0, src.width, src.height).data; let l = src.width; let t = src.height; let r = -1; let b = -1;
      for (let yy = 0; yy < src.height; yy++) for (let xx = 0; xx < src.width; xx++) if (d[(yy * src.width + xx) * 4 + 3] > 8) { if (xx < l) l = xx; if (xx > r) r = xx; if (yy < t) t = yy; if (yy > b) b = yy; }
      if (r < 0) return null;
      const pd = 6; l = Math.max(0, l - pd); t = Math.max(0, t - pd); r = Math.min(src.width - 1, r + pd); b = Math.min(src.height - 1, b + pd);
      const c = document.createElement('canvas'); c.width = r - l + 1; c.height = b - t + 1; c.getContext('2d').drawImage(src, l, t, c.width, c.height, 0, 0, c.width, c.height); return c;
    }
    async function makeSignature() {
      let c = null;
      if (mode === 'draw') c = hasInk ? trimmed(padC) : null;
      else if (mode === 'type') { const v = typed.value.trim(); if (v) { await document.fonts.load('700 120px Caveat'); c = trimmed(textPng(v, { px: 120, color: ink.value, font: '700 120px Caveat, cursive' }).canvas); } }
      else c = imgCanvas ? trimmed(imgCanvas) : null;
      if (!c) throw new Error(mode === 'draw' ? 'Draw your signature in the box first.' : mode === 'type' ? 'Type your name first.' : 'Choose a picture first.');
      return { png: await blobBytes(await canvasBlob(c, 'image/png')), url: c.toDataURL('image/png'), ratio: c.height / c.width };
    }
    const placeBtn = btn('btn', 'Add to this page'); placeBtn.dataset.act = 'place';
    placeBtn.addEventListener('click', async () => {
      const s = await guard(makeSignature); if (!s) return; sig = s;
      const cv = stage.querySelector('canvas'); const cw = cv.clientWidth; const ch = cv.clientHeight; const w = Math.min(cw * 0.32, 220); const h = w * sig.ratio;
      if (box) box.remove();
      box = el('div', 'sigbox'); box.id = 'sigBox'; const im = el('img'); im.src = sig.url; im.alt = 'Your signature'; const hd = el('span', 'h'); box.append(im, hd);
      box.style.width = `${w}px`; box.style.height = `${h}px`; box.style.left = `${cw * 0.58}px`; box.style.top = `${Math.min(ch - h - 10, ch * 0.78)}px`; stage.append(box);
      let act = null;
      const down = (e, kind) => { e.preventDefault(); e.stopPropagation(); act = { kind, sx: e.clientX, sy: e.clientY, l: box.offsetLeft, t: box.offsetTop, w: box.offsetWidth }; box.setPointerCapture(e.pointerId); };
      box.addEventListener('pointerdown', (e) => down(e, e.target === hd ? 'size' : 'move'));
      box.addEventListener('pointermove', (e) => {
        if (!act) return; const dx = e.clientX - act.sx; const dy = e.clientY - act.sy; const W = cv.clientWidth; const H = cv.clientHeight;
        if (act.kind === 'move') { box.style.left = `${Math.max(0, Math.min(W - box.offsetWidth, act.l + dx))}px`; box.style.top = `${Math.max(0, Math.min(H - box.offsetHeight, act.t + dy))}px`; }
        else { const nw = Math.max(40, Math.min(W - act.l, act.w + dx)); box.style.width = `${nw}px`; box.style.height = `${nw * sig.ratio}px`; }
      });
      box.addEventListener('pointerup', () => { act = null; });
      hintP.textContent = 'Drag the signature into place. Pull the dot to resize it.';
    });
    const hintP = el('div', 'hint', 'Make your signature, then add it to the page shown.');
    side.append(el('h3', '', 'Your signature'), tabs, panes.draw, panes.type, panes.image, field('Ink', ink), clear, placeBtn, hintP); const rep = replaceOption(side);
    runner(side, 'Save signed PDF', async (r) => {
      if (!box || !sig) throw new Error('Add your signature to the page first.');
      const out = await opSign(f.bytes, { pageIndex: pageNo - 1, png: sig.png, vx: box.offsetLeft / stageScale, vy: box.offsetTop / stageScale, vw: box.offsetWidth / stageScale, vh: box.offsetHeight / stageScale }, f.name); r.tick(0.9);
      const saved = await saveOut(out, f, `${stemOf(f.name)}-signed.pdf`, rep); if (saved) r.done(`Saved ${baseName(saved)}.`, [saved]);
    });
    side.append(el('div', 'hint', 'This places a picture of your signature. It is not a certificate-based digital signature.'));
    await show(); setMode('draw');
  }
  start();
  const init = takeInitial(); if (init) load(init[0]);
}

// ---------------------------------------------------------------- boot
function renderTools() {
  const grid = $('pdfTools'); grid.textContent = '';
  for (const t of TOOLS) {
    const c = el('button', 'tool-card'); c.dataset.tool = t.id; const i = el('i'); i.append(icon(t.icon)); const tx = el('div'); tx.append(el('b', '', t.name), el('span', '', t.desc));
    c.append(i, tx); c.addEventListener('click', () => t.mount(t)); grid.append(c);
  }
}
renderTools();
function openTool(id, paths) {
  const t = TOOLS.find((x) => x.id === id); if (!t) return;
  Y.activateHome(); Y.showPage('pdf'); pendingPaths = paths && paths.length ? paths : null;
  try { t.mount(t); } finally { pendingPaths = null; }
}
// Used by the build script that makes template previews, and handy for support.
window.YOSPDF = { openTool, ops: { opMerge, opPick, opCompress, opToImages, opFromImages, opWatermark, opPageNumbers, opSign }, loadView, renderPage, parseRanges };
