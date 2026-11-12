/* YourOfficeSuite: read, remove and write file details (metadata) for PDF, Word, Excel, PowerPoint, JPG and PNG.
   Everything runs on this computer. Pictures are never re-encoded, so their quality does not change. */
(() => {
  'use strict';
  const { PDFDocument, PDFName, PDFDict } = window.PDFLib;
  const fflate = window.fflate;
  const dec = new TextDecoder(); const enc = new TextEncoder();
  const utf8 = new TextDecoder('utf-8', { fatal: true });
  const text8 = (u8) => { try { return utf8.decode(u8); } catch (_) { return latin1(u8); } };
  const latin1 = (u8) => { let s = ''; for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]); return s; };
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const concat = (parts) => { const n = parts.reduce((a, p) => a + p.length, 0); const out = new Uint8Array(n); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };
  const keywordsOf = (s) => String(s || '').split(/[,;]/).map((k) => k.trim()).filter(Boolean);
  const typeOf = (name) => { const e = name.split('.').pop().toLowerCase(); return e === 'jpeg' ? 'jpg' : e; };

  // ------------------------------------------------------------------ PDF
  async function pdfLoad(bytes, name) {
    let doc; let pages = 0;
    try { doc = await PDFDocument.load(bytes, { updateMetadata: false }); pages = doc.getPageCount(); } catch (e) {
      if (/encrypt/i.test(String(e && e.message))) throw new Error(`${name} is password-protected. Remove the password first.`);
      throw new Error(`${name} could not be read. It may be damaged.`);
    }
    if (!pages) throw new Error(`${name} could not be read. It may be damaged.`);
    return doc;
  }
  const xmpRefs = (doc) => doc.context.enumerateIndirectObjects().filter(([, o]) => o && o.dict instanceof PDFDict && o.dict.lookup(PDFName.of('Type')) === PDFName.of('Metadata'));
  async function pdfRead(bytes, name) {
    const doc = await pdfLoad(bytes, name); const out = [];
    const add = (k, v) => { if (v != null && String(v).trim()) out.push([k, v instanceof Date ? v.toLocaleString() : String(v)]); };
    const safe = (fn) => { try { return fn(); } catch (_) { return undefined; } };
    add('Title', safe(() => doc.getTitle())); add('Author', safe(() => doc.getAuthor())); add('Subject', safe(() => doc.getSubject())); add('Keywords', safe(() => doc.getKeywords()));
    add('Created with', safe(() => doc.getCreator())); add('PDF maker', safe(() => doc.getProducer())); add('Created', safe(() => doc.getCreationDate())); add('Changed', safe(() => doc.getModificationDate()));
    if (xmpRefs(doc).length) out.push(['Extra details block (XMP)', 'Present']);
    return out;
  }
  function pdfStripXmp(doc) {
    for (const [ref] of xmpRefs(doc)) doc.context.delete(ref);
    doc.catalog.delete(PDFName.of('Metadata'));
  }
  async function pdfRemove(bytes, name) {
    const doc = await pdfLoad(bytes, name); const info = doc.getInfoDict();
    for (const k of info.keys()) info.delete(k);
    pdfStripXmp(doc);
    return doc.save({ useObjectStreams: true });
  }
  async function pdfEdit(bytes, f, name) {
    const doc = await pdfLoad(bytes, name);
    if (f.title) doc.setTitle(f.title); if (f.author) doc.setAuthor(f.author);
    if (f.subject || f.comments) doc.setSubject(f.subject || f.comments);
    if (f.keywords) doc.setKeywords(keywordsOf(f.keywords));
    doc.setModificationDate(new Date());
    pdfStripXmp(doc); // the old block would otherwise still show the old values
    return doc.save({ useObjectStreams: true });
  }

  // ------------------------------------------------------------------ Word, Excel, PowerPoint
  const NS = { cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', dc: 'http://purl.org/dc/elements/1.1/', dcterms: 'http://purl.org/dc/terms/',
    ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties', vt: 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes' };
  const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const CORE_EMPTY = `${XML_HEAD}<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>`;
  const APP_EMPTY = `${XML_HEAD}<Properties xmlns="${NS.ep}" xmlns:vt="${NS.vt}"></Properties>`;
  const CUSTOM_EMPTY = `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="${NS.vt}"></Properties>`;
  const CORE_LABEL = { title: 'Title', creator: 'Author', subject: 'Subject', keywords: 'Keywords', description: 'Comments', lastModifiedBy: 'Last saved by', created: 'Created', modified: 'Changed',
    revision: 'Revision', category: 'Category', contentStatus: 'Status', lastPrinted: 'Last printed', language: 'Language', identifier: 'Identifier', version: 'Version' };
  const APP_LABEL = { Company: 'Company', Manager: 'Manager', Application: 'Made with', AppVersion: 'App version', Template: 'Template', TotalTime: 'Editing time (minutes)', HyperlinkBase: 'Link base' };
  const parseXml = (u8) => { const d = new DOMParser().parseFromString(dec.decode(u8), 'application/xml'); if (d.querySelector('parsererror')) throw new Error('bad xml'); return d; };
  const toXml = (doc) => enc.encode(XML_HEAD + new XMLSerializer().serializeToString(doc.documentElement));
  function unzip(bytes, name, filter) {
    try { return fflate.unzipSync(bytes, filter ? { filter } : undefined); } catch (_) { throw new Error(`${name} could not be read. It may be damaged or password-protected.`); }
  }
  function ooxmlRead(bytes, name) {
    const z = unzip(bytes, name, (f) => /^docProps\/(core|app|custom)\.xml$/.test(f.name)); const out = [];
    try {
      if (z['docProps/core.xml']) for (const e of parseXml(z['docProps/core.xml']).documentElement.children) { const v = e.textContent.trim(); if (v) out.push([CORE_LABEL[e.localName] || e.localName, v]); }
      if (z['docProps/app.xml']) for (const e of parseXml(z['docProps/app.xml']).documentElement.children) { const v = e.textContent.trim(); if (v && APP_LABEL[e.localName] && !e.children.length) out.push([APP_LABEL[e.localName], v]); }
      if (z['docProps/custom.xml']) for (const e of parseXml(z['docProps/custom.xml']).documentElement.children) { const v = e.textContent.trim(); if (v) out.push([`Custom: ${e.getAttribute('name') || 'property'}`, v]); }
    } catch (_) { out.push(['Details', 'Present, but could not be read']); }
    return out;
  }
  function ensurePart(files, part, contentType, relType) {
    if (files[part]) return;
    const ct = '[Content_Types].xml'; const rels = '_rels/.rels';
    if (!files[ct] || !files[rels]) throw new Error('This is not a valid Office file.');
    let c = dec.decode(files[ct]);
    if (!c.includes(`PartName="/${part}"`)) c = c.replace('</Types>', `<Override PartName="/${part}" ContentType="${contentType}"/></Types>`);
    files[ct] = enc.encode(c);
    let r = dec.decode(files[rels]);
    if (!r.includes(`Target="${part}"`)) { let id = 'rIdYos1'; let n = 1; while (r.includes(`Id="${id}"`)) id = `rIdYos${++n}`; r = r.replace('</Relationships>', `<Relationship Id="${id}" Type="${relType}" Target="${part}"/></Relationships>`); }
    files[rels] = enc.encode(r);
  }
  function setChild(doc, ns, qname, value) {
    const local = qname.split(':').pop(); let e = [...doc.documentElement.children].find((c) => c.localName === local && (c.namespaceURI === ns));
    if (!e) { e = doc.createElementNS(ns, qname); doc.documentElement.append(e); }
    e.textContent = value;
  }
  function zipUp(files) {
    const input = {}; for (const k of Object.keys(files)) input[k] = [files[k], { level: /\.(png|jpe?g|gif|zip|mp4|mp3)$/i.test(k) ? 0 : 6 }];
    return fflate.zipSync(input);
  }
  function ooxmlRemove(bytes, name) {
    const files = unzip(bytes, name); if (!files['[Content_Types].xml']) throw new Error(`${name} is not a valid Office file.`);
    if (files['docProps/core.xml']) files['docProps/core.xml'] = enc.encode(CORE_EMPTY);
    if (files['docProps/app.xml']) files['docProps/app.xml'] = enc.encode(APP_EMPTY);
    if (files['docProps/custom.xml']) files['docProps/custom.xml'] = enc.encode(CUSTOM_EMPTY);
    return zipUp(files);
  }
  function ooxmlEdit(bytes, f, name) {
    const files = unzip(bytes, name); if (!files['[Content_Types].xml']) throw new Error(`${name} is not a valid Office file.`);
    ensurePart(files, 'docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties');
    let core; try { core = parseXml(files['docProps/core.xml'] || enc.encode(CORE_EMPTY)); } catch (_) { core = parseXml(enc.encode(CORE_EMPTY)); }
    if (f.title) setChild(core, NS.dc, 'dc:title', f.title); if (f.author) setChild(core, NS.dc, 'dc:creator', f.author); if (f.subject) setChild(core, NS.dc, 'dc:subject', f.subject);
    if (f.keywords) setChild(core, NS.cp, 'cp:keywords', keywordsOf(f.keywords).join(', ')); if (f.comments) setChild(core, NS.dc, 'dc:description', f.comments);
    files['docProps/core.xml'] = toXml(core);
    if (f.company) {
      ensurePart(files, 'docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties');
      let app; try { app = parseXml(files['docProps/app.xml'] || enc.encode(APP_EMPTY)); } catch (_) { app = parseXml(enc.encode(APP_EMPTY)); }
      setChild(app, NS.ep, 'Company', f.company); files['docProps/app.xml'] = toXml(app);
    }
    return zipUp(files);
  }

  // ------------------------------------------------------------------ JPG
  function jpegSegments(b, name) {
    if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) throw new Error(`${name} is not a valid JPG picture.`);
    const segs = []; let p = 2;
    while (p + 4 <= b.length) {
      if (b[p] !== 0xff) throw new Error(`${name} could not be read. It may be damaged.`);
      const m = b[p + 1];
      if (m === 0xff) { p++; continue; }
      if (m === 0xda || m === 0xd9) { segs.push({ m: 0, start: p, end: b.length }); return segs; } // picture data: keep as is
      const len = (b[p + 2] << 8) | b[p + 3]; if (len < 2 || p + 2 + len > b.length) throw new Error(`${name} could not be read. It may be damaged.`);
      segs.push({ m, start: p, end: p + 2 + len }); p += 2 + len;
    }
    throw new Error(`${name} could not be read. It may be damaged.`);
  }
  const startsWith = (b, off, text) => { for (let i = 0; i < text.length; i++) if (b[off + i] !== text.charCodeAt(i)) return false; return true; };
  function exifParse(b, start, end) {
    const out = {}; const t = start; if (end - t < 14) return out;
    const le = b[t] === 0x49; const u16 = (o) => (le ? b[o] | (b[o + 1] << 8) : (b[o] << 8) | b[o + 1]); const u32 = (o) => (le ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0);
    const ifd = (off, names) => {
      const p = t + off; if (p + 2 > end) return; const n = u16(p);
      for (let i = 0; i < n; i++) {
        const e = p + 2 + i * 12; if (e + 12 > end) return; const tag = u16(e); const type = u16(e + 2); const count = u32(e + 4);
        if (!(tag in names)) continue;
        if (type === 2) { const o = count <= 4 ? e + 8 : t + u32(e + 8); if (o + count <= end) out[names[tag]] = text8(b.subarray(o, o + count).filter((x) => x !== 0)).trim(); }
        else if (type === 3) out[names[tag]] = u16(e + 8); else if (type === 4) out[names[tag]] = u32(e + 8);
      }
    };
    try {
      ifd(u32(t + 4), { 0x010e: 'desc', 0x010f: 'make', 0x0110: 'model', 0x0112: 'orientation', 0x0131: 'software', 0x0132: 'date', 0x013b: 'artist', 0x8298: 'copyright', 0x8825: 'gps', 0x8769: 'exifIfd' });
      if (out.exifIfd) ifd(out.exifIfd, { 0x9003: 'taken' });
    } catch (_) { /* partial data is fine */ }
    return out;
  }
  function jpegScan(b, name) {
    const segs = jpegSegments(b, name); let exif = null; const blocks = [];
    for (const s of segs) {
      const d = s.start + 4;
      if (s.m === 0xe1 && startsWith(b, d, 'Exif\0\0')) { exif = exifParse(b, d + 6, s.end); blocks.push('camera data (EXIF)'); }
      else if (s.m === 0xe1 && startsWith(b, d, 'http://ns.adobe.com/xap/')) blocks.push('XMP');
      else if (s.m === 0xed) blocks.push('IPTC'); else if (s.m === 0xfe) blocks.push('comment');
    }
    return { segs, exif: exif || {}, blocks };
  }
  function jpegRead(b, name) {
    const { exif, blocks, segs } = jpegScan(b, name); const out = [];
    const add = (k, v) => { if (v != null && String(v).trim()) out.push([k, String(v)]); };
    add('Title', exif.desc); add('Author', exif.artist); add('Copyright', exif.copyright); add('Camera', [exif.make, exif.model].filter(Boolean).join(' ')); add('Date taken', exif.taken || exif.date); add('Software', exif.software);
    if (exif.gps) out.push(['Location (GPS)', 'Present']);
    const xmp = segs.find((s) => s.m === 0xe1 && startsWith(b, s.start + 4, 'http://ns.adobe.com/xap/'));
    if (xmp) for (const [k, v] of xmpFields(dec.decode(b.subarray(xmp.start + 4, xmp.end)))) if (!out.some((r) => r[0] === k)) out.push([k, v]);
    if (blocks.length) out.push(['Hidden blocks', [...new Set(blocks)].join(', ')]);
    return out;
  }
  function xmpFields(xml) {
    const out = []; const grab = (tag, label) => { const m = xml.match(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`)); if (!m) return; const li = [...m[1].matchAll(/<rdf:li[^>]*>([\s\S]*?)<\/rdf:li>/g)].map((x) => x[1].trim()).filter(Boolean); const v = li.length ? li.join(', ') : m[1].replace(/<[^>]+>/g, '').trim(); if (v) out.push([label, v.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')]); };
    grab('dc:title', 'Title'); grab('dc:creator', 'Author'); grab('dc:description', 'Comments'); grab('dc:subject', 'Keywords'); grab('dc:rights', 'Copyright');
    return out;
  }
  function exifBuild({ orientation, desc, artist, copyright }) {
    const ents = [];
    const ascii = (tag, text) => { if (text) ents.push({ tag, type: 2, data: concat([enc.encode(text), new Uint8Array(1)]) }); };
    ascii(0x010e, desc); if (orientation > 1 && orientation <= 8) ents.push({ tag: 0x0112, type: 3, value: orientation }); ascii(0x013b, artist); ascii(0x8298, copyright);
    if (!ents.length) return null;
    ents.sort((a, b) => a.tag - b.tag);
    const ifdSize = 2 + ents.length * 12 + 4; let dataOff = 8 + ifdSize; const extra = [];
    const body = new Uint8Array(8 + ifdSize); const v = new DataView(body.buffer);
    body.set([0x49, 0x49, 0x2a, 0x00]); v.setUint32(4, 8, true); v.setUint16(8, ents.length, true);
    ents.forEach((e, i) => {
      const o = 10 + i * 12; v.setUint16(o, e.tag, true); v.setUint16(o + 2, e.type, true);
      if (e.type === 3) { v.setUint32(o + 4, 1, true); v.setUint16(o + 8, e.value, true); }
      else { v.setUint32(o + 4, e.data.length, true); if (e.data.length <= 4) body.set(e.data, o + 8); else { v.setUint32(o + 8, dataOff, true); extra.push(e.data); dataOff += e.data.length; if (e.data.length % 2) { extra.push(new Uint8Array(1)); dataOff += 1; } } }
    });
    return concat([enc.encode('Exif\0\0'), body, ...extra]);
  }
  const jpegSeg = (marker, payload) => { const n = payload.length + 2; if (n > 65535) throw new Error('That text is too long to store in a picture.'); return concat([new Uint8Array([0xff, marker, n >> 8, n & 255]), payload]); };
  function xmpPacket(f) {
    const alt = (tag, v) => (v ? `<${tag}><rdf:Alt><rdf:li xml:lang="x-default">${esc(v)}</rdf:li></rdf:Alt></${tag}>` : '');
    const kw = keywordsOf(f.keywords);
    const body = alt('dc:title', f.title) + (f.author ? `<dc:creator><rdf:Seq><rdf:li>${esc(f.author)}</rdf:li></rdf:Seq></dc:creator>` : '') + alt('dc:description', f.comments || f.subject)
      + (kw.length ? `<dc:subject><rdf:Bag>${kw.map((k) => `<rdf:li>${esc(k)}</rdf:li>`).join('')}</rdf:Bag></dc:subject>` : '') + alt('dc:rights', f.copyright);
    if (!body) return null;
    return `<?xpacket begin="﻿" id="W5M0MpCehiHzreSzNTczkc9d"?><x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/">${body}</rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end="w"?>`;
  }
  function jpegWrite(b, name, f) {
    const { segs, exif } = jpegScan(b, name); const parts = [b.subarray(0, 2)]; const kept = [];
    for (const s of segs) {
      const d = s.start + 4; const isApp = s.m >= 0xe0 && s.m <= 0xef;
      const keep = s.m === 0 || (!isApp && s.m !== 0xfe) || s.m === 0xe0 || s.m === 0xee || (s.m === 0xe2 && startsWith(b, d, 'ICC_PROFILE'));
      if (keep) kept.push(s);
    }
    let i = 0; if (kept[0] && kept[0].m === 0xe0) { parts.push(b.subarray(kept[0].start, kept[0].end)); i = 1; } // JFIF header stays first
    // The turn-the-picture hint is kept so a cleaned photo does not end up sideways.
    const ex = exifBuild({ orientation: exif.orientation, desc: f && f.title, artist: f && f.author, copyright: f && f.copyright }); if (ex) parts.push(jpegSeg(0xe1, ex));
    const xmp = f ? xmpPacket(f) : null; if (xmp) parts.push(jpegSeg(0xe1, concat([enc.encode('http://ns.adobe.com/xap/1.0/\0'), enc.encode(xmp)])));
    for (; i < kept.length; i++) parts.push(b.subarray(kept[i].start, kept[i].end));
    return concat(parts);
  }

  // ------------------------------------------------------------------ PNG
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (u8) => { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const PNG_SIG = [137, 80, 78, 71, 13, 10, 26, 10];
  const PNG_META = ['tEXt', 'zTXt', 'iTXt', 'eXIf', 'tIME'];
  function pngChunks(b, name) {
    if (b.length < 20 || PNG_SIG.some((v, i) => b[i] !== v)) throw new Error(`${name} is not a valid PNG picture.`);
    const out = []; let p = 8; const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
    while (p + 12 <= b.length) {
      const len = v.getUint32(p); const type = latin1(b.subarray(p + 4, p + 8)); const end = p + 12 + len;
      if (end > b.length) throw new Error(`${name} could not be read. It may be damaged.`);
      out.push({ type, start: p, end, data: b.subarray(p + 8, p + 8 + len) }); p = end;
      if (type === 'IEND') break;
    }
    if (!out.length || out[0].type !== 'IHDR') throw new Error(`${name} could not be read. It may be damaged.`);
    return out;
  }
  function pngChunk(type, data) {
    const body = concat([enc.encode(type), data]); const out = new Uint8Array(12 + data.length); const v = new DataView(out.buffer);
    v.setUint32(0, data.length); out.set(body, 4); v.setUint32(8 + data.length, crc32(body)); return out;
  }
  const PNG_LABEL = { Title: 'Title', Author: 'Author', Description: 'Comments', Comment: 'Comments', Copyright: 'Copyright', Software: 'Software', 'Creation Time': 'Created', Source: 'Source', Keywords: 'Keywords', Subject: 'Subject' };
  function pngRead(b, name) {
    const out = []; const other = [];
    for (const c of pngChunks(b, name)) {
      if (c.type === 'tEXt') { const z = c.data.indexOf(0); if (z > 0) out.push([PNG_LABEL[latin1(c.data.subarray(0, z))] || latin1(c.data.subarray(0, z)), latin1(c.data.subarray(z + 1)).slice(0, 300)]); }
      else if (c.type === 'iTXt') {
        const z = c.data.indexOf(0); if (z <= 0) continue; const key = latin1(c.data.subarray(0, z)); const comp = c.data[z + 1]; const l = c.data.indexOf(0, z + 3); const t = l >= 0 ? c.data.indexOf(0, l + 1) : -1;
        if (key === 'XML:com.adobe.xmp') { other.push('XMP'); if (!comp && t >= 0) for (const r of xmpFields(dec.decode(c.data.subarray(t + 1)))) out.push(r); }
        else if (!comp && t >= 0) out.push([PNG_LABEL[key] || key, dec.decode(c.data.subarray(t + 1)).slice(0, 300)]); else other.push('compressed text');
      } else if (c.type === 'zTXt') other.push('compressed text'); else if (c.type === 'eXIf') other.push('camera data (EXIF)'); else if (c.type === 'tIME') other.push('time stamp');
    }
    if (other.length) out.push(['Hidden blocks', [...new Set(other)].join(', ')]);
    return out;
  }
  function pngWrite(b, name, f) {
    const chunks = pngChunks(b, name); const parts = [b.subarray(0, 8)];
    const text = (key, value) => {
      if (!value) return null;
      const plain = !/[^\x20-\x7e\xa0-\xff]/.test(value);
      return plain ? pngChunk('tEXt', concat([enc.encode(key), new Uint8Array(1), Uint8Array.from(value, (ch) => ch.charCodeAt(0))])) : pngChunk('iTXt', concat([enc.encode(key), new Uint8Array(5), enc.encode(value)]));
    };
    for (const c of chunks) {
      if (PNG_META.includes(c.type)) continue;
      parts.push(b.subarray(c.start, c.end));
      if (c.type === 'IHDR' && f) for (const ch of [text('Title', f.title), text('Author', f.author), text('Description', f.comments || f.subject), text('Keywords', keywordsOf(f.keywords).join(', ')), text('Copyright', f.copyright)]) if (ch) parts.push(ch);
    }
    return concat(parts);
  }

  // ------------------------------------------------------------------ public
  const OFFICE = ['docx', 'xlsx', 'pptx'];
  async function read(bytes, name) {
    const t = typeOf(name);
    if (t === 'pdf') return pdfRead(bytes, name); if (OFFICE.includes(t)) return ooxmlRead(bytes, name);
    if (t === 'jpg') return jpegRead(bytes, name); if (t === 'png') return pngRead(bytes, name);
    throw new Error(`${name} is not a supported file type.`);
  }
  async function remove(bytes, name) {
    const t = typeOf(name);
    if (t === 'pdf') return pdfRemove(bytes, name); if (OFFICE.includes(t)) return ooxmlRemove(bytes, name);
    if (t === 'jpg') return jpegWrite(bytes, name, null); if (t === 'png') return pngWrite(bytes, name, null);
    throw new Error(`${name} is not a supported file type.`);
  }
  async function edit(bytes, name, fields) {
    const t = typeOf(name); const f = {}; for (const k of Object.keys(fields || {})) { const v = String(fields[k] == null ? '' : fields[k]).trim(); if (v) f[k] = v.slice(0, 2000); }
    if (!Object.keys(f).length) throw new Error('Fill in at least one box first.');
    if (t === 'pdf') return pdfEdit(bytes, f, name); if (OFFICE.includes(t)) return ooxmlEdit(bytes, f, name);
    if (t === 'jpg') return jpegWrite(bytes, name, f); if (t === 'png') return pngWrite(bytes, name, f);
    throw new Error(`${name} is not a supported file type.`);
  }
  window.YOSMETA = { read, remove, edit, typeOf, TYPES: ['pdf', ...OFFICE, 'jpg', 'jpeg', 'png'] };
})();
