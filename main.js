// YourOfficeSuite - main process.
// Serves the bundled editor engine from a local-only HTTP server and gives the
// shell window file access to the user's OneDrive folder.
const { app, BrowserWindow, ipcMain, dialog, shell, Menu, session, nativeTheme } = require('electron');
const childProcess = require('child_process');
const http = require('http');
const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const os = require('os');

const ENGINE_DIR = app.isPackaged ? path.join(process.resourcesPath, 'engine') : path.join(__dirname, 'engine');
const SHELL_DIR = path.join(__dirname, 'shell');
const PREFERRED_PORT = 38917;
const EDIT_EXT = ['docx', 'xlsx', 'pptx', 'csv', 'pdf'];
const IMAGE_EXT = ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif'];
const OPEN_EXT = [...EDIT_EXT, 'doc', 'xls', 'ppt', 'odt', 'ods', 'odp', 'rtf', 'txt'];

let win = null;
let baseUrl = '';
let pendingFiles = [];
let dirtyCount = 0;
let forceClose = false;
let activeDoc = { ext: '', dir: '' };
let docState = { open: [], opening: null };
let crashCount = 0;

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.wasm': 'application/wasm',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif',
  '.ico': 'image/x-icon', '.webp': 'image/webp', '.woff': 'font/woff', '.woff2': 'font/woff2', '.ttf': 'font/ttf',
  '.otf': 'font/otf', '.txt': 'text/plain; charset=utf-8', '.xml': 'application/xml', '.map': 'application/json',
  '.bin': 'application/octet-stream', '.bcmap': 'application/octet-stream', '.pfb': 'application/octet-stream', '.webmanifest': 'application/manifest+json',
};

// ---------- static server (127.0.0.1 only) ----------
const SW_SOURCE = `
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('message', (e) => {
  const port = e.ports && e.ports[0];
  const d = e.data || {};
  if (d.type === 'SKIP_WAITING') self.skipWaiting();
  if (!port) return;
  if (d.type === 'VERSION') port.postMessage({ type: 'VERSION', cacheVersion: 'local', vendorVersion: 'local' });
  if (d.type === 'CLIENT_COUNT') self.clients.matchAll().then((c) => port.postMessage({ type: 'CLIENT_COUNT', count: c.length, editors: c.length }));
  if (d.type === 'PREPARE_ASSISTANT') port.postMessage({ type: 'ASSISTANT_PREPARED', ok: false });
});
`;
async function resolveStatic(urlPath) {
  let root = ENGINE_DIR;
  let rel = urlPath;
  if (rel.startsWith('/__shell/')) { root = SHELL_DIR; rel = rel.slice('/__shell'.length); }
  const target = path.normalize(path.join(root, rel));
  if (target !== root && !target.startsWith(root + path.sep)) return null;
  const candidates = [target];
  if (!path.extname(target)) candidates.push(target + '.html', path.join(target, 'index.html'));
  for (const c of candidates) {
    try { const st = await fsp.stat(c); if (st.isFile()) return { file: c, size: st.size }; } catch (_) { /* next */ }
  }
  return null;
}

function startServer() {
  const server = http.createServer(async (req, res) => {
    const headers = {
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
      'Cross-Origin-Resource-Policy': 'same-origin',
      'X-Content-Type-Options': 'nosniff',
    };
    try {
      if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405, headers); return res.end(); }
      const host = String(req.headers.host || '');
      if (!/^(127\.0\.0\.1|localhost):\d+$/.test(host)) { res.writeHead(403, headers); return res.end(); }
      const urlPath = decodeURIComponent(new URL(req.url, 'http://127.0.0.1').pathname);
      if (urlPath === '/sw.js') {
        // Everything is already on disk, so the engine's caching worker is replaced by a
        // pass-through one. The editor only needs a worker to be in control before it starts.
        res.writeHead(200, { ...headers, 'Content-Type': MIME['.js'], 'Cache-Control': 'no-cache', 'Service-Worker-Allowed': '/' });
        return res.end(SW_SOURCE);
      }
      const hit = await resolveStatic(urlPath);
      if (!hit) { res.writeHead(404, { ...headers, 'Content-Type': 'text/plain' }); return res.end('Not found'); }
      const ext = path.extname(hit.file).toLowerCase();
      headers['Content-Type'] = MIME[ext] || 'application/octet-stream';
      headers['Content-Length'] = hit.size;
      headers['Cache-Control'] = 'no-cache';
      res.writeHead(200, headers);
      if (req.method === 'HEAD') return res.end();
      fs.createReadStream(hit.file).on('error', () => res.destroy()).pipe(res);
    } catch (e) {
      res.writeHead(500, headers); res.end();
    }
  });
  return new Promise((resolve, reject) => {
    server.once('error', (err) => {
      if (err.code === 'EADDRINUSE') {
        server.removeAllListeners('error');
        server.once('error', reject);
        server.listen(0, '127.0.0.1', () => resolve(server.address().port));
      } else reject(err);
    });
    server.listen(PREFERRED_PORT, '127.0.0.1', () => resolve(server.address().port));
  });
}

// ---------- settings / recent ----------
const dataFile = (name) => path.join(app.getPath('userData'), name);
async function readJson(name, fallback) {
  try { return JSON.parse(await fsp.readFile(dataFile(name), 'utf8')); } catch (_) { return fallback; }
}
async function writeJson(name, value) {
  await fsp.mkdir(app.getPath('userData'), { recursive: true });
  await fsp.writeFile(dataFile(name), JSON.stringify(value, null, 2));
}

async function isDir(p) { try { return (await fsp.stat(p)).isDirectory(); } catch (_) { return false; } }

async function findRoots() {
  const roots = [];
  const seen = new Set();
  const add = async (name, p, kind) => {
    if (!p) return;
    const key = path.resolve(p).toLowerCase();
    if (seen.has(key) || !(await isDir(p))) return;
    seen.add(key); roots.push({ name, path: path.resolve(p), kind });
  };
  for (const v of ['OneDriveCommercial', 'OneDriveConsumer', 'OneDrive']) {
    const p = process.env[v];
    if (p) await add(path.basename(p), p, 'onedrive');
  }
  if (process.platform === 'darwin') {
    // On a Mac the OneDrive app keeps its folders under ~/Library/CloudStorage.
    const cloud = path.join(os.homedir(), 'Library', 'CloudStorage');
    try {
      for (const name of await fsp.readdir(cloud)) { if (/^OneDrive/i.test(name)) await add(name, path.join(cloud, name), 'onedrive'); }
    } catch (_) { /* no cloud folders */ }
  }
  try {
    for (const e of await fsp.readdir(os.homedir(), { withFileTypes: true })) {
      if (e.isDirectory() && /^OneDrive/i.test(e.name)) await add(e.name, path.join(os.homedir(), e.name), 'onedrive');
    }
  } catch (_) { /* ignore */ }
  for (const p of settings.extraRoots || []) await add(path.basename(p) || p, p, 'folder');
  try { await add('Documents', app.getPath('documents'), 'folder'); } catch (_) { /* ignore */ }
  return roots;
}

function officeArgs(argv) {
  return argv.filter((a) => {
    if (!a || a.startsWith('-')) return false;
    const ext = path.extname(a).slice(1).toLowerCase();
    return OPEN_EXT.includes(ext) && fs.existsSync(a);
  }).map((a) => path.resolve(a));
}

// ---------- settings, search, updates ----------
let settings = {};
const THEMES = ['system', 'light', 'dark'];
const settingsView = () => ({ theme: THEMES.includes(settings.theme) ? settings.theme : 'system', restore: settings.restore !== false,
  pins: Array.isArray(settings.pins) ? settings.pins : [], updateDir: settings.updateDir || '', autoUpdate: settings.autoUpdate !== false });
function applyTheme() { nativeTheme.themeSource = settingsView().theme; }

async function searchFiles(query) {
  const q = String(query || '').trim().toLowerCase();
  if (q.length < 2) return { results: [], truncated: false };
  const out = []; const seen = new Set(); const deadline = Date.now() + 4000; let scanned = 0; let truncated = false;
  const walk = async (dir, depth) => {
    if (out.length >= 200 || Date.now() > deadline || scanned > 60000) { truncated = true; return; }
    let ents; try { ents = await fsp.readdir(dir, { withFileTypes: true }); } catch (_) { return; }
    const dirs = [];
    for (const e of ents) {
      scanned++;
      if (e.name.startsWith('.') || e.name.startsWith('~$') || e.name.startsWith('$') || e.name === 'node_modules') continue;
      const full = path.join(dir, e.name);
      if (e.isDirectory()) { if (depth < 7) dirs.push(full); continue; }
      const ext = path.extname(e.name).slice(1).toLowerCase();
      if (!OPEN_EXT.includes(ext) || !e.name.toLowerCase().includes(q)) continue;
      const key = full.toLowerCase(); if (seen.has(key)) continue; seen.add(key);
      try { const st = await fsp.stat(full); out.push({ name: e.name, path: full, isDir: false, ext, size: st.size, mtimeMs: st.mtimeMs, sub: dir }); } catch (_) { /* gone */ }
      if (out.length >= 200) { truncated = true; return; }
    }
    for (const d of dirs) await walk(d, depth + 1);
  };
  for (const r of await findRoots()) await walk(r.path, 0);
  out.sort((a, b) => b.mtimeMs - a.mtimeMs);
  return { results: out, truncated };
}

const SETUP_RE = /^YourOfficeSuite-Setup-(\d+)\.(\d+)\.(\d+)\.exe$/i;
const verOf = (name) => { const m = SETUP_RE.exec(name); return m ? m.slice(1).map(Number) : null; };
const newer = (a, b) => { for (let i = 0; i < 3; i++) { if (a[i] !== b[i]) return a[i] > b[i]; } return false; };
async function checkFolder() {
  const dir = settings.updateDir;
  if (!dir || process.platform === 'darwin') return { configured: false, available: false };
  const cur = app.getVersion().split('.').map(Number); let best = null;
  try {
    for (const name of await fsp.readdir(dir)) { const v = verOf(name); if (v && newer(v, cur) && (!best || newer(v, best.v))) best = { v, name }; }
  } catch (_) { return { configured: true, dir, available: false, error: 'The updates folder could not be read.' }; }
  return best ? { configured: true, dir, available: true, source: 'folder', version: best.v.join('.'), file: path.join(dir, best.name) } : { configured: true, dir, available: false };
}

// Online updates come from this project's GitHub releases. This is the only place the app goes online,
// through its own session that can reach GitHub and nothing else.
const UPDATE_REPO = 'aglr-afk/YourOfficeSuite';
const UPDATE_HOSTS = /^(github\.com|api\.github\.com|[a-z0-9-]+\.githubusercontent\.com)$/i;
let updateSes = null; let checkedOnlineAt = null; let lastOffer = null;
function updateSession() {
  if (!updateSes) {
    updateSes = session.fromPartition('yos-update', { cache: false });
    updateSes.webRequest.onBeforeRequest((d, cb) => {
      let ok = false; try { const u = new URL(d.url); ok = u.protocol === 'https:' && UPDATE_HOSTS.test(u.hostname); } catch (_) { ok = false; }
      cb({ cancel: !ok });
    });
  }
  return updateSes;
}
async function fetchOnline(url, ms) {
  const ctl = new AbortController(); const timer = setTimeout(() => ctl.abort(), ms || 15000);
  try { return await updateSession().fetch(url, { signal: ctl.signal, redirect: 'follow', headers: { 'User-Agent': `YourOfficeSuite/${app.getVersion()}`, Accept: 'application/json, application/octet-stream' } }); }
  finally { clearTimeout(timer); }
}
async function checkOnline() {
  checkedOnlineAt = Date.now(); let info = null;
  try {
    // Each release carries a small latest.json; the GitHub API is the fallback for releases without one.
    let r = await fetchOnline(`https://github.com/${UPDATE_REPO}/releases/latest/download/latest.json`);
    if (r.ok) { const m = await r.json(); info = { version: String(m.version || ''), file: String(m.file || ''), sha256: m.sha256, size: m.size, notes: m.notes }; }
    else {
      r = await fetchOnline(`https://api.github.com/repos/${UPDATE_REPO}/releases/latest`);
      if (r.ok) {
        const j = await r.json(); const asset = (j.assets || []).find((x) => SETUP_RE.test(x.name));
        if (asset) info = { version: verOf(asset.name).join('.'), file: asset.name, size: asset.size, notes: j.body || '', sha256: (/sha-?256[:\s`*]+([a-f0-9]{64})/i.exec(j.body || '') || [])[1] };
      }
    }
  } catch (_) { return { available: false, error: 'Could not reach GitHub to check for updates.' }; }
  if (!info) return { available: false };
  const v = verOf(info.file); const cur = app.getVersion().split('.').map(Number);
  if (!v || v.join('.') !== info.version) return { available: false };
  if (!newer(v, cur)) return { available: false, latest: info.version };
  return { available: true, source: 'online', version: info.version, file: info.file, size: Number(info.size) > 0 ? Number(info.size) : null,
    sha256: /^[a-f0-9]{64}$/i.test(String(info.sha256 || '')) ? String(info.sha256).toLowerCase() : null, notes: String(info.notes || '').slice(0, 4000),
    url: `https://github.com/${UPDATE_REPO}/releases/download/v${info.version}/${info.file}`, page: `https://github.com/${UPDATE_REPO}/releases/tag/v${info.version}` };
}
// The newest offer from the updates folder and, when asked, from GitHub.
async function checkUpdate(opts) {
  const folder = await checkFolder(); const online = opts && opts.online ? await checkOnline() : null;
  let pick = folder;
  if (online && online.available && (!folder.available || newer(online.version.split('.').map(Number), folder.version.split('.').map(Number)))) pick = online;
  lastOffer = pick.available ? pick : null;
  return { ...pick, configured: folder.configured, dir: folder.dir, checkedOnlineAt, onlineError: online && online.error ? online.error : null, folderError: folder.error || null, latest: online ? online.latest || null : null };
}
async function downloadUpdate(offer) {
  const crypto = require('crypto'); const target = path.join(app.getPath('temp'), offer.file);
  const res = await fetchOnline(offer.url, 30 * 60 * 1000);
  if (!res.ok || !res.body) throw new Error('The update could not be downloaded.');
  const total = offer.size || Number(res.headers.get('content-length')) || 0; const hash = crypto.createHash('sha256'); const out = fs.createWriteStream(target);
  let received = 0; let first = null; let lastSent = 0; const reader = res.body.getReader();
  try {
    for (;;) {
      const { done, value } = await reader.read(); if (done) break;
      const chunk = Buffer.from(value); if (!first) first = chunk; hash.update(chunk); received += chunk.length;
      if (!out.write(chunk)) await new Promise((r) => out.once('drain', r));
      if (win && Date.now() - lastSent > 200) { lastSent = Date.now(); win.webContents.send('update-progress', { received, total }); }
    }
  } finally { await new Promise((r) => out.end(r)); }
  const bad = async (why) => { await fsp.unlink(target).catch(() => {}); throw new Error(why); };
  if (!first || first.length < 2 || first[0] !== 0x4d || first[1] !== 0x5a) await bad('The downloaded update is not a setup file.');
  if (offer.size && received !== offer.size) await bad('The update download was incomplete. Please try again.');
  if (offer.sha256 && hash.digest('hex') !== offer.sha256) await bad('The downloaded update did not pass its safety check, so it was not started.');
  if (win) win.webContents.send('update-progress', { received, total: received });
  return target;
}

// ---------- file validation ----------
// A damaged Office file can take the whole editor down, so the container is checked first.
function checkOfficeZip(buf, name) {
  if (buf.length >= 8 && buf.readUInt32BE(0) === 0xd0cf11e0) {
    throw new Error(`${name} is password-protected or saved in an old binary format. Remove the password in Microsoft Office first.`);
  }
  const damaged = new Error(`${name} is damaged or incomplete and cannot be opened. If it is still syncing, wait for OneDrive to finish and try again.`);
  if (buf.length < 22 || buf.readUInt32LE(0) !== 0x04034b50) throw damaged;
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 22 - 65535); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw damaged;
  const count = buf.readUInt16LE(eocd + 10);
  const cdSize = buf.readUInt32LE(eocd + 12);
  let pos = buf.readUInt32LE(eocd + 16);
  if (pos === 0xffffffff || count === 0xffff) return; // zip64: leave it to the editor
  if (pos + cdSize > eocd) throw damaged;
  let hasTypes = false;
  for (let i = 0; i < count; i++) {
    if (pos + 46 > buf.length || buf.readUInt32LE(pos) !== 0x02014b50) throw damaged;
    const compSize = buf.readUInt32LE(pos + 20);
    const nameLen = buf.readUInt16LE(pos + 28);
    const extraLen = buf.readUInt16LE(pos + 30);
    const commentLen = buf.readUInt16LE(pos + 32);
    const localOff = buf.readUInt32LE(pos + 42);
    const entry = buf.toString('utf8', pos + 46, pos + 46 + nameLen);
    if (localOff + 30 > buf.length || buf.readUInt32LE(localOff) !== 0x04034b50) throw damaged;
    if (compSize !== 0xffffffff && localOff + 30 + compSize > buf.length) throw damaged;
    if (entry === '[Content_Types].xml') hasTypes = true;
    pos += 46 + nameLen + extraLen + commentLen;
  }
  if (!hasTypes) throw new Error(`${name} is not a valid Office file.`);
}

// ---------- IPC ----------
function guard(event) {
  const url = event.senderFrame ? event.senderFrame.url : '';
  if (!url.startsWith(baseUrl + '/__shell/')) throw new Error('Blocked');
}
function handle(channel, fn) {
  ipcMain.handle(channel, async (event, ...args) => { guard(event); return fn(...args); });
}

function registerIpc() {
  handle('app:init', async () => {
    const pending = pendingFiles.splice(0); let restore = { files: [], active: null };
    if (!pending.length && settingsView().restore) {
      const last = await readJson('session.json', {});
      restore = { files: (Array.isArray(last.open) ? last.open : []).filter((f) => typeof f === 'string' && fs.existsSync(f)).slice(0, 20), active: last.active || null };
    }
    return { baseUrl, roots: await findRoots(), recent: await readJson('recent.json', []), pendingFiles: pending, restore,
      version: app.getVersion(), sep: path.sep, platform: process.platform, settings: settingsView(), update: await checkUpdate() };
  });

  handle('settings:get', async () => settingsView());
  handle('settings:set', async (patch) => {
    const p = patch || {};
    if (THEMES.includes(p.theme)) settings.theme = p.theme;
    if (typeof p.restore === 'boolean') settings.restore = p.restore;
    if (Array.isArray(p.pins)) settings.pins = p.pins.filter((x) => x && typeof x.path === 'string').slice(0, 60).map((x) => ({ path: x.path, isDir: !!x.isDir }));
    if (p.updateDir === '') settings.updateDir = '';
    if (typeof p.autoUpdate === 'boolean') settings.autoUpdate = p.autoUpdate;
    await writeJson('settings.json', settings); applyTheme();
    return settingsView();
  });
  handle('settings:pickUpdateDir', async () => {
    const r = await dialog.showOpenDialog(win, { title: 'Choose the folder where new setup files are kept', properties: ['openDirectory'] });
    if (r.canceled || !r.filePaths[0]) return null;
    settings.updateDir = r.filePaths[0]; await writeJson('settings.json', settings);
    return { settings: settingsView(), update: await checkUpdate() };
  });
  let sessionTimer = null;
  handle('session:save', async (st) => {
    const data = { open: Array.isArray(st && st.open) ? st.open.filter((f) => typeof f === 'string').slice(0, 40) : [], active: (st && st.active) || null };
    clearTimeout(sessionTimer); sessionTimer = setTimeout(() => { writeJson('session.json', data).catch(() => {}); }, 300);
  });
  handle('fs:search', async (q) => searchFiles(q));
  handle('fs:rename', async ({ file, stem }) => {
    const name = String(stem || '').trim();
    if (!name || /[\\/:*?"<>|]/.test(name) || /^\.+$/.test(name) || name.length > 180) throw new Error('That name cannot be used. Avoid \\ / : * ? " < > |');
    const target = path.join(path.dirname(file), name + path.extname(file));
    if (target === file) return { path: file };
    if (fs.existsSync(target) && target.toLowerCase() !== file.toLowerCase()) throw new Error('A file with that name is already in this folder.');
    await fsp.rename(file, target);
    const recent = (await readJson('recent.json', [])).map((r) => (r.path === file ? { ...r, path: target, name: path.basename(target) } : r));
    await writeJson('recent.json', recent);
    if (Array.isArray(settings.pins) && settings.pins.some((x) => x.path === file)) { settings.pins = settings.pins.map((x) => (x.path === file ? { ...x, path: target } : x)); await writeJson('settings.json', settings); }
    return { path: target, recent, settings: settingsView() };
  });
  handle('fs:duplicate', async (file) => {
    const ext = path.extname(file); const stem = path.basename(file, ext); const dir = path.dirname(file);
    for (let i = 1; i < 500; i++) {
      const target = path.join(dir, `${stem} (copy${i > 1 ? ' ' + i : ''})${ext}`);
      if (!fs.existsSync(target)) { await fsp.copyFile(file, target, fs.constants.COPYFILE_EXCL); return { path: target }; }
    }
    throw new Error('Could not find a free file name');
  });
  handle('update:check', async (opts) => checkUpdate({ online: !!(opts && opts.online) }));
  handle('update:install', async (version) => {
    const offer = lastOffer;
    if (!offer || !offer.available || offer.version !== version) throw new Error('That update is no longer available.');
    if (process.platform === 'darwin') {
      // A Mac gets the new version from its release page.
      if (offer.page && offer.page.startsWith(`https://github.com/${UPDATE_REPO}/`)) { shell.openExternal(offer.page); return { started: false, page: true }; }
      throw new Error('That update is no longer available.');
    }
    if (process.platform !== 'win32') throw new Error('Updates can only be installed on Windows.');
    let file = offer.file;
    if (offer.source === 'online') file = await downloadUpdate(offer);
    else if (!fs.existsSync(file)) throw new Error('That update is no longer available.');
    const child = childProcess.spawn(file, [], { detached: true, stdio: 'ignore' }); child.unref();
    forceClose = true; setTimeout(() => app.quit(), 400);
    return { started: true };
  });
  handle('update:openPage', async () => { if (lastOffer && lastOffer.page && lastOffer.page.startsWith(`https://github.com/${UPDATE_REPO}/`)) shell.openExternal(lastOffer.page); });

  handle('fs:list', async (dir) => {
    const out = [];
    for (const e of await fsp.readdir(dir, { withFileTypes: true })) {
      if (e.name.startsWith('.') || e.name.startsWith('~$')) continue;
      const full = path.join(dir, e.name);
      if (e.isDirectory()) { out.push({ name: e.name, path: full, isDir: true }); continue; }
      const ext = path.extname(e.name).slice(1).toLowerCase();
      if (!OPEN_EXT.includes(ext)) continue;
      let st = null;
      try { st = await fsp.stat(full); } catch (_) { continue; }
      out.push({ name: e.name, path: full, isDir: false, ext, size: st.size, mtimeMs: st.mtimeMs });
    }
    out.sort((a, b) => (a.isDir === b.isDir ? a.name.localeCompare(b.name, undefined, { numeric: true }) : a.isDir ? -1 : 1));
    return { dir, parent: path.dirname(dir) === dir ? null : path.dirname(dir), entries: out };
  });

  handle('fs:read', async (file) => {
    const st = await fsp.stat(file);
    const kind = path.extname(file).slice(1).toLowerCase();
    if (!OPEN_EXT.includes(kind) && !IMAGE_EXT.includes(kind)) throw new Error('This file type is not supported');
    if (st.size > 600 * 1024 * 1024) throw new Error('This file is too large to open (over 600 MB)');
    const data = await fsp.readFile(file);
    if (kind === 'pdf' && data.length && data.subarray(0, 1024).indexOf('%PDF') < 0) throw new Error(`${path.basename(file)} is not a valid PDF file.`);
    if (data.length && ['docx', 'xlsx', 'pptx'].includes(path.extname(file).slice(1).toLowerCase())) checkOfficeZip(data, path.basename(file));
    return { data, mtimeMs: st.mtimeMs, size: st.size, name: path.basename(file) };
  });

  handle('fs:write', async ({ file, data, expectedMtimeMs, force }) => {
    if (!file || !data) throw new Error('Nothing to write');
    const buf = Buffer.from(data);
    if (buf.length < 4) throw new Error('Refusing to write an empty file');
    if (!force && expectedMtimeMs != null) {
      try {
        const st = await fsp.stat(file);
        if (Math.abs(st.mtimeMs - expectedMtimeMs) > 2) return { conflict: true, mtimeMs: st.mtimeMs };
      } catch (_) { /* file was removed: write it again */ }
    }
    // Write to a temp file first so a crash never leaves a half-written document.
    const tmp = path.join(path.dirname(file), `~$yos-${process.pid}-${Date.now()}.tmp`);
    await fsp.writeFile(tmp, buf);
    try {
      await fsp.copyFile(tmp, file); // keeps the same file identity, so OneDrive syncs it as a new version
    } finally {
      await fsp.unlink(tmp).catch(() => {});
    }
    const st = await fsp.stat(file);
    return { mtimeMs: st.mtimeMs, size: st.size };
  });

  handle('fs:stat', async (file) => {
    try { const st = await fsp.stat(file); return { exists: true, mtimeMs: st.mtimeMs, size: st.size }; } catch (_) { return { exists: false }; }
  });

  handle('fs:conflictCopyPath', async (file) => {
    const dir = path.dirname(file); const ext = path.extname(file); const stem = path.basename(file, ext);
    for (let i = 1; i < 500; i++) {
      const p = path.join(dir, `${stem} (my copy${i > 1 ? ' ' + i : ''})${ext}`);
      if (!fs.existsSync(p)) return p;
    }
    throw new Error('Could not find a free file name');
  });

  handle('dlg:open', async (opts) => {
    const o = opts || {};
    const sets = { office: { name: 'Documents', extensions: OPEN_EXT }, pdf: { name: 'PDF files', extensions: ['pdf'] },
      image: { name: 'Images', extensions: IMAGE_EXT }, convertible: { name: 'Word, Excel, PowerPoint', extensions: OPEN_EXT.filter((e) => e !== 'pdf') },
      meta: { name: 'PDF, Word, Excel, PowerPoint, JPG, PNG', extensions: ['pdf', 'docx', 'xlsx', 'pptx', 'jpg', 'jpeg', 'png'] } };
    const r = await dialog.showOpenDialog(win, {
      title: o.title || 'Open',
      properties: o.multi === false ? ['openFile'] : ['openFile', 'multiSelections'],
      filters: [sets[o.kind] || sets.office],
    });
    return r.canceled ? [] : r.filePaths;
  });

  handle('dlg:saveAs', async ({ defaultPath, ext }) => {
    const names = { docx: 'Word Document', xlsx: 'Excel Workbook', pptx: 'PowerPoint Presentation', csv: 'CSV', pdf: 'PDF', png: 'PNG image', jpg: 'JPEG image', txt: 'Text' };
    const r = await dialog.showSaveDialog(win, { defaultPath, filters: [{ name: names[ext] || ext, extensions: [ext] }] });
    if (r.canceled || !r.filePath) return null;
    return path.extname(r.filePath).toLowerCase() === '.' + ext ? r.filePath : r.filePath + '.' + ext;
  });

  handle('dlg:pickFolder', async (title) => {
    const r = await dialog.showOpenDialog(win, { title: title || 'Choose a folder', properties: ['openDirectory', 'createDirectory'] });
    return r.canceled || !r.filePaths[0] ? null : r.filePaths[0];
  });
  handle('fs:uniquePath', async (dir, name) => {
    const ext = path.extname(name); const stem = path.basename(name, ext);
    for (let i = 0; i < 1000; i++) {
      const p = path.join(dir, i ? `${stem} (${i})${ext}` : name);
      if (!fs.existsSync(p)) return p;
    }
    throw new Error('Could not find a free file name');
  });
  handle('app:licenses', async () => {
    const read = async (p) => { try { return await fsp.readFile(p, 'utf8'); } catch (_) { return ''; } };
    return { credits: await read(path.join(SHELL_DIR, 'licenses.txt')), notice: await read(path.join(ENGINE_DIR, 'NOTICE')) };
  });
  handle('dlg:addFolder', async () => {
    const r = await dialog.showOpenDialog(win, { properties: ['openDirectory'] });
    if (r.canceled || !r.filePaths[0]) return null;
    settings.extraRoots = Array.from(new Set([...(settings.extraRoots || []), r.filePaths[0]]));
    await writeJson('settings.json', settings);
    return findRoots();
  });

  handle('dlg:message', async (opts) => (await dialog.showMessageBox(win, opts)).response);

  handle('recent:add', async (file) => {
    let list = await readJson('recent.json', []);
    list = [{ path: file, name: path.basename(file), at: Date.now() }, ...list.filter((r) => r.path !== file)].slice(0, 40);
    await writeJson('recent.json', list);
    return list;
  });
  handle('recent:remove', async (file) => {
    const list = (await readJson('recent.json', [])).filter((r) => r.path !== file);
    await writeJson('recent.json', list);
    return list;
  });

  handle('fs:reveal', async (file) => { shell.showItemInFolder(file); });
  handle('path:join', async (dir, name) => path.join(dir, name));
  handle('app:docState', async (st) => { docState = { open: Array.isArray(st.open) ? st.open : [], opening: st.opening || null }; });
  handle('app:dirty', async (n) => { dirtyCount = Number(n) || 0; });
  handle('app:closeNow', async () => { forceClose = true; if (win) win.close(); });
  handle('app:active', async ({ title, ext, dir }) => {
    activeDoc = { ext: ext || '', dir: dir || '' };
    if (win) win.setTitle(title ? `${title} - YourOfficeSuite` : 'YourOfficeSuite');
  });
}

// ---------- window ----------
// Windows and Linux run without a menu bar. A Mac needs one for Quit, Hide and the standard editing keys.
function setAppMenu() {
  if (process.platform !== 'darwin') { Menu.setApplicationMenu(null); return; }
  const send = (cmd) => () => { if (win) win.webContents.send('cmd', cmd); };
  try { app.setAboutPanelOptions({ applicationName: 'YourOfficeSuite', applicationVersion: app.getVersion(), copyright: 'Free software under AGPL-3.0' }); } catch (_) { /* ignore */ }
  Menu.setApplicationMenu(Menu.buildFromTemplate([
    { label: 'YourOfficeSuite', submenu: [{ role: 'about', label: 'About YourOfficeSuite' }, { type: 'separator' }, { role: 'hide', label: 'Hide YourOfficeSuite' }, { role: 'hideOthers' }, { role: 'unhide' }, { type: 'separator' }, { role: 'quit', label: 'Quit YourOfficeSuite' }] },
    { label: 'File', submenu: [{ label: 'New Document', click: send('new') }, { label: 'Open…', click: send('open') }, { type: 'separator' }, { label: 'Save', click: send('save') }, { label: 'Save As…', click: send('saveAs') }, { type: 'separator' }, { label: 'Close Tab', click: send('closeTab') }] },
    { label: 'Edit', submenu: [{ role: 'undo' }, { role: 'redo' }, { type: 'separator' }, { role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { role: 'selectAll' }] },
    { label: 'Window', submenu: [{ role: 'minimize' }, { role: 'zoom' }, { role: 'togglefullscreen' }] },
    { label: 'Help', submenu: [{ label: 'Keyboard Shortcuts', click: send('shortcuts') }] },
  ]));
}
function createWindow() {
  win = new BrowserWindow({
    width: 1400, height: 900, minWidth: 900, minHeight: 600, show: false,
    title: 'YourOfficeSuite', backgroundColor: '#14121c',
    icon: path.join(SHELL_DIR, 'icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: false },
  });
  setAppMenu();
  win.once('ready-to-show', () => { win.maximize(); win.show(); });

  // Shortcuts work even while keyboard focus is inside an editor frame.
  win.webContents.on('before-input-event', (event, input) => {
    // Escape is passed on as well, so a menu or window in the app closes even while the keyboard is inside an editor.
    if (input.type === 'keyDown' && input.key === 'Escape' && !input.control && !input.alt) win.webContents.send('cmd', 'escape');
    if (input.type === 'keyDown' && input.key === 'F1' && !input.control && !input.alt) { event.preventDefault(); win.webContents.send('cmd', 'shortcuts'); return; }
    if (input.type !== 'keyDown' || !(input.control || input.meta) || input.alt) return;
    const key = String(input.key).toLowerCase();
    const map = { s: input.shift ? 'saveAs' : 'save', o: 'open', w: 'closeTab', n: 'new', tab: input.shift ? 'prevTab' : 'nextTab' };
    let cmd = map[key];
    if (!cmd && /^[1-9]$/.test(key) && !input.shift) cmd = 'tab' + key;
    if (cmd) { event.preventDefault(); win.webContents.send('cmd', cmd); }
    if (key === 'i' && input.shift && !app.isPackaged) win.webContents.openDevTools({ mode: 'detach' });
  });

  // Editor frames: make the editor's own "save to disk" use a plain download, which
  // will-download (below) routes into the app's save instead of a second file picker.
  const NO_PICKER = 'try { delete Window.prototype.showSaveFilePicker; delete window.showSaveFilePicker; } catch (e) {}';
  win.webContents.on('frame-created', (_e, { frame }) => {
    if (!frame) return;
    const strip = () => { try { if (frame.parent && frame.url.startsWith(baseUrl + '/')) frame.executeJavaScript(NO_PICKER).catch(() => {}); } catch (_) { /* frame gone */ } };
    frame.on('dom-ready', strip);
  });
  win.webContents.on('did-frame-navigate', (_e, _url, _code, _text, isMainFrame, pid, rid) => {
    if (isMainFrame) return;
    try { const f = require('electron').webFrameMain.fromId(pid, rid); if (f && f.url.startsWith(baseUrl + '/')) f.executeJavaScript(NO_PICKER).catch(() => {}); } catch (_) { /* frame gone */ }
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:\/\//i.test(url) && !url.startsWith(baseUrl)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith(baseUrl + '/__shell/')) e.preventDefault(); });

  win.on('close', (e) => {
    if (forceClose || dirtyCount === 0) return;
    e.preventDefault();
    const choice = dialog.showMessageBoxSync(win, {
      type: 'warning', buttons: ['Save all and close', 'Close without saving', 'Cancel'], defaultId: 0, cancelId: 2,
      title: 'Unsaved changes', message: `${dirtyCount} file${dirtyCount > 1 ? 's have' : ' has'} unsaved changes.`,
    });
    if (choice === 0) win.webContents.send('cmd', 'saveAllAndClose');
    else if (choice === 1) { forceClose = true; win.close(); }
  });
  // If the editor engine crashes, bring the window back with the files that were open.
  win.webContents.on('render-process-gone', (_e, details) => {
    if (!win || details.reason === 'clean-exit') return;
    crashCount += 1;
    const culprit = docState.opening;
    const reopen = crashCount > 2 ? [] : docState.open.filter((f) => f !== culprit);
    pendingFiles = reopen;
    docState = { open: [], opening: null }; dirtyCount = 0;
    dialog.showMessageBoxSync(win, {
      type: 'error', title: 'The editor stopped working', buttons: ['OK'],
      message: culprit ? `"${path.basename(culprit)}" could not be opened and the editor had to restart.` : 'The editor ran into a problem and had to restart.',
      detail: 'Your other files will be reopened from their last saved version. Changes made after the last save are lost.',
    });
    if (win) win.loadURL(baseUrl + '/__shell/index.html');
  });
  win.on('closed', () => { win = null; });
  win.loadURL(baseUrl + '/__shell/index.html');
}

// No background traffic from the browser engine (update checks, component downloads).
app.commandLine.appendSwitch('disable-component-update');
app.commandLine.appendSwitch('disable-background-networking');
app.commandLine.appendSwitch('disable-features', 'OptimizationHints,MediaRouter');

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  pendingFiles = officeArgs(process.argv.slice(1));
  // A Mac hands over files opened from Finder through this event, before or after the window exists.
  app.on('open-file', (event, file) => {
    event.preventDefault();
    const files = officeArgs([file]); if (!files.length) return;
    if (win) { if (win.isMinimized()) win.restore(); win.focus(); win.webContents.send('open-files', files); }
    else pendingFiles.push(...files);
  });
  app.on('second-instance', (_e, argv) => {
    const files = officeArgs(argv.slice(1));
    if (win) {
      if (win.isMinimized()) win.restore();
      win.focus();
      if (files.length) win.webContents.send('open-files', files);
    } else pendingFiles.push(...files);
  });

  app.whenReady().then(async () => {
    settings = await readJson('settings.json', {}); applyTheme();
    const port = await startServer();
    baseUrl = `http://127.0.0.1:${port}`;
    // The app works fully offline: block every request that is not for the local engine.
    session.defaultSession.webRequest.onBeforeRequest((details, cb) => {
      const u = details.url;
      const ok = u.startsWith(baseUrl + '/') || /^(devtools|blob|data|chrome-extension|chrome|about):/i.test(u);
      cb({ cancel: !ok });
    });
    session.defaultSession.setPermissionRequestHandler((_wc, permission, cb) => cb(['clipboard-read', 'clipboard-sanitized-write', 'fullscreen'].includes(permission)));
    // The editor's own Save button produces a download. A download in the document's own
    // format is a normal save, so it goes through the app's save. Other formats (PDF etc.)
    // are exports and get a Save dialog that starts in the document's folder.
    session.defaultSession.on('will-download', (event, item) => {
      const name = item.getFilename() || 'export';
      const ext = path.extname(name).slice(1).toLowerCase();
      if (activeDoc.ext && ext === activeDoc.ext) {
        event.preventDefault();
        if (win) win.webContents.send('cmd', 'save');
        return;
      }
      let dir = activeDoc.dir;
      if (!dir) { try { dir = app.getPath('documents'); } catch (_) { dir = os.homedir(); } }
      item.setSaveDialogOptions({ title: 'Export', defaultPath: path.join(dir, name) });
    });
    // The built-in spelling checker would fetch its dictionaries from the internet, so it stays off.
    try { session.defaultSession.setSpellCheckerEnabled(false); session.defaultSession.setSpellCheckerDictionaryDownloadURL('http://127.0.0.1:9/'); } catch (_) { /* not available on this platform */ }
    registerIpc();
    createWindow();
  });
  app.on('window-all-closed', () => app.quit());
}
