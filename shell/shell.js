/* YourOfficeSuite shell: tabs, home pages, file browser, open/save bridge to the editor frames. */
(() => {
  'use strict';
  const api = window.office;
  const $ = (id) => document.getElementById(id);
  const ORIGIN = location.origin;
  const KIND = { docx: 'docx', doc: 'docx', odt: 'docx', rtf: 'docx', txt: 'docx', xlsx: 'xlsx', xls: 'xlsx', ods: 'xlsx', csv: 'xlsx', pptx: 'pptx', ppt: 'pptx', odp: 'pptx', pdf: 'pdf' };
  const NATIVE = ['docx', 'xlsx', 'pptx', 'csv', 'pdf'];
  const ICON = { docx: 'i-doc', xlsx: 'i-sheet', pptx: 'i-slides', pdf: 'i-pdf' };
  const NEW_NAME = { docx: 'Document', xlsx: 'Spreadsheet', pptx: 'Presentation' };
  const AUTOSAVE_MS = 5000;
  const SVG = 'http://www.w3.org/2000/svg';

  const state = { tabs: [], workers: [], active: 'home', page: 'start', roots: [], recent: [], loc: { type: 'recent' }, entries: [], sep: '\\', newCount: {}, seq: 0, version: '', settings: { theme: 'system', restore: true, pins: [], updateDir: '', autoUpdate: true }, update: { available: false }, restoring: false, lastSession: '' };

  // ---------- helpers ----------
  const extOf = (name) => (name.includes('.') ? name.split('.').pop().toLowerCase() : '');
  const baseName = (p) => p.split(/[\\/]/).pop();
  const dirName = (p) => p.slice(0, Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/')));
  const stemOf = (name) => name.replace(/\.[^.]+$/, '');
  const uid = () => `${Date.now()}-${++state.seq}`;
  const fmtDate = (ms) => (ms ? new Date(ms).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' }) : '');
  const fmtTime = (ms) => new Date(ms).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  const fmtSize = (n) => (n == null ? '' : n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(0)} KB` : `${(n / 1048576).toFixed(1)} MB`);
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const icon = (id) => { const s = document.createElementNS(SVG, 'svg'); const u = document.createElementNS(SVG, 'use'); u.setAttribute('href', '#' + id); s.append(u); return s; };
  const badge = (cls, id) => { const b = el('i', `ico-b ${cls}`); b.append(icon(id)); return b; };
  const inOneDrive = (p) => !!p && state.roots.some((r) => r.kind === 'onedrive' && p.toLowerCase().startsWith(r.path.toLowerCase()));

  let toastTimer = null;
  function toast(msg, isErr) {
    const t = $('toast'); t.textContent = msg; t.className = 'show' + (isErr ? ' err' : '');
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.className = ''; }, isErr ? 6000 : 2600);
  }
  const errText = (e) => String((e && e.message) || e).replace(/^Error invoking remote method '[^']+': (Error: )?/, '');

  // ---------- modal ----------
  function openModal(build, wide) {
    const sheet = $('modal').firstElementChild; sheet.textContent = ''; sheet.className = 'sheet' + (wide ? ' wide' : '');
    build(sheet); $('modal').hidden = false;
    // keyboard focus moves into the window, so Tab and Enter work there and not in the document behind it
    if (!sheet.contains(document.activeElement)) { const b = sheet.querySelector('.btn.primary') || sheet.querySelector('button'); if (b) b.focus(); }
  }
  function closeModal() { $('modal').hidden = true; }

  function showAbout() {
    openModal((s) => {
      const top = el('div', 'about-top'); const img = el('img'); img.src = 'icon.png'; img.alt = '';
      const tt = el('div'); tt.append(el('h1', '', 'YourOfficeSuite'), el('span', '', `Version ${state.version}`)); top.append(img, tt);
      const pills = el('div', 'pills'); for (const p of ['Free to use', 'Works offline', 'No sign-in', 'No ads or tracking']) pills.append(el('span', '', p));
      const foot = el('div', 'foot'); const lic = el('button', 'link', 'Open-source licences'); lic.id = 'btnLicences'; const ok = el('button', 'btn primary', 'Close');
      lic.addEventListener('click', showLicences); ok.addEventListener('click', closeModal); foot.append(lic, ok);
      const set = el('div', 'set');
      const l1 = el('label', 'line'); const d1 = el('div'); d1.append(el('b', '', 'Reopen my files when the app starts'), el('span', '', 'Brings back the tabs you had open last time'));
      const c1 = el('input'); c1.type = 'checkbox'; c1.id = 'setRestore'; c1.checked = state.settings.restore; c1.addEventListener('change', async () => { state.settings = await api.invoke('settings:set', { restore: c1.checked }); }); l1.append(d1, c1);
      const l2 = el('div', 'line'); const d2 = el('div'); const upText = el('span', '', updateText()); upText.id = 'setUpdateText'; d2.append(el('b', '', 'Updates'), upText);
      const b2 = el('div'); b2.style.display = 'flex'; b2.style.gap = '6px'; b2.style.flex = 'none';
      const pick = el('button', 'btn', state.settings.updateDir ? 'Change folder' : 'Choose folder'); pick.id = 'setUpdateDir'; if (MAC) pick.style.display = 'none'; const chk = el('button', 'btn', 'Check now'); chk.id = 'setUpdateCheck';
      pick.addEventListener('click', async () => { const r = await api.invoke('settings:pickUpdateDir'); if (r) { state.settings = r.settings; state.update = r.update; pick.textContent = 'Change folder'; upText.textContent = updateText(); renderUpdate(); } });
      chk.addEventListener('click', async () => { chk.disabled = true; upText.textContent = 'Checking…'; try { state.update = await api.invoke('update:check', { online: true }); } catch (_) { /* shown as no update */ } chk.disabled = false; upText.textContent = updateText(true); renderUpdate(); });
      b2.append(pick, chk); l2.append(d2, b2);
      const l3 = el('label', 'line'); const d3 = el('div'); d3.append(el('b', '', 'Check for updates when the app starts'), el('span', '', 'Looks at this project\'s page on GitHub for a newer version. Nothing about you or your files is sent.'));
      const c3 = el('input'); c3.type = 'checkbox'; c3.id = 'setAutoUpdate'; c3.checked = state.settings.autoUpdate; c3.addEventListener('change', async () => { state.settings = await api.invoke('settings:set', { autoUpdate: c3.checked }); }); l3.append(d3, c3);
      set.append(l1, l3, l2);
      const keys = el('button', 'link', 'Keyboard shortcuts'); keys.id = 'btnKeys'; keys.prepend(icon('i-keyboard')); keys.addEventListener('click', showShortcuts);
      const left = el('div'); left.style.display = 'flex'; left.style.gap = '4px'; left.append(lic, keys); foot.textContent = ''; foot.append(left, ok);
      s.append(top,
        el('p', '', 'YourOfficeSuite is a free office suite made for everyone. Write documents, build spreadsheets, design presentations and handle PDFs, all in one app.'),
        el('p', '', 'It costs nothing and asks for nothing. There is no account to create, and your files stay on your own computer and sync through your own OneDrive.'),
        pills, set, foot);
    });
  }
  function updateText(manual) {
    const u = state.update || {};
    if (u.available && MAC) return `Version ${u.version} is available on GitHub.`;
    if (u.available) return `Version ${u.version} is ready to install${u.source === 'online' ? ' from GitHub.' : `. Folder: ${state.settings.updateDir}`}`;
    if (manual && u.onlineError) return u.onlineError;
    if (u.folderError) return u.folderError;
    const when = u.checkedOnlineAt ? ` Last checked ${fmtTime(u.checkedOnlineAt)}.` : '';
    return `You have the latest version.${when}${state.settings.updateDir ? ` Updates folder: ${state.settings.updateDir}` : (MAC ? '' : ' You can also choose a folder to pick up setup files from.')}`;
  }
  let updating = false;
  function renderUpdate() {
    const bar = $('updateBar'); const u = state.update || {}; bar.hidden = !u.available;
    if (!u.available || updating) return;
    bar.querySelector('span').textContent = `Version ${u.version} of YourOfficeSuite is ready to install.`;
    $('btnUpdateNotes').hidden = !u.page; $('btnUpdate').disabled = false; $('btnUpdate').textContent = MAC ? 'Download' : 'Install';
    if (MAC) bar.querySelector('span').textContent = `Version ${u.version} of YourOfficeSuite is available.`;
  }
  async function installUpdate() {
    const u = state.update || {}; if (!u.available || updating) return;
    if (MAC) { try { await api.invoke('update:install', u.version); } catch (e) { toast(errText(e), true); } return; }
    if (state.tabs.some((t) => t.dirty)) { toast('Save your open files first, then install the update.', true); return; }
    const notes = u.notes ? `\n\nWhat is new:\n${u.notes.split('\n').filter((l) => l.trim()).slice(0, 12).join('\n')}` : '';
    const r = await api.invoke('dlg:message', { type: 'question', buttons: ['Install now', 'Later'], defaultId: 0, cancelId: 1, title: 'Install update',
      message: `Install YourOfficeSuite ${u.version}?`, detail: `${u.source === 'online' ? 'The update will be downloaded from GitHub, then the app will close and the setup will start.' : `The app will close and the setup will start.\n${u.file}`}${notes}` });
    if (r !== 0) return;
    updating = true; $('btnUpdate').disabled = true;
    if (u.source === 'online') $('updateBar').querySelector('span').textContent = `Downloading version ${u.version}…`;
    try { await api.invoke('update:install', u.version); }
    catch (e) { toast(errText(e), true); }
    updating = false; renderUpdate();
  }
  const MAC = /Mac/i.test(navigator.platform || '');
  if (MAC) document.querySelectorAll('[title*="Ctrl"]').forEach((n) => { n.title = n.title.replace(/Ctrl/g, 'Cmd'); });
  const SHORTCUTS = [['Ctrl + N', 'New document'], ['Ctrl + O', 'Open a file'], ['Ctrl + S', 'Save'], ['Ctrl + Shift + S', 'Save as'], ['Ctrl + W', 'Close the current tab'],
    ['Ctrl + Tab', 'Next tab'], ['Ctrl + Shift + Tab', 'Previous tab'], ['Ctrl + 1', 'Go to Home'], ['Ctrl + 2 to 9', 'Go to that tab'], ['F1', 'Show this list'], ['Esc', 'Close a menu or window']].map(([k, v]) => [MAC ? k.replace(/Ctrl/g, 'Cmd') : k, v]);
  function showShortcuts() {
    openModal((s) => {
      const h = el('div', 'about-top'); const tt = el('div'); tt.append(el('h1', '', 'Keyboard shortcuts'), el('span', '', 'These work anywhere in the app')); h.append(tt);
      const g = el('div', 'keys'); for (const [k, v] of SHORTCUTS) g.append(el('kbd', '', k), el('span', '', v));
      const foot = el('div', 'foot'); const ok = el('button', 'btn primary', 'Close'); ok.addEventListener('click', closeModal); foot.append(el('span'), ok); s.append(h, g, foot);
    });
  }

  // ---------- theme ----------
  const THEME_NEXT = { system: 'light', light: 'dark', dark: 'system' };
  const THEME_UI = { system: ['i-auto', MAC ? 'Theme: follows macOS' : 'Theme: follows Windows'], light: ['i-sun', 'Theme: light'], dark: ['i-moon', 'Theme: dark'] };
  function renderTheme() { const b = $('btnTheme'); const [ic, title] = THEME_UI[state.settings.theme] || THEME_UI.system; b.textContent = ''; b.append(icon(ic)); b.title = `${title}. Click to change.`; b.dataset.theme = state.settings.theme; }
  async function cycleTheme() { state.settings = await api.invoke('settings:set', { theme: THEME_NEXT[state.settings.theme] || 'light' }); renderTheme(); toast(THEME_UI[state.settings.theme][1]); }

  // ---------- context menu ----------
  let menuAt = 0;
  function closeMenu() { $('ctx').hidden = true; }
  // a scroll that was already under way when the menu opened must not close it straight away
  const closeMenuOnScroll = () => { if (Date.now() - menuAt > 400) closeMenu(); };
  function showMenu(x, y, items) {
    const m = $('ctx'); m.textContent = '';
    for (const it of items) {
      if (it.label) { m.append(el('div', 'lab', it.label)); continue; }
      const b = el('button', '', it.text); b.setAttribute('role', 'menuitem'); b.dataset.act = it.id; b.prepend(icon(it.icon || 'i-doc'));
      b.addEventListener('click', () => { closeMenu(); it.run(); }); m.append(b);
    }
    m.hidden = false; menuAt = Date.now(); const r = m.getBoundingClientRect();
    m.style.left = `${Math.max(6, Math.min(x, innerWidth - r.width - 6))}px`; m.style.top = `${Math.max(6, Math.min(y, innerHeight - r.height - 6))}px`;
    const first = m.querySelector('button'); if (first) first.focus();
  }
  const isPinned = (p) => state.settings.pins.some((x) => x.path === p);
  async function togglePin(e) {
    const pins = isPinned(e.path) ? state.settings.pins.filter((x) => x.path !== e.path) : [{ path: e.path, isDir: !!e.isDir }, ...state.settings.pins];
    state.settings = await api.invoke('settings:set', { pins }); renderPins(); toast(isPinned(e.path) ? 'Pinned to Home' : 'Removed from pinned');
  }
  function renameFile(e) {
    if (state.tabs.some((t) => t.path && t.path.toLowerCase() === e.path.toLowerCase())) { toast('Close this file before renaming it.', true); return; }
    openModal((s) => {
      const h = el('div', 'about-top'); const tt = el('div'); tt.append(el('h1', '', 'Rename'), el('span', '', e.name)); h.append(tt);
      const inp = el('input', 'field'); inp.id = 'renameInput'; inp.value = stemOf(e.name); inp.style.width = '100%'; inp.style.marginBottom = '20px'; inp.maxLength = 180;
      const foot = el('div', 'foot'); const cancel = el('button', 'btn', 'Cancel'); const ok = el('button', 'btn primary', 'Rename'); ok.id = 'renameOk'; foot.append(cancel, ok);
      const go2 = async () => {
        try { const r = await api.invoke('fs:rename', { file: e.path, stem: inp.value }); if (r.recent) state.recent = r.recent; if (r.settings) state.settings = r.settings; closeModal(); toast(`Renamed to ${baseName(r.path)}`); renderPins(); refreshHome(); runSearch(); }
        catch (err) { toast(errText(err), true); }
      };
      cancel.addEventListener('click', closeModal); ok.addEventListener('click', go2); inp.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') go2(); });
      s.append(h, inp, foot); setTimeout(() => { inp.focus(); inp.select(); }, 30);
    });
  }
  async function duplicateFile(e) {
    try { const r = await api.invoke('fs:duplicate', e.path); toast(`Created ${baseName(r.path)}`); refreshHome(); runSearch(); } catch (err) { toast(errText(err), true); }
  }
  function fileMenu(e) {
    const items = [];
    const bench = (id, text, ic) => ({ id: 'bench-' + id, text, icon: ic, run: () => window.YOSPDF && window.YOSPDF.openTool(id, [e.path]) });
    if (e.isDir) items.push({ id: 'open', text: 'Open folder', icon: 'i-folder', run: () => openFolder(e.path) });
    else items.push({ id: 'open', text: 'Open', icon: 'i-open', run: () => openFile(e.path) });
    items.push({ id: 'pin', text: isPinned(e.path) ? 'Unpin from Home' : 'Pin to Home', icon: 'i-pin', run: () => togglePin(e) });
    if (!e.isDir) { items.push({ id: 'rename', text: 'Rename', icon: 'i-edit', run: () => renameFile(e) }, { id: 'duplicate', text: 'Make a copy', icon: 'i-copy', run: () => duplicateFile(e) }); }
    items.push({ id: 'reveal', text: 'Show in folder', icon: 'i-folder', run: () => api.invoke('fs:reveal', e.path) });
    if (!e.isDir && state.recent.some((r) => r.path === e.path)) items.push({ id: 'forget', text: 'Remove from Recent', icon: 'i-x', run: async () => { state.recent = await api.invoke('recent:remove', e.path); refreshHome(); } });
    const ext = e.isDir ? '' : extOf(e.name);
    if (ext === 'pdf') items.push({ label: 'PDF Bench' }, bench('compress', 'Compress', 'i-compress'), bench('sign', 'Sign', 'i-sign'), bench('organize', 'Organise pages', 'i-organize'), bench('split', 'Split or extract', 'i-split'),
      bench('pdf2word', 'Convert to Word', 'i-doc'), bench('protect', 'Protect with a password', 'i-lock'), bench('metadata', 'File details', 'i-tag'));
    else if (['docx', 'xlsx', 'pptx'].includes(ext)) items.push({ label: 'PDF Bench' }, bench('office2pdf', 'Convert to PDF', 'i-convert'), bench('metadata', 'File details', 'i-tag'));
    return items;
  }
  const withMenu = (node, e) => { node.addEventListener('contextmenu', (ev) => { ev.preventDefault(); showMenu(ev.clientX, ev.clientY, fileMenu(e)); }); return node; };
  function openFolder(dir) {
    const root = state.roots.find((r) => dir.toLowerCase().startsWith(r.path.toLowerCase()));
    state.loc = { type: 'dir', root: root ? root.path : dir, dir }; $('filter').value = '';
    if (state.active !== 'home') activate('home');
    showPage('files');
  }

  // ---------- cards (recent, pinned, search results) ----------
  function fileCard(e, sub) {
    const kind = e.isDir ? 'dir' : KIND[extOf(e.name)] || 'docx';
    const b = el('button', 'recent'); b.title = e.path; b.dataset.path = e.path;
    const tx = el('div'); tx.append(el('b', '', e.name), el('span', '', sub));
    b.append(badge(kind, e.isDir ? 'i-folder' : ICON[kind]), tx);
    if (isPinned(e.path)) { const p = el('span', 'pinmark'); p.append(icon('i-pin')); b.append(p); }
    b.addEventListener('click', () => (e.isDir ? openFolder(e.path) : openFile(e.path)));
    return withMenu(b, e);
  }
  function renderPins() {
    const pins = state.settings.pins || []; $('pinSec').hidden = !pins.length; const g = $('pinGrid'); g.textContent = '';
    for (const p of pins) g.append(fileCard({ path: p.path, name: baseName(p.path), isDir: p.isDir }, dirName(p.path)));
    if (state.page === 'start') renderRecent();
  }
  let searchSeq = 0; let searchTimer = null;
  async function runSearch() {
    const q = $('search').value.trim(); const box = $('searchBox'); const grid = $('searchGrid'); const seq = ++searchSeq;
    if (q.length < 2) { box.hidden = true; grid.textContent = ''; return; }
    box.hidden = false; $('searchTitle').textContent = 'Searching…';
    let res; try { res = await api.invoke('fs:search', q); } catch (e) { if (seq === searchSeq) $('searchTitle').textContent = errText(e); return; }
    if (seq !== searchSeq) return;
    grid.textContent = ''; const n = res.results.length;
    $('searchTitle').textContent = n ? `${n}${res.truncated ? '+' : ''} file${n === 1 ? '' : 's'} found` : 'No files match that name';
    for (const e of res.results) grid.append(fileCard(e, `${fmtDate(e.mtimeMs)} · ${e.sub}`));
  }

  async function showLicences() {
    let data = { credits: '', notice: '' };
    try { data = await api.invoke('app:licenses'); } catch (_) { /* shown empty */ }
    openModal((s) => {
      const h = el('div', 'about-top'); const tt = el('div'); tt.append(el('h1', '', 'Open-source licences'), el('span', '', 'The free software this app is built with')); h.append(tt);
      const pre = el('pre', '', `${data.credits}\n${data.notice ? '\n----------------------------------------\n\n' + data.notice : ''}`);
      const foot = el('div', 'foot'); const back = el('button', 'link', 'Back'); const ok = el('button', 'btn primary', 'Close');
      back.addEventListener('click', showAbout); ok.addEventListener('click', closeModal); foot.append(back, ok);
      s.append(h, pre, foot);
    }, true);
  }

  // ---------- pages ----------
  function showPage(name) {
    state.page = name;
    for (const n of document.querySelectorAll('#rail .nav[data-page]')) n.classList.toggle('active', n.dataset.page === name);
    for (const p of document.querySelectorAll('#pages .page')) p.classList.toggle('active', p.id === 'page-' + name);
    if (name === 'files') refreshList();
    if (name === 'start') renderRecent();
    document.dispatchEvent(new CustomEvent('yos:page', { detail: name }));
  }

  // ---------- tabs ----------
  const activeTab = () => state.tabs.find((t) => t.id === state.active) || null;

  function activate(id) {
    state.active = id;
    $('homeTab').classList.toggle('active', id === 'home');
    $('home').classList.toggle('active', id === 'home');
    for (const t of state.tabs) {
      t.tabEl.classList.toggle('active', t.id === id);
      t.iframe.classList.toggle('active', t.id === id);
    }
    const t = activeTab();
    if (t) { try { t.iframe.contentWindow.focus(); } catch (_) { /* not ready */ } } else refreshHome();
    updateChrome();
  }

  function updateChrome() {
    const t = activeTab();
    $('btnSave').disabled = !t || !t.opened || !!t.saving;
    $('btnSaveAs').disabled = !t || !t.opened || !!t.saving;
    let s = '';
    if (t) {
      if (!t.opened) s = 'Opening…';
      else if (t.saving) s = 'Saving…';
      else if (t.dirty) s = t.path ? 'Unsaved changes' : 'Not saved yet';
      else if (t.savedAt) s = `Saved ${fmtTime(t.savedAt)}${inOneDrive(t.path) ? ' · syncing through OneDrive' : ''}`;
      else if (t.path) s = inOneDrive(t.path) ? 'In your OneDrive folder' : 'Up to date';
    }
    $('status').textContent = s;
    api.invoke('app:active', { title: t ? t.name : '', ext: t ? t.saveExt : '', dir: t ? (t.path ? dirName(t.path) : t.sourceDir || '') : '' });
    api.invoke('app:dirty', state.tabs.filter((x) => x.dirty).length);
    const opening = state.tabs.find((x) => !x.opened && (x.path || x.sourcePath));
    api.invoke('app:docState', { open: state.tabs.filter((x) => x.opened && x.path).map((x) => x.path), opening: opening ? opening.path || opening.sourcePath : null });
    if (!state.restoring) {
      const sess = JSON.stringify({ open: state.tabs.filter((x) => x.path).map((x) => x.path), active: t && t.path ? t.path : null });
      if (sess !== state.lastSession) { state.lastSession = sess; api.invoke('session:save', JSON.parse(sess)); }
    }
  }

  function renderTab(t) {
    t.tabEl.classList.toggle('dirty', !!t.dirty);
    t.tabEl.querySelector('.name').textContent = t.name;
    t.tabEl.title = t.path || t.name;
    const b = t.tabEl.querySelector('.ico'); b.className = `ico ${t.kind}`; b.textContent = ''; b.append(icon(ICON[t.kind]));
  }

  function editorFrame() {
    const iframe = document.createElement('iframe');
    iframe.setAttribute('allow', 'clipboard-read; clipboard-write; fullscreen');
    iframe.src = `${ORIGIN}/editor?embed=1&locale=en&embedOrigin=${encodeURIComponent(ORIGIN)}`;
    return iframe;
  }

  function createTab({ name, path, kind, saveExt, mtimeMs, sourceDir }) {
    const id = uid();
    const tabEl = el('button', 'tab'); tabEl.setAttribute('role', 'tab');
    const x = el('span', 'x', '×'); x.title = MAC ? 'Close (Cmd+W)' : 'Close (Ctrl+W)'; x.setAttribute('role', 'button');
    tabEl.append(el('span', 'ico'), el('span', 'name'), el('span', 'dot'), x);
    const iframe = editorFrame();
    const t = { id, name, path: path || null, kind, saveExt, mtimeMs: mtimeMs || null, sourceDir: sourceDir || null, dirty: false,
      opened: false, ready: false, saving: null, savedAt: null, pending: new Map(), autoTimer: null, autoBlocked: false, tabEl, iframe, onReady: null };
    tabEl.addEventListener('click', (e) => { if (e.target === x) closeTab(t); else activate(id); });
    tabEl.addEventListener('auxclick', (e) => { if (e.button === 1) closeTab(t); });
    $('tabs').append(tabEl); $('frames').append(iframe);
    state.tabs.push(t); renderTab(t); activate(id);
    return t;
  }

  function destroyTab(t) {
    clearTimeout(t.autoTimer);
    for (const p of t.pending.values()) p.reject(new Error('Tab closed'));
    t.tabEl.remove(); t.iframe.remove();
    const i = state.tabs.indexOf(t); state.tabs.splice(i, 1);
    if (state.active === t.id) activate(state.tabs[i] ? state.tabs[i].id : state.tabs[i - 1] ? state.tabs[i - 1].id : 'home');
    else updateChrome();
  }

  async function closeTab(t) {
    if (t.dirty) {
      activate(t.id);
      const r = await api.invoke('dlg:message', { type: 'warning', buttons: ['Save', "Don't save", 'Cancel'], defaultId: 0, cancelId: 2,
        title: 'Unsaved changes', message: `Save changes to "${t.name}"?` });
      if (r === 2) return false;
      if (r === 0 && !(await saveTab(t, {}))) return false;
    }
    destroyTab(t);
    return true;
  }

  // ---------- editor messaging ----------
  function post(t, type, payload, transfer) {
    const id = uid();
    t.iframe.contentWindow.postMessage({ id, type, payload: payload || {} }, ORIGIN, transfer || []);
    return id;
  }

  window.addEventListener('message', async (event) => {
    if (event.origin !== ORIGIN) return;
    const msg = event.data || {};
    if (typeof msg.type !== 'string' || !msg.type.startsWith('document:')) return;
    const w = state.workers.find((x) => x.iframe.contentWindow === event.source);
    if (w) { w.onMessage(msg); return; }
    const t = state.tabs.find((x) => x.iframe.contentWindow === event.source);
    if (!t) return;
    const payload = msg.payload || {};
    switch (msg.type) {
      case 'document:ready':
        t.ready = true; if (t.onReady) { const f = t.onReady; t.onReady = null; f(); }
        break;
      case 'document:opened':
        t.opened = true; updateChrome();
        break;
      case 'document:dirty-changed':
        setDirty(t, !!payload.dirty);
        break;
      case 'document:saved': {
        const p = t.pending.get(msg.id);
        if (!p) break;
        t.pending.delete(msg.id);
        try { p.resolve({ bytes: new Uint8Array(await payload.file.arrayBuffer()), dirty: !!payload.dirty }); } catch (e) { p.reject(e); }
        break;
      }
      case 'document:error': {
        const p = t.pending.get(msg.id);
        if (p) { t.pending.delete(msg.id); p.reject(new Error(payload.message || 'The editor reported an error')); break; }
        if (!t.opened) { toast(`Could not open "${t.name}": ${payload.message || 'unknown error'}`, true); destroyTab(t); } else toast(payload.message || 'The editor reported an error', true);
        break;
      }
      default: break;
    }
  });

  function setDirty(t, dirty) {
    t.dirty = dirty; renderTab(t); updateChrome();
    if (dirty) scheduleAutosave(t);
  }

  function scheduleAutosave(t) {
    clearTimeout(t.autoTimer);
    if (!$('autosave').checked || !t.path || t.autoBlocked) return;
    t.autoTimer = setTimeout(() => { if (t.dirty && state.tabs.includes(t)) saveTab(t, { auto: true }); }, AUTOSAVE_MS);
  }

  function exportBytes(t) {
    return new Promise((resolve, reject) => {
      const id = post(t, 'document:save', { targetExt: t.saveExt.toUpperCase() });
      const timer = setTimeout(() => { if (t.pending.delete(id)) reject(new Error('The editor took too long to export the file')); }, 180000);
      t.pending.set(id, { resolve: (v) => { clearTimeout(timer); resolve(v); }, reject: (e) => { clearTimeout(timer); reject(e); } });
    });
  }

  // Runs a file through a hidden editor to convert it (for example Word to PDF).
  function convertWithEngine(bytes, fileName, targetExt) {
    return new Promise((resolve, reject) => {
      const iframe = editorFrame(); iframe.classList.add('worker');
      let saveId = null; let tries = 0; let done = false;
      const finish = (fn, v) => { if (done) return; done = true; clearTimeout(timer); state.workers.splice(state.workers.indexOf(w), 1); iframe.remove(); fn(v); };
      const timer = setTimeout(() => finish(reject, new Error('The conversion took too long')), 240000);
      const send = (type, payload, transfer) => { const id = uid(); iframe.contentWindow.postMessage({ id, type, payload }, ORIGIN, transfer || []); return id; };
      const save = () => { tries += 1; saveId = send('document:save', { targetExt: targetExt.toUpperCase() }); };
      const w = { iframe, onMessage: async (msg) => {
        const payload = msg.payload || {};
        if (msg.type === 'document:ready') {
          const u8 = new Uint8Array(bytes); const buffer = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
          send('document:open-buffer', { fileName, buffer, readonly: false }, [buffer]);
        } else if (msg.type === 'document:opened') setTimeout(save, 1200);
        else if (msg.type === 'document:saved' && msg.id === saveId) {
          try { finish(resolve, new Uint8Array(await payload.file.arrayBuffer())); } catch (e) { finish(reject, e); }
        } else if (msg.type === 'document:error') {
          if (saveId && tries < 2) setTimeout(save, 2500); else finish(reject, new Error(payload.message || 'The file could not be converted'));
        }
      } };
      state.workers.push(w); $('frames').append(iframe);
    });
  }

  // ---------- save ----------
  function saveTab(t, opts) {
    if (t.saving) return t.saving;
    if (!t.opened) return Promise.resolve(false);
    t.saving = doSave(t, opts || {}).catch((e) => { toast(`Could not save "${t.name}": ${errText(e)}`, true); return false; })
      .finally(() => {
        t.saving = null; updateChrome(); if (t.dirty) scheduleAutosave(t);
        if (state.active === t.id && document.activeElement && document.activeElement.tagName === 'BUTTON') { document.activeElement.blur(); try { t.iframe.contentWindow.focus(); } catch (_) { /* frame gone */ } }
      });
    updateChrome();
    return t.saving;
  }

  async function doSave(t, { saveAs, auto }) {
    clearTimeout(t.autoTimer);
    let target = t.path;
    if (saveAs || !target) {
      if (auto) return false;
      const dir = t.path ? dirName(t.path) : t.sourceDir || (state.loc.type === 'dir' ? state.loc.dir : (state.roots[0] && state.roots[0].path) || '');
      const defaultPath = dir ? await api.invoke('path:join', dir, `${stemOf(t.name)}.${t.saveExt}`) : `${stemOf(t.name)}.${t.saveExt}`;
      target = await api.invoke('dlg:saveAs', { defaultPath, ext: t.saveExt });
      if (!target) return false;
      const other = state.tabs.find((x) => x !== t && x.path && x.path.toLowerCase() === target.toLowerCase());
      if (other) { toast('That file is open in another tab. Choose a different name.', true); return false; }
    }
    const { bytes, dirty } = await exportBytes(t);
    const sameFile = target === t.path;
    let res = await api.invoke('fs:write', { file: target, data: bytes, expectedMtimeMs: sameFile ? t.mtimeMs : null });
    if (res.conflict) {
      activate(t.id);
      const r = await api.invoke('dlg:message', { type: 'warning', buttons: ['Save as a copy', 'Overwrite', 'Cancel'], defaultId: 0, cancelId: 2,
        title: 'File changed outside this app',
        message: `"${t.name}" was changed on disk after you opened it.`,
        detail: 'This usually means OneDrive synced a newer version from another device. Saving a copy keeps both versions.' });
      if (r === 2) { t.autoBlocked = true; return false; }
      if (r === 0) target = await api.invoke('fs:conflictCopyPath', target);
      res = await api.invoke('fs:write', { file: target, data: bytes, force: true });
    }
    t.path = target; t.name = baseName(target); t.mtimeMs = res.mtimeMs; t.savedAt = Date.now(); t.autoBlocked = false;
    t.dirty = dirty; renderTab(t);
    state.recent = await api.invoke('recent:add', target);
    if (!auto) toast(`Saved ${t.name}`);
    if (state.active === 'home') refreshHome();
    return true;
  }

  // ---------- open / new ----------
  async function openFile(file) {
    const existing = state.tabs.find((x) => x.path && x.path.toLowerCase() === file.toLowerCase());
    if (existing) { activate(existing.id); return; }
    const ext = extOf(baseName(file));
    const kind = KIND[ext];
    if (!kind) { toast(`${baseName(file)} is not a supported file type`, true); return; }
    let read;
    try { read = await api.invoke('fs:read', file); } catch (e) {
      toast(`Could not read ${baseName(file)}: ${errText(e)}`, true);
      state.recent = await api.invoke('recent:remove', file); if (state.active === 'home') refreshHome();
      return;
    }
    if (!read.size) { toast(`${read.name} is empty (0 bytes) and cannot be opened`, true); return; }
    const native = NATIVE.includes(ext);
    const t = createTab({ name: read.name, path: native ? file : null, kind, saveExt: native ? ext : kind, mtimeMs: read.mtimeMs, sourceDir: dirName(file) });
    t.sourcePath = file; updateChrome();
    const u8 = new Uint8Array(read.data);
    const buffer = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
    const send = () => post(t, 'document:open-buffer', { fileName: read.name, buffer, readonly: false }, [buffer]);
    if (t.ready) send(); else t.onReady = send;
    if (native) state.recent = await api.invoke('recent:add', file);
    else toast(`${read.name} is an older format. Saving will create a .${kind} copy.`);
  }

  // Opens bytes as a new, not yet saved file (blank documents and templates).
  function openBuffer(name, kind, buffer) {
    const t = createTab({ name, kind, saveExt: kind, sourceDir: state.loc.type === 'dir' ? state.loc.dir : null });
    const send = () => post(t, 'document:open-buffer', { fileName: name, buffer, readonly: false }, [buffer]);
    if (t.ready) send(); else t.onReady = send;
    return t;
  }

  async function fetchBytes(url) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`missing file (${res.status})`);
    return res.arrayBuffer();
  }

  async function newDoc(kind) {
    const n = (state.newCount[kind] = (state.newCount[kind] || 0) + 1);
    try { openBuffer(`${NEW_NAME[kind]}${n}.${kind}`, kind, await fetchBytes(`templates/blank.${kind}`)); }
    catch (e) { toast(`Could not create a new file: ${errText(e)}`, true); }
  }

  async function openDialog() {
    const files = await api.invoke('dlg:open');
    for (const f of files) await openFile(f);
  }

  // ---------- templates ----------
  let templates = [];
  let tplKind = 'all';
  function renderTemplates() {
    const grid = $('tplGrid'); grid.textContent = '';
    const list = templates.filter((t) => tplKind === 'all' || t.kind === tplKind);
    if (!list.length) { grid.append(el('div', 'empty', 'No templates here yet.')); return; }
    for (const t of list) {
      const card = el('button', 'tpl-card'); card.dataset.tpl = t.id; card.title = `Use "${t.name}"`;
      const th = el('div', 'thumb'); const img = el('img'); img.src = `templates/${t.thumb}`; img.alt = ''; img.loading = 'lazy'; th.append(img);
      const meta = el('div', 'meta'); const tx = el('div'); tx.append(el('b', '', t.name), el('span', '', t.desc)); meta.append(badge(t.kind, ICON[t.kind]), tx);
      card.append(th, meta);
      card.addEventListener('click', async () => {
        try { openBuffer(`${t.name}.${t.kind}`, t.kind, await fetchBytes(`templates/${t.file}`)); }
        catch (e) { toast(`Could not open the template: ${errText(e)}`, true); }
      });
      grid.append(card);
    }
  }

  // ---------- file browser ----------
  function renderRoots() {
    const box = $('roots'); box.textContent = '';
    const mk = (label, cls, ic, active, onClick, title) => {
      const b = el('button', 'root' + (active ? ' active' : '')); b.title = title || label;
      b.append(badge(cls, ic), el('span', '', label)); b.addEventListener('click', onClick); box.append(b);
    };
    mk('Recent', '', 'i-clock', state.loc.type === 'recent', () => go({ type: 'recent' }));
    for (const r of state.roots) {
      const active = state.loc.type === 'dir' && state.loc.root === r.path;
      mk(r.name, r.kind === 'onedrive' ? 'cloud' : 'dir', r.kind === 'onedrive' ? 'i-cloud' : 'i-folder', active, () => go({ type: 'dir', root: r.path, dir: r.path }), r.path);
    }
  }

  function renderCrumbs() {
    const c = $('crumbs'); c.textContent = '';
    if (state.loc.type === 'recent') { c.append(el('button', '', 'Recent files')); return; }
    const root = state.roots.find((r) => r.path === state.loc.root) || { name: state.loc.root, path: state.loc.root };
    const rest = state.loc.dir.slice(root.path.length).split(/[\\/]/).filter(Boolean);
    let acc = root.path;
    const add = (label, dir) => { const b = el('button', '', label); b.addEventListener('click', () => go({ type: 'dir', root: root.path, dir })); c.append(b); };
    add(root.name, root.path);
    for (const part of rest) { acc = acc + state.sep + part; c.append(el('span', 'sep', '›')); add(part, acc); }
  }

  function renderList() {
    const list = $('list'); list.textContent = '';
    const q = $('filter').value.trim().toLowerCase();
    const rows = state.entries.filter((e) => !q || e.name.toLowerCase().includes(q));
    if (!rows.length) {
      list.append(el('div', 'empty', state.loc.type === 'recent' ? 'No recent files yet. Open a file or pick a location on the left.' : q ? 'Nothing matches the filter.' : 'No documents, spreadsheets, presentations or PDFs in this folder.'));
      return;
    }
    for (const e of rows) {
      const row = el('button', 'row'); row.setAttribute('role', 'listitem'); row.title = e.path;
      const nm = el('div', 'nm');
      const kind = e.isDir ? 'dir' : KIND[e.ext] || '';
      nm.append(badge(kind, e.isDir ? 'i-folder' : ICON[kind] || 'i-doc'));
      const txt = el('div'); txt.style.minWidth = '0';
      txt.append(el('span', '', e.name)); if (e.sub) { txt.append(el('div', 'sub', e.sub)); }
      nm.append(txt);
      row.append(nm, el('span', 'sub', fmtDate(e.mtimeMs)), el('span', 'sub', e.isDir ? '' : fmtSize(e.size)));
      row.addEventListener('click', () => (e.isDir ? go({ type: 'dir', root: state.loc.root, dir: e.path }) : openFile(e.path)));
      list.append(withMenu(row, e));
    }
  }

  let listSeq = 0;
  async function refreshList() {
    const seq = ++listSeq; let entries = []; // a slower, older listing must not overwrite a newer one
    try {
      if (state.loc.type === 'recent') entries = state.recent.map((r) => ({ name: r.name, path: r.path, isDir: false, ext: extOf(r.name), mtimeMs: r.at, sub: dirName(r.path) }));
      else entries = (await api.invoke('fs:list', state.loc.dir)).entries;
    } catch (e) { if (seq === listSeq) toast(`Could not read that folder: ${errText(e)}`, true); }
    if (seq !== listSeq) return;
    state.entries = entries;
    renderRoots(); renderCrumbs(); renderList();
  }

  function renderRecent() {
    const g = $('recentGrid'); g.textContent = '';
    const list = state.recent.slice(0, 9);
    if (!list.length) { g.append(el('div', 'empty', 'Files you open or save will show up here.')); g.style.display = 'block'; return; }
    g.style.display = '';
    for (const r of list) g.append(fileCard({ path: r.path, name: r.name, isDir: false }, `${fmtDate(r.at)} · ${baseName(dirName(r.path)) || dirName(r.path)}`));
  }

  function refreshHome() { if (state.page === 'files') refreshList(); if (state.page === 'start') renderRecent(); }
  function go(loc) { state.loc = loc; $('filter').value = ''; refreshList(); }

  // Escape can arrive twice for one key press (from the page and from the app window), so repeats are ignored.
  let escAt = 0;
  function escapeKey() {
    if (Date.now() - escAt < 200) return; escAt = Date.now();
    if (!$('ctx').hidden) closeMenu(); else if (!$('modal').hidden) closeModal();
  }

  // ---------- commands ----------
  async function command(name) {
    const t = activeTab();
    if (name === 'save' && t) saveTab(t, {});
    else if (name === 'saveAs' && t) saveTab(t, { saveAs: true });
    else if (name === 'open') openDialog();
    else if (name === 'closeTab' && t) closeTab(t);
    else if (name === 'escape') escapeKey();
    else if (name === 'new') newDoc('docx');
    else if (name === 'shortcuts') showShortcuts();
    else if (name === 'nextTab' || name === 'prevTab') {
      const order = ['home', ...state.tabs.map((x) => x.id)]; const i = order.indexOf(state.active);
      activate(order[(i + (name === 'nextTab' ? 1 : order.length - 1)) % order.length]);
    } else if (/^tab[1-9]$/.test(name)) {
      const order = ['home', ...state.tabs.map((x) => x.id)]; const id = order[Number(name.slice(3)) - 1]; if (id) activate(id);
    }
    else if (name === 'saveAllAndClose') {
      for (const x of state.tabs.filter((y) => y.dirty)) { activate(x.id); if (!(await saveTab(x, {}))) return; }
      api.invoke('app:closeNow');
    }
  }

  // Shared with the PDF Bench module.
  window.YOS = { api, state, el, icon, badge, toast, errText, fmtSize, baseName, dirName, stemOf, extOf, openFile, convertWithEngine, showPage, refreshHome: () => refreshHome(),
    isOpen: (p) => state.tabs.some((t) => t.path && t.path.toLowerCase() === String(p).toLowerCase()),
    activateHome: () => activate('home') };

  // ---------- boot ----------
  async function boot() {
    const init = await api.invoke('app:init');
    state.roots = init.roots; state.recent = init.recent; state.sep = init.sep; state.version = init.version;
    state.settings = init.settings || state.settings; state.update = init.update || state.update; renderTheme(); renderUpdate(); renderPins();
    const od = state.roots.find((r) => r.kind === 'onedrive');
    state.loc = od ? { type: 'dir', root: od.path, dir: od.path } : { type: 'recent' };
    const h = new Date().getHours();
    $('greeting').textContent = h < 5 ? 'Working late' : h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening';
    renderRecent();

    $('homeTab').addEventListener('click', () => activate('home'));
    $('btnSave').addEventListener('click', () => command('save'));
    $('btnSaveAs').addEventListener('click', () => command('saveAs'));
    $('btnOpen').addEventListener('click', openDialog);
    $('btnRefresh').addEventListener('click', refreshList);
    $('btnAbout').addEventListener('click', showAbout);
    $('btnTheme').addEventListener('click', cycleTheme);
    $('btnUpdate').addEventListener('click', installUpdate);
    $('btnUpdateNotes').addEventListener('click', () => api.invoke('update:openPage'));
    api.onUpdateProgress((p) => { if (!updating) return; const u = state.update || {}; const pct = p.total ? Math.min(100, Math.round(p.received / p.total * 100)) : null;
      $('updateBar').querySelector('span').textContent = `Downloading version ${u.version}… ${pct == null ? fmtSize(p.received) : pct + '%'}`; });
    // Quiet check against the project's GitHub releases; any failure simply means no update is shown.
    if (state.settings.autoUpdate) api.invoke('update:check', { online: true }).then((u) => { state.update = u; renderUpdate(); }).catch(() => {});
    $('search').addEventListener('input', () => { clearTimeout(searchTimer); searchTimer = setTimeout(runSearch, 280); });
    $('searchClear').addEventListener('click', () => { $('search').value = ''; runSearch(); });
    document.addEventListener('mousedown', (e) => { if (!$('ctx').hidden && !$('ctx').contains(e.target)) closeMenu(); });
    window.addEventListener('blur', closeMenu); window.addEventListener('resize', closeMenu); document.addEventListener('scroll', closeMenuOnScroll, true);
    $('filter').addEventListener('input', renderList);
    $('modal').addEventListener('mousedown', (e) => { if (e.target === $('modal')) closeModal(); });

    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') escapeKey(); });
    // An editor that finishes loading in the background must not pull the keyboard away from Home, a menu or a window.
    let lastFocus = null;
    document.addEventListener('focusin', (e) => { if (e.target && e.target.tagName !== 'IFRAME') lastFocus = e.target; });
    window.addEventListener('blur', () => setTimeout(() => {
      const a = document.activeElement; if (!a || a.tagName !== 'IFRAME') return;
      const t = activeTab(); const allowed = t && a === t.iframe && $('modal').hidden && $('ctx').hidden;
      if (allowed) return;
      a.blur(); if (lastFocus && document.contains(lastFocus) && lastFocus.offsetParent !== null) lastFocus.focus(); else window.focus();
    }, 0));
    $('autosave').addEventListener('change', () => { for (const x of state.tabs) if (x.dirty) scheduleAutosave(x); });
    $('btnAddFolder').addEventListener('click', async () => { const roots = await api.invoke('dlg:addFolder'); if (roots) { state.roots = roots; renderRoots(); } });
    for (const b of document.querySelectorAll('[data-new]')) b.addEventListener('click', () => newDoc(b.dataset.new));
    for (const b of document.querySelectorAll('#rail .nav[data-page]')) b.addEventListener('click', () => showPage(b.dataset.page));
    for (const b of document.querySelectorAll('[data-goto]')) b.addEventListener('click', () => showPage(b.dataset.goto));
    for (const b of document.querySelectorAll('#tplFilter button')) b.addEventListener('click', () => {
      tplKind = b.dataset.kind; for (const o of document.querySelectorAll('#tplFilter button')) o.classList.toggle('on', o === b); renderTemplates();
    });
    try { templates = await (await fetch('templates/manifest.json')).json(); } catch (_) { templates = []; }
    renderTemplates();

    // Drop files anywhere on the home pages to open them (the PDF Bench handles its own drops).
    const home = $('home');
    home.addEventListener('dragover', (e) => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) e.preventDefault(); });
    home.addEventListener('drop', async (e) => {
      if (e.defaultPrevented || !e.dataTransfer || !e.dataTransfer.files.length) return;
      e.preventDefault();
      for (const f of e.dataTransfer.files) { const p = api.pathForFile(f); if (p) await openFile(p); }
    });

    api.onCommand(command);
    api.onOpenFiles(async (files) => { for (const f of files) await openFile(f); });
    for (const f of init.pendingFiles) await openFile(f);
    if (init.restore && init.restore.files.length) {
      state.restoring = true;
      try { for (const f of init.restore.files) await openFile(f); } finally { state.restoring = false; }
      const a = state.tabs.find((x) => x.path === init.restore.active); if (a) activate(a.id); else updateChrome();
    }
  }
  boot().catch((e) => toast(`Startup failed: ${errText(e)}`, true));
})();
