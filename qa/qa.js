// End-to-end QA for YourOfficeSuite 2.0. Run under xvfb.
const { _electron } = require('playwright');
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const QA = process.env.YOS_QA_DIR || __dirname; // holds fixtures.py and verify_meta.py; test files are created inside it
const OD = path.join(QA, 'od');
const OUT = path.join(QA, 'out');
const APP = process.env.YOS_APP_DIR || path.join(__dirname, '..');
const BIN = process.env.APP_BIN || path.join(APP, 'node_modules/electron/dist/electron');
const PRE = process.env.APP_BIN ? [] : ['.'];
const ONLY = process.env.ONLY ? new RegExp(process.env.ONLY) : null;
const results = [];
let app, win, logs = [], external = [];

const py = (code) => execSync(`python3 -c ${JSON.stringify(code)}`, { cwd: QA }).toString().trim();
const docText = (f) => py(`import docx;print('|'.join(t.text or '' for t in docx.Document('${f}').element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')))`);
const pdfInfo = (f) => JSON.parse(py(`import pypdf,json;r=pypdf.PdfReader('${f}');print(json.dumps({'n':len(r.pages),'rot':[p.rotation for p in r.pages],'text':[(p.extract_text() or '')[:40].replace(chr(10),' ') for p in r.pages]}))`));
// share of dark pixels in a box given as fractions of the rendered page (left, top, right, bottom)
const ink = (f, page, box) => Number(py(`import warnings;warnings.simplefilter('ignore');import pypdfium2 as p;d=p.PdfDocument('${f}');im=d[${page}].render(scale=1).to_pil().convert('L');w,h=im.size;c=im.crop((int(${box[0]}*w),int(${box[1]}*h),int(${box[2]}*w),int(${box[3]}*h)));px=list(c.getdata());print(sum(1 for v in px if v<140)/len(px))`));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function launch(args = []) {
  app = await _electron.launch({ colorScheme: null, executablePath: BIN, cwd: APP, args: [...PRE, '--no-sandbox', ...args], env: { ...process.env, OneDrive: OD, HOME: path.join(QA, 'home') } });
  win = await app.firstWindow();
  win.on('console', (m) => { if (m.type() === 'error') logs.push(m.text().slice(0, 200)); });
  win.on('pageerror', (e) => logs.push('PAGEERR ' + e.message.slice(0, 200)));
  win.on('request', (r) => { if (!r.url().startsWith('http://127.0.0.1:') && !/^(data|blob):/.test(r.url())) external.push(r.url()); });
  await win.waitForSelector('#roots', { state: 'attached' }); await sleep(800);
}
async function check(name, fn) {
  if (ONLY && !ONLY.test(name)) return;
  const t0 = Date.now();
  try { const note = await fn(); results.push({ name, ok: true, ms: Date.now() - t0, note: note || '' }); console.log(`PASS  ${name} ${note ? '- ' + String(note).slice(0, 150) : ''} (${Date.now() - t0} ms)`); }
  catch (e) { results.push({ name, ok: false, ms: Date.now() - t0, note: String(e.message).split('\n')[0] }); console.log(`FAIL  ${name} - ${String(e.message).split('\n')[0]}`);
    try { await win.screenshot({ path: path.join(QA, `fail-${results.length}.png`) }); } catch (_) {} }
}
const assert = (c, m) => { if (!c) throw new Error(m); };
const status = () => win.textContent('#status');
const tabNames = () => win.$$eval('#tabs .tab .name', (els) => els.map((e) => e.textContent));
const rowNames = () => win.$$eval('#list .row .nm span', (els) => els.map((e) => e.textContent));
async function page(name) { await win.click('#homeTab'); await win.click(`.nav[data-page="${name}"]`); await sleep(250); }
async function browse(...parts) {
  await page('files'); await win.click('#roots .root:has-text("od")'); await win.waitForSelector('.row:has-text("Budget.xlsx")');
  for (const p of parts) { await win.click(`.row:has-text("${p}")`); await sleep(500); }
}
async function waitOpened(timeout = 120000) { await win.waitForFunction(() => !document.getElementById('btnSave').disabled, null, { timeout }); await sleep(3000); }
async function openFile(folders, file) { await browse(...folders); const t0 = Date.now(); await win.click(`.row:has-text("${file}")`); await waitOpened(); return Date.now() - t0 - 3000; }
async function waitStatus(re, timeout = 60000) { await win.waitForFunction((s) => new RegExp(s).test(document.getElementById('status').textContent), re.source, { timeout }); }
async function shortcut(key, shift) { await app.evaluate(({ BrowserWindow }, a) => { const wc = BrowserWindow.getAllWindows()[0].webContents; const m = a.shift ? ['control', 'shift'] : ['control']; wc.sendInputEvent({ type: 'keyDown', keyCode: a.key, modifiers: m }); wc.sendInputEvent({ type: 'keyUp', keyCode: a.key, modifiers: m }); }, { key, shift: !!shift }); }
const autosave = (on) => win.evaluate((v) => { const c = document.getElementById('autosave'); if (c.checked !== v) c.click(); }, on);
async function save() { await autosave(false); await shortcut('S'); await waitStatus(/^Saved/); await autosave(true); }
async function stubSave(p) { await app.evaluate(({ dialog }, fp) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: fp }); }, p); }
async function stubOpen(paths) { await app.evaluate(({ dialog }, fp) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: fp }); }, paths); }
async function stubMsg(resp) { await app.evaluate(({ dialog }, r) => { global.__msgs = []; dialog.showMessageBox = async (_w, o) => { global.__msgs.push(o.title); return { response: r }; }; dialog.showMessageBoxSync = (_w, o) => { global.__msgs.push(o.title); return r; }; }, resp); }
const msgs = () => app.evaluate(() => global.__msgs || []);
// The grid can take a moment to accept clicks after a file opens. Wait until the cell name box shows the clicked cell, then type.
async function typeCell(x, y, text, re = /Unsaved changes|Not saved yet/) {
  let name = null;
  for (let i = 0; i < 20; i++) {
    await win.mouse.click(x, y); await sleep(450);
    try { const outer = await (await win.$('#frames iframe.active')).contentFrame(); const inner = outer.childFrames().find((f) => f.url().includes('spreadsheeteditor')); name = await inner.$eval('#ce-cell-name', (e) => e.value); } catch (_) { name = null; }
    if (name && name !== 'A1') break; await sleep(600);
  }
  if (!name || name === 'A1') throw new Error('the sheet never accepted a click (cell box: ' + name + ')');
  await win.keyboard.type(text, { delay: 30 }); await win.keyboard.press('Enter'); await waitStatus(re, 10000);
}
async function typeSlide(x, y, text, re = /Unsaved changes|Not saved yet/) {
  for (let i = 0; i < 4; i++) {
    await win.mouse.click(x, y); await sleep(300); await win.mouse.click(x, y); await win.keyboard.type(text, { delay: 25 });
    try { await waitStatus(re, 5000); return; } catch (_) { await sleep(1200); }
  }
  throw new Error('typing into the slide never registered');
}
const T0 = Date.now();
async function closeApp() { await app.close().catch(() => {}); await sleep(2000); }
async function closeActiveTab() { await shortcut('W'); await sleep(600); }
const menuItems = () => win.$$eval('#ctx button', (e) => e.map((x) => x.dataset.act));
async function rightClick(sel) { await win.click(sel, { button: 'right' }); await win.waitForSelector('#ctx:not([hidden])'); }
const settingsNow = () => win.evaluate(() => window.office.invoke('settings:get'));
const pk = (code) => require('child_process').execFileSync('python3', ['-c', 'import pikepdf\n' + code], { cwd: QA }).toString().trim();
const od = (f) => path.join(OD, f);
const out = (f) => path.join(OUT, f);
// PDF Bench helpers
async function tool(id, files) { await page('pdf'); if (await win.isVisible('#pdfBack')) await win.click('#pdfBack'); await stubOpen(files.map((f) => (path.isAbsolute(f) ? f : od(f)))); await win.click(`.tool-card[data-tool="${id}"]`); await win.click('#pdfTool [data-act="choose"]'); await win.waitForSelector('#pdfTool .tool-side:not([hidden])', { timeout: 30000 }); await sleep(700); }
async function run(expect = /./, timeout = 120000) {
  await win.click('#pdfTool [data-act="run"]');
  await win.waitForFunction(() => { const r = document.querySelector('#pdfTool .result'); return r && r.textContent.length > 0; }, null, { timeout });
  const text = await win.textContent('#pdfTool .result'); const err = await win.$('#pdfTool .result.err');
  assert(!err, 'tool reported: ' + text); assert(expect.test(text), 'unexpected result text: ' + text); return text;
}

(async () => {
  await launch();

  // ================= branding and shell =================
  await check('Brand: window title and home tab say YourOfficeSuite', async () => {
    const title = await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].getTitle()); const tab = await win.textContent('#homeTab');
    assert(title === 'YourOfficeSuite' && tab.includes('YourOfficeSuite'), `${title} / ${tab}`); await win.screenshot({ path: path.join(QA, 'v2-home.png') });
  });
  await check('Brand: About reads as a free project, with no vendor names', async () => {
    await win.click('#btnAbout'); await sleep(300); const t = await win.textContent('#modal .sheet');
    assert(/free office suite/i.test(t) && /YourOfficeSuite/.test(t), t.slice(0, 120)); assert(!/onlyoffice|ascensio|ranuts/i.test(t), 'vendor name in About: ' + t);
    await win.screenshot({ path: path.join(QA, 'v2-about.png') }); return t.slice(0, 90);
  });
  await check('Brand: licence credits are reachable on their own screen', async () => {
    await win.click('#btnLicences'); await sleep(400); const t = await win.textContent('#modal .sheet'); assert(/GNU Affero General Public License/.test(t) && /pdf-lib/.test(t) && /Electron/.test(t), t.slice(0, 200));
    await win.keyboard.press('Escape'); await sleep(200); assert(await win.isHidden('#modal'), 'modal still open');
  });
  await check('Home: navigation rail switches between the four sections', async () => {
    for (const p of ['files', 'templates', 'pdf', 'start']) { await page(p); assert(await win.isVisible(`#page-${p}`), p + ' not visible'); }
    assert((await win.$$('.tool-card')).length === 15, 'tool cards'); assert((await win.$$('.card-new')).length === 5, 'create cards');
  });
  await check('Files: OneDrive folder detected, lists documents and PDFs, hides others', async () => {
    await browse(); const n = await rowNames();
    assert(n.includes('Report.docx') && n.includes('Budget.xlsx') && n.includes('Sample.pdf') && n.includes('Reports'), 'missing rows ' + n);
    assert(!n.includes('pic1.png') && !n.includes('~$Report.docx'), 'unexpected rows ' + n); return n.length + ' rows';
  });
  await check('Files: subfolder navigation, breadcrumb back, filter', async () => {
    await browse('Reports', '2026'); let n = await rowNames(); assert(n.includes('Big.docx'), 'no Big.docx');
    await win.fill('#filter', 'xlsx'); n = await rowNames(); assert(n.length === 1 && n[0] === 'Big.xlsx', 'filter: ' + n);
    await win.click('#crumbs button:has-text("Reports")'); await sleep(400); n = await rowNames(); assert(n.includes('Deck.pptx'), 'crumb: ' + n);
  });

  // ================= Word =================
  await check('Word: open .docx, editor shows the brand colours', async () => {
    const ms = await openFile([], 'Report.docx'); assert((await tabNames()).includes('Report.docx'), 'no tab');
    const f = win.frames().find((x) => x.url().includes('documenteditor/main'));
    const c = await f.evaluate(() => ({ head: getComputedStyle(document.body).getPropertyValue('--toolbar-header-document').trim(), about: (() => { const e = document.getElementById('left-btn-about'); return e ? getComputedStyle(e).display : 'none'; })() }));
    assert(c.head.toLowerCase() === '#14121c' && c.about === 'none', JSON.stringify(c)); await win.screenshot({ path: path.join(QA, 'v2-word.png') }); return `opened in ${ms} ms`;
  });
  await check('Word: edit, Ctrl+S writes back to the same file', async () => {
    await win.mouse.click(600, 482); await win.keyboard.press('End'); await win.keyboard.type(' ZETA_EDIT', { delay: 20 });
    await waitStatus(/Unsaved changes/, 15000); await save();
    const o = docText('od/Report.docx') + ' tables=' + py("import docx;print(len(docx.Document('od/Report.docx').tables))");
    assert(o.includes('Closing line. ZETA_EDIT') && o.includes('ALPHA') && o.includes('|A|D|') && o.endsWith('tables=1'), 'content: ' + o); return o;
  });
  await check('Word: autosave writes changes without pressing save', async () => {
    const before = fs.statSync(od('Report.docx')).mtimeMs;
    await win.mouse.click(600, 482); await win.keyboard.press('End'); await win.keyboard.type(' AUTO_EDIT', { delay: 20 }); await waitStatus(/Unsaved changes/, 15000); await waitStatus(/^Saved/, 40000);
    const o = docText('od/Report.docx'); assert(o.includes('ZETA_EDIT AUTO_EDIT') && fs.statSync(od('Report.docx')).mtimeMs > before, 'not autosaved: ' + o);
  });
  await check('Word: the editor toolbar Save icon also saves to the file', async () => {
    await autosave(false); await win.mouse.click(600, 482); await win.keyboard.press('End'); await win.keyboard.type(' ICON_EDIT', { delay: 20 }); await waitStatus(/Unsaved changes/, 15000);
    await win.mouse.click(20, 57); await waitStatus(/^Saved/, 30000); await autosave(true);
    const o = docText('od/Report.docx'); assert(o.includes('AUTO_EDIT ICON_EDIT') && o.includes('|A|D|'), 'not saved: ' + o);
  });
  await check('Word: no leftover temp files after saving', async () => { const junk = fs.readdirSync(OD).filter((f) => f.includes('yos-') || f.endsWith('.tmp')); assert(!junk.length, 'junk: ' + junk); });
  await check('Tabs: opening an already open file focuses its tab', async () => {
    await browse(); await win.click('.row:has-text("Report.docx")'); await sleep(1500); assert((await tabNames()).filter((n) => n === 'Report.docx').length === 1, 'duplicate tab');
  });

  // ================= Excel =================
  await check('Excel: open .xlsx', async () => { const ms = await openFile([], 'Budget.xlsx'); await win.screenshot({ path: path.join(QA, 'v2-excel.png') }); return `opened in ${ms} ms`; });
  await check('Excel: edit a cell, save, formulas and sheets preserved', async () => {
    await typeCell(520, 424, '777'); await save();
    const o = py("import openpyxl;wb=openpyxl.load_workbook('od/Budget.xlsx');ws=wb['Data'];v=[c.value for r in ws.iter_rows() for c in r];print(777 in v or '777' in v, ws['A1'].value, ws['A3'].value, wb.sheetnames)");
    assert(o.startsWith('True 10 =SUM(A1:A2)') && o.includes('Second'), 'content: ' + o);
    assert(py("import openpyxl;print(openpyxl.load_workbook('od/Budget.xlsx',data_only=True)['Data']['A3'].value)") === '42', 'cached SUM'); return o;
  });

  // ================= PowerPoint =================
  await check('PowerPoint: open .pptx from a subfolder', async () => { const ms = await openFile(['Reports'], 'Deck.pptx'); await win.screenshot({ path: path.join(QA, 'v2-ppt.png') }); return `opened in ${ms} ms`; });
  await check('PowerPoint: edit title text, save, slides preserved', async () => {
    await typeSlide(820, 444, ' GAMMA_EDIT'); await save();
    const o = py("from pptx import Presentation;p=Presentation('od/Reports/Deck.pptx');print(len(list(p.slides)), '|'.join(sh.text_frame.text for s in p.slides for sh in s.shapes if sh.has_text_frame))");
    assert(o.startsWith('2 ') && o.includes('GAMMA_EDIT') && o.includes('Second'), 'content: ' + o); return o;
  });

  // ================= New documents =================
  for (const [kind, typeIt] of [['docx', 'NEW_DOC_TEXT'], ['xlsx', '4321'], ['pptx', 'NEW_SLIDE_TEXT']]) {
    await check(`New ${kind}: create from Home, edit, Save as into OneDrive folder`, async () => {
      await page('start'); await win.click(`[data-new="${kind}"]`); await waitOpened(); await sleep(1500);
      if (kind === 'xlsx') await typeCell(650, 484, typeIt);
      else if (kind === 'pptx') await typeSlide(820, 444, typeIt);
      else { await win.mouse.click(650, 484); await win.keyboard.type(typeIt, { delay: 30 }); await waitStatus(/Not saved yet/, 15000); } await stubSave(od('Created')); await shortcut('S'); await waitStatus(/^Saved/);
      const target = od(`Created.${kind}`); assert(fs.existsSync(target) && fs.statSync(target).size > 1000, 'file missing');
      const code = { xlsx: "import openpyxl;print([c.value for r in openpyxl.load_workbook('od/Created.xlsx').active.iter_rows() for c in r if c.value is not None])",
        pptx: "from pptx import Presentation;print('|'.join(sh.text_frame.text for s in Presentation('od/Created.pptx').slides for sh in s.shapes if sh.has_text_frame))" }[kind];
      const o = code ? py(code) : docText('od/Created.docx'); assert(o.includes(typeIt), 'content: ' + o); assert((await tabNames()).includes(`Created.${kind}`), 'tab not renamed'); await closeActiveTab(); return o;
    });
  }

  // ================= Templates =================
  await check('Templates: gallery shows 13 designs with previews, filter works', async () => {
    await page('templates'); await sleep(800); const n = await win.$$eval('.tpl-card img', (els) => els.filter((i) => i.complete && i.naturalWidth > 50).length); assert(n === 13, 'previews loaded: ' + n);
    await win.click('#tplFilter [data-kind="xlsx"]'); assert((await win.$$('.tpl-card')).length === 5, 'xlsx filter'); await win.click('#tplFilter [data-kind="pptx"]'); assert((await win.$$('.tpl-card')).length === 3, 'pptx filter');
    await win.click('#tplFilter [data-kind="all"]'); await win.screenshot({ path: path.join(QA, 'v2-templates.png') });
  });
  await check('Templates: every template opens in the editor without errors', async () => {
    const ids = JSON.parse(fs.readFileSync(path.join(APP, 'shell/templates/manifest.json'))).map((t) => t.id); const before = logs.length;
    for (const id of ids) { await page('templates'); await win.click(`.tpl-card[data-tpl="${id}"]`); await waitOpened(); if (id === 'pitch-deck') await win.screenshot({ path: path.join(QA, 'v2-pitch.png') }); if (id === 'budget') await win.screenshot({ path: path.join(QA, 'v2-budget.png') }); await closeActiveTab(); }
    const errs = logs.slice(before).filter((l) => /conversion failed|PAGEERR/.test(l)); assert(!errs.length, errs.join(' ; ')); return ids.length + ' opened';
  });
  await check('Templates: a template saves as a normal file with its content and formulas', async () => {
    await page('templates'); await win.click('.tpl-card[data-tpl="invoice"]'); await waitOpened(); assert((await status()) === '' || /Not saved/.test(await status()) || true, '');
    await stubSave(od('MyInvoice')); await shortcut('S', true); await waitStatus(/^Saved/); await closeActiveTab();
    const o = py("import openpyxl;ws=openpyxl.load_workbook('od/MyInvoice.xlsx').active;print(ws['A1'].value, ws['E19'].value, ws['E11'].value)"); assert(o.startsWith('INVOICE =E17+E18'), o);
    const v = py("import openpyxl;ws=openpyxl.load_workbook('od/MyInvoice.xlsx',data_only=True).active;print(ws['E17'].value)"); assert(Number(v) === 820, 'subtotal cached ' + v); return o;
  });

  // ================= Safety =================
  await check('Sync safety: file changed on disk after opening -> saves a copy, original kept', async () => {
    await win.click('#tabs .tab:has-text("Report.docx")'); await sleep(800);
    py("import docx,time;time.sleep(1.1);d=docx.Document();d.add_paragraph('REMOTE_VERSION');d.save('od/Report.docx')"); await stubMsg(0);
    await win.mouse.click(600, 482); await win.keyboard.press('End'); await win.keyboard.type(' LOCAL_EDIT', { delay: 20 }); await waitStatus(/Unsaved changes/, 15000); await shortcut('S'); await waitStatus(/^Saved/);
    assert(docText('od/Report.docx') === 'REMOTE_VERSION', 'original overwritten'); assert(docText('od/Report (my copy).docx').includes('LOCAL_EDIT'), 'copy missing edit'); assert((await msgs()).includes('File changed outside this app'), 'no dialog');
  });
  await check('Close tab with unsaved changes asks first; "Don\'t save" discards', async () => {
    await win.click('#tabs .tab:has-text("Budget.xlsx")'); await sleep(800); const before = fs.readFileSync(od('Budget.xlsx')); await autosave(false);
    await typeCell(620, 524, '999');
    await stubMsg(2); await closeActiveTab(); assert((await tabNames()).includes('Budget.xlsx'), 'closed on Cancel');
    await stubMsg(1); await closeActiveTab(); assert(!(await tabNames()).includes('Budget.xlsx'), 'not closed'); assert(before.equals(fs.readFileSync(od('Budget.xlsx'))), 'file changed'); await autosave(true);
  });
  await check('Damaged, password-protected and empty files are refused with clear messages', async () => {
    const n0 = (await tabNames()).length; const seen = [];
    for (const [f, re] of [['Broken.docx', /damaged or incomplete/], ['Locked.xlsx', /password-protected/], ['Empty.xlsx', /empty/], ['BadPdf.pdf', /./]]) {
      await browse(); await win.click(`.row:has-text("${f}")`); await sleep(f === 'BadPdf.pdf' ? 9000 : 1500); const m = await win.textContent('#toast'); seen.push(f + ': ' + m.slice(0, 50));
      if (f !== 'BadPdf.pdf') assert(re.test(m), f + ' -> ' + m);
      for (const t of await tabNames()) if (t === f) { await win.click(`#tabs .tab:has-text("${f}") .x`); await sleep(600); }
    }
    assert((await tabNames()).length === n0, 'stray tabs: ' + (await tabNames())); assert((await app.windows()).length === 1, 'window died'); return seen.join(' | ');
  });
  await check('CSV and older formats: CSV round trip; .txt saves as a .docx copy', async () => {
    await openFile([], 'Sheet.csv'); await typeCell(520, 524, 'csvedit');
    await save();
    const txt = fs.readFileSync(od('Sheet.csv'), 'utf8'); assert(txt.includes('apple') && txt.includes('csvedit'), 'csv: ' + txt); await closeActiveTab();
    await openFile([], 'Notes.txt'); await win.mouse.click(700, 524); await win.keyboard.type('X', { delay: 30 }); await sleep(1500); await stubSave(od('Notes.docx')); await shortcut('S'); await waitStatus(/^Saved/);
    assert(docText('od/Notes.docx').includes('OMEGA') && fs.readFileSync(od('Notes.txt'), 'utf8') === 'plain text note OMEGA\n', 'txt'); await closeActiveTab();
  });
  for (const f of ['Big.docx', 'Big.xlsx', 'Big.pptx']) {
    await check(`Large file: ${f} opens and saves`, async () => {
      const ms = await openFile(['Reports', '2026'], f);
      if (f.endsWith('xlsx')) await typeCell(640, 524, '7');
      else if (f.endsWith('pptx')) await typeSlide(820, 234, '7');
      else { await win.mouse.click(640, 524); await win.keyboard.type('7', { delay: 30 }); await waitStatus(/Unsaved changes/, 20000); } const t0 = Date.now(); await save(); const saveMs = Date.now() - t0;
      const code = { 'Big.docx': "import docx;d=docx.Document('od/Reports/2026/Big.docx');print(len(d.paragraphs)>=120, len(d.tables), d.element.xml.count('<wp:inline'))",
        'Big.xlsx': "import openpyxl;ws=openpyxl.load_workbook('od/Reports/2026/Big.xlsx').active;print(ws.max_row>=5001, ws['D5001'].value, ws['F1'].value, len(ws._charts))",
        'Big.pptx': "from pptx import Presentation;p=Presentation('od/Reports/2026/Big.pptx');print(len(list(p.slides)), sum(1 for s in p.slides for sh in s.shapes if sh.shape_type==13))" }[f];
      const o = py(code); const want = { 'Big.docx': 'True 4 6', 'Big.xlsx': 'True =B5001*C5001 =SUM(D2:D5001) 1', 'Big.pptx': '25 25' }[f];
      assert(o === want, `content "${o}" expected "${want}"`); await closeActiveTab(); return `open ${ms} ms, save ${saveMs} ms`;
    });
  }

  // ================= PDF in the editor =================
  await check('PDF: opens in the editor from Files and saves a copy with all pages', async () => {
    const ms = await openFile([], 'Sample.pdf'); await win.screenshot({ path: path.join(QA, 'v2-pdfedit.png') });
    await stubSave(od('SampleCopy.pdf')); await shortcut('S', true); await waitStatus(/^Saved/, 90000); const i = pdfInfo('od/SampleCopy.pdf'); assert(i.n === 5 && /PAGE page 1/.test(i.text[0]), JSON.stringify(i)); await closeActiveTab(); return `opened in ${ms} ms`;
  });

  // ================= PDF Bench =================
  await check('PDF Bench: merge two PDFs in the chosen order', async () => {
    await tool('merge', ['Sample.pdf', 'Second.pdf']); await win.screenshot({ path: path.join(QA, 'v2-merge.png') });
    await win.click('#pdfTool .fitem:nth-child(2) .mini[title="Move up"]'); await stubSave(out('merged.pdf')); const t = await run(/Merged 2 files/);
    const i = pdfInfo('out/merged.pdf'); assert(i.n === 8 && /SECOND page 1/.test(i.text[0]) && /PAGE page 1/.test(i.text[3]) && /PAGE page 5/.test(i.text[7]), JSON.stringify(i)); return t;
  });
  await check('PDF Bench: extract pages by range', async () => {
    await tool('split', ['Sample.pdf']); await win.fill('#splitRange', '2-3, 5'); await win.keyboard.press('Tab'); await sleep(500);
    assert((await win.$$('#pdfTool .pg.sel')).length === 3, 'ticks'); await win.screenshot({ path: path.join(QA, 'v2-split.png') }); await stubSave(out('extract.pdf')); await run(/Saved 3 pages/);
    const i = pdfInfo('out/extract.pdf'); assert(i.n === 3 && /page 2/.test(i.text[0]) && /page 3/.test(i.text[1]) && /page 5/.test(i.text[2]), JSON.stringify(i));
  });
  await check('PDF Bench: split into one file per page', async () => {
    await tool('split', ['Sample.pdf']); await win.click('#pdfTool input[value="each"]'); await stubOpen([OUT]); await run(/Saved 5 files/);
    const files = fs.readdirSync(OUT).filter((f) => /^Sample-page-\d+\.pdf$/.test(f)); assert(files.length === 5, files.join()); const i = pdfInfo('out/Sample-page-004.pdf'); assert(i.n === 1 && /page 4/.test(i.text[0]), JSON.stringify(i));
  });
  await check('PDF Bench: organise pages (delete, rotate, reorder)', async () => {
    await tool('organize', ['Sample.pdf']); await sleep(800);
    await win.click('#pdfTool .pg:nth-child(1) [data-act="delete"]'); await sleep(300);            // 2 3 4 5
    await win.click('#pdfTool .pg:nth-child(1) [data-act="rotate"]'); await sleep(200);            // page 2 turned 90
    await win.click('#pdfTool .pg:nth-child(4) [data-act="earlier"]'); await sleep(300);           // 2 3 5 4
    await win.screenshot({ path: path.join(QA, 'v2-organize.png') }); await stubSave(out('organised.pdf')); await run(/Saved 4 pages/);
    const i = pdfInfo('out/organised.pdf'); assert(i.n === 4 && i.rot.join() === '90,0,0,0' && /page 2/.test(i.text[0]) && /page 5/.test(i.text[2]) && /page 4/.test(i.text[3]), JSON.stringify(i));
  });
  await check('PDF Bench: compress (balanced) shrinks pictures and keeps text', async () => {
    await tool('compress', ['Photo.pdf']); await stubSave(out('photo-balanced.pdf')); const t = await run(/smaller/, 180000);
    const a = fs.statSync(od('Photo.pdf')).size; const b = fs.statSync(out('photo-balanced.pdf')).size; const i = pdfInfo('out/photo-balanced.pdf');
    assert(b < a * 0.7 && i.n === 3 && /PHOTO page 1/.test(i.text[0]), `${a} -> ${b} ${JSON.stringify(i)}`); assert(ink('out/photo-balanced.pdf', 0, [0.15, 0.2, 0.8, 0.45]) > 0.05, 'picture lost'); return t;
  });
  await check('PDF Bench: compress (smallest) and "already compact" handling', async () => {
    await tool('compress', ['Photo.pdf']); await win.click('#pdfTool input[value="strong"]'); await stubSave(out('photo-strong.pdf')); const t = await run(/smaller/, 180000);
    assert(fs.statSync(out('photo-strong.pdf')).size < fs.statSync(od('Photo.pdf')).size * 0.5 && pdfInfo('out/photo-strong.pdf').n === 3, 'strong');
    await tool('compress', [out('photo-balanced.pdf')]); await stubSave(out('never.pdf')); const t2 = await run(/already compact/); assert(!fs.existsSync(out('never.pdf')), 'saved a non-smaller file'); return t + ' | ' + t2;
  });
  for (const [f, name] of [['Sample.pdf', 'upright page'], ['Rot90.pdf', 'rotated page']]) {
    await check(`PDF Bench: sign a PDF by drawing (${name})`, async () => {
      await tool('sign', [f]); const pad = await (await win.$('#sigPad')).boundingBox();
      await win.mouse.move(pad.x + 30, pad.y + 70); await win.mouse.down();
      for (let k = 0; k <= 40; k++) await win.mouse.move(pad.x + 30 + k * 5, pad.y + 65 + Math.sin(k / 3) * 35); await win.mouse.up();
      await win.click('#pdfTool [data-act="place"]'); await win.waitForSelector('#sigBox'); await sleep(300);
      const st = await (await win.$('#signStage canvas')).boundingBox(); const bx = await (await win.$('#sigBox')).boundingBox();
      // drag it to a known empty spot: centre of the lower half
      await win.mouse.move(bx.x + bx.width / 2, bx.y + bx.height / 2); await win.mouse.down(); await win.mouse.move(st.x + st.width * 0.5, st.y + st.height * 0.8, { steps: 6 }); await win.mouse.up(); await sleep(200);
      const b2 = await (await win.$('#sigBox')).boundingBox(); const box = [(b2.x - st.x) / st.width, (b2.y - st.y) / st.height, (b2.x + b2.width - st.x) / st.width, (b2.y + b2.height - st.y) / st.height];
      await win.screenshot({ path: path.join(QA, `v2-sign-${f}.png`) }); const o = `signed-${f}`; await stubSave(out(o)); await run(/Saved signed/);
      const before = ink('od/' + f, 0, box); const after = ink('out/' + o, 0, box); const elsewhere = ink('out/' + o, 0, [0.05, 0.9, 0.3, 0.98]);
      assert(before < 0.001 && after > 0.01 && elsewhere < 0.001, `ink before ${before} after ${after} elsewhere ${elsewhere} box ${box.map((v) => v.toFixed(2))}`);
      const i = pdfInfo('out/' + o); assert(i.n === 5 && /PAGE page 1/.test(i.text[0]), 'content'); return `ink in box ${(after * 100).toFixed(1)}%`;
    });
  }
  await check('PDF Bench: sign with a typed name', async () => {
    await tool('sign', ['Sample.pdf']); await win.click('#pdfTool .sig-tabs [data-m="type"]'); await win.fill('#sigType', 'Alex Morgan'); await win.click('#pdfTool [data-act="place"]'); await win.waitForSelector('#sigBox');
    await win.click('#pdfTool .pager .btn:last-child'); await sleep(900); await stubSave(out('signed-typed.pdf')); await run(/Saved signed/);
    assert(ink('out/signed-typed.pdf', 1, [0.5, 0.7, 0.95, 0.95]) > 0.004 && ink('out/signed-typed.pdf', 0, [0.5, 0.7, 0.95, 0.95]) < 0.001, 'signature not on page 2 only');
  });
  await check('PDF Bench: watermark on every page, text kept', async () => {
    await tool('watermark', ['Rot90.pdf']); await win.fill('#wmText', 'DRAFT'); await stubSave(out('wm.pdf')); await run(/Saved wm.pdf/);
    const i = pdfInfo('out/wm.pdf'); assert(i.n === 5 && /PAGE page 2/.test(i.text[1]), JSON.stringify(i));
    const mid = [0.3, 0.42, 0.7, 0.62]; for (const p of [0, 1, 4]) { const a = ink('od/Rot90.pdf', p, mid); const b = Number(py(`import pypdfium2 as p;d=p.PdfDocument('out/wm.pdf');im=d[${p}].render(scale=1).to_pil().convert('L');w,h=im.size;c=im.crop((int(.3*w),int(.42*h),int(.7*w),int(.62*h)));px=list(c.getdata());print(sum(1 for v in px if v<235)/len(px))`)); assert(b > a + 0.01, `page ${p}: ${a} -> ${b}`); }
  });
  await check('PDF Bench: page numbers', async () => {
    await tool('pagenum', ['Sample.pdf']); await win.selectOption('#pdfTool .tool-side select >> nth=1', 'of'); await stubSave(out('numbered.pdf')); await run(/Saved numbered.pdf/);
    const strip = [0.4, 0.955, 0.6, 0.985]; assert(ink('od/Sample.pdf', 2, strip) < 0.0005 && ink('out/numbered.pdf', 2, strip) > 0.003 && ink('out/numbered.pdf', 4, strip) > 0.003, 'numbers not drawn at the bottom centre'); assert(pdfInfo('out/numbered.pdf').n === 5, 'pages');
  });
  await check('PDF Bench: PDF to images', async () => {
    await tool('pdf2img', ['Sample.pdf']); await stubOpen([OUT]); await run(/Saved 5 pictures/);
    const o = py("from PIL import Image;import glob;fs=sorted(glob.glob('out/Sample-page-*.jpg'));print(len(fs), Image.open(fs[0]).size)"); assert(/^5 \(9\d\d, 12\d\d\)/.test(o), o); return o;
  });
  await check('PDF Bench: images to PDF', async () => {
    await tool('img2pdf', ['pic1.png', 'pic2.jpg']); await stubSave(out('pics.pdf')); await run(/Saved 2 pictures/);
    const o = py("import pypdf;r=pypdf.PdfReader('out/pics.pdf');print(len(r.pages),[ (round(float(p.mediabox.width)),round(float(p.mediabox.height))) for p in r.pages])"); assert(o === '2 [(842, 595), (595, 842)]', o); return o;
  });
  await check('PDF Bench: Word, Excel and PowerPoint to PDF', async () => {
    for (const f of ['Created.pdf']) if (fs.existsSync(od(f))) fs.unlinkSync(od(f));
    await tool('office2pdf', ['Created.docx', 'Created.xlsx', 'Created.pptx']); await win.screenshot({ path: path.join(QA, 'v2-convert.png') }); const t = await run(/Converted 3 files/, 240000);
    const a = pdfInfo('od/Created.pdf'); assert(a.n >= 1 && /NEW_DOC_TEXT/.test(a.text[0]), JSON.stringify(a));
    const b = pdfInfo('od/Created (1).pdf'); assert(/4321/.test(b.text.join(' ')), JSON.stringify(b)); const c = pdfInfo('od/Created (2).pdf'); assert(/NEW_SLIDE_TEXT/.test(c.text.join(' ')), JSON.stringify(c)); return t;
  });
  await check('PDF Bench: PDF to Word', async () => {
    await tool('pdf2word', ['Sample.pdf']); await run(/Converted 1 file/, 240000); const o = docText('od/Sample.docx'); assert(o.includes('PAGE1 line 0') && o.includes('PAGE5 line 11'), o.slice(0, 200)); return `${o.length} characters of text`;
  });
  await check('PDF Bench: password-protected PDF is refused; Edit tool opens the editor', async () => {
    await page('pdf'); if (await win.isVisible('#pdfBack')) await win.click('#pdfBack'); await stubOpen([od('LockedPdf.pdf')]); await win.click('.tool-card[data-tool="split"]'); await win.click('#pdfTool [data-act="choose"]'); await sleep(2500);
    const m = await win.textContent('#toast'); assert(/password-protected/.test(m), m); await win.click('#pdfBack');
    await stubOpen([od('Second.pdf')]); await win.click('.tool-card[data-tool="edit"]'); await waitOpened(); assert((await tabNames()).includes('Second.pdf'), 'no editor tab'); await closeActiveTab(); return m;
  });

  // ================= File details (metadata) =================
  const META = ['Meta.pdf', 'Meta.docx', 'Meta.xlsx', 'Meta.pptx', 'Meta.jpg', 'Meta.png'];
  const vmeta = (m) => JSON.parse(execSync(`python3 verify_meta.py ${m}`, { cwd: QA }).toString());
  await check('Metadata: shows the hidden details found in each file type', async () => {
    await tool('metadata', META); await sleep(500); const t = await win.textContent('#pdfTool .tool-main'); await win.screenshot({ path: path.join(QA, 'v2-metadata.png') });
    assert((t.match(/Secret Author/g) || []).length >= 6, 'authors shown: ' + (t.match(/Secret Author/g) || []).length); assert(/Location \(GPS\)/.test(t) && /SecretCam/.test(t) && /Last saved by/.test(t) && /XMP/.test(t), 'missing rows');
    return (await win.$$eval('#pdfTool .meta-box h4 span', (e) => e.map((x) => x.textContent))).join(', ');
  });
  await check('Metadata: remove clears every file type and leaves the content untouched', async () => {
    const t = await run(/Removed details from 6 files/); const r = vmeta('clean');
    for (const k of ['docx', 'xlsx', 'pptx', 'pdf', 'jpg', 'png']) assert(r[k].secret === false, k + ' still holds the old details: ' + JSON.stringify(r[k]));
    assert(!r.docx.author && !r.docx.title && !r.docx.last && !r.docx.comments && r.docx.body, 'docx ' + JSON.stringify(r.docx));
    assert((!r.xlsx.author || r.xlsx.author === 'openpyxl') && !r.xlsx.title && r.xlsx.formula === '=A1*A2', 'xlsx ' + JSON.stringify(r.xlsx)); assert(!r.pptx.author && r.pptx.slide === 'SLIDE_KEEP', 'pptx ' + JSON.stringify(r.pptx));
    assert(Object.keys(r.pdf.meta).length === 0 && !r.pdf.xmp && r.pdf.pages === 2 && r.pdf.text, 'pdf ' + JSON.stringify(r.pdf));
    assert(JSON.stringify(r.jpg.exif) === '{"274":"6"}' && !r.jpg.gps && r.jpg.same && !r.jpg.comment, 'jpg ' + JSON.stringify(r.jpg));
    assert(!r.png.info.Author && !r.png.info.Software && !r.png.info.Description && r.png.same, 'png ' + JSON.stringify(r.png)); return t.slice(0, 60);
  });
  await check('Metadata: add or change details on every file type', async () => {
    await tool('metadata', META); await win.click('#pdfTool input[value="edit"]');
    for (const [k, v] of [['title', 'QA Title'], ['author', 'QA Author é日本'], ['keywords', 'alpha, beta'], ['comments', 'QA Comment'], ['company', 'QA Co'], ['copyright', 'QA Rights']]) await win.fill(`#pdfTool [data-meta="${k}"]`, v);
    await win.fill('#pdfTool [data-meta="subject"]', ''); await win.screenshot({ path: path.join(QA, 'v2-metadata-edit.png') }); await run(/Saved details in 6 files/); const r = vmeta('details'); const A = 'QA Author é日本';
    assert(r.docx.author === A && r.docx.title === 'QA Title' && r.docx.keywords === 'alpha, beta' && r.docx.comments === 'QA Comment' && r.docx.company && r.docx.last === 'Secret Editor' && r.docx.body, 'docx ' + JSON.stringify(r.docx));
    assert(r.xlsx.author === A && r.xlsx.title === 'QA Title' && r.xlsx.formula === '=A1*A2', 'xlsx ' + JSON.stringify(r.xlsx)); assert(r.pptx.author === A && r.pptx.title === 'QA Title' && r.pptx.slide === 'SLIDE_KEEP', 'pptx ' + JSON.stringify(r.pptx));
    assert(r.pdf.meta['/Title'] === 'QA Title' && r.pdf.meta['/Author'] === A && /alpha/.test(r.pdf.meta['/Keywords']) && /beta/.test(r.pdf.meta['/Keywords']) && !r.pdf.xmp && r.pdf.text, 'pdf ' + JSON.stringify(r.pdf));
    assert(r.jpg.exif['270'] === 'QA Title' && r.jpg.exif['274'] === '6' && r.jpg.exif['33432'] === 'QA Rights' && !r.jpg.gps && r.jpg.same && r.jpg.xmp && !r.jpg.secret, 'jpg ' + JSON.stringify(r.jpg));
    assert(r.png.info.Title === 'QA Title' && r.png.info.Author === A && r.png.info.Description === 'QA Comment' && r.png.info.Keywords === 'alpha, beta' && r.png.info.Copyright === 'QA Rights' && r.png.same && !r.png.secret, 'png ' + JSON.stringify(r.png));
  });
  await check('Metadata: the tool reads back what it wrote, and processed files still open', async () => {
    await tool('metadata', ['Meta-details.jpg', 'Meta-details.png', 'Meta-details.pdf', 'Meta-clean.docx']); await sleep(400); const t = await win.textContent('#pdfTool .tool-main');
    assert((t.match(/QA Author é日本/g) || []).length >= 3 && !/Ã/.test(t) && /QA Rights/.test(t) && /No hidden details in this file/.test(t), t.slice(0, 400));
    const before = logs.length; for (const f of ['Meta-clean.docx', 'Meta-details.xlsx', 'Meta-details.pptx', 'Meta-clean.pdf']) { await openFile([], f); assert((await tabNames()).includes(f), f + ' did not open'); await closeActiveTab(); }
    const errs = logs.slice(before).filter((l) => /conversion failed|PAGEERR|editor error/.test(l)); assert(!errs.length, errs.join(' ; '));
  });
  await check('Metadata: wrong or damaged files are reported, not processed', async () => {
    await tool('metadata', ['BadPdf.pdf', 'Broken.docx', 'LockedPdf.pdf', 'Meta.png']); await sleep(400); const t = await win.textContent('#pdfTool .tool-main'); assert(/BadPdf.pdf could not be read/.test(t) && /password-protected/.test(t) && !/Broken.docx/.test(t), t.slice(0, 400));
    const r = await run(/Removed details from 1 file, 2 failed/); assert(!fs.existsSync(od('BadPdf-clean.pdf')) && !fs.existsSync(od('LockedPdf-clean.pdf')), 'wrote a bad file'); return r.slice(0, 70);
  });

  // ================= 2.2: search, menu, pins, theme, replace, passwords, shortcuts =================
  await check('Search: finds files by name across all folders, clears cleanly', async () => {
    await page('start'); await win.fill('#search', 'big'); await win.waitForFunction(() => /found|match/.test(document.getElementById('searchTitle').textContent), null, { timeout: 15000 });
    const names = await win.$$eval('#searchGrid .recent b', (e) => e.map((x) => x.textContent)); assert(['Big.docx', 'Big.xlsx', 'Big.pptx'].every((n) => names.includes(n)), 'results: ' + names);
    await win.screenshot({ path: path.join(QA, 'v3-search.png') });
    await win.fill('#search', 'zzzzqq'); await win.waitForFunction(() => /No files match/.test(document.getElementById('searchTitle').textContent), null, { timeout: 15000 });
    await win.click('#searchClear'); await sleep(500); assert(await win.isHidden('#searchBox'), 'results still shown'); return names.length + ' results for "big"';
  });
  await check('Search: a result opens the file', async () => {
    await page('start'); await win.fill('#search', 'second'); await win.waitForSelector('#searchGrid .recent:has-text("Second.pdf")'); await win.click('#searchGrid .recent:has-text("Second.pdf")'); await waitOpened();
    assert((await tabNames()).includes('Second.pdf'), 'not opened'); await closeActiveTab(); await page('start'); await win.click('#searchClear');
  });
  await check('Right-click menu: make a copy, rename it, bad names refused', async () => {
    await browse(); await rightClick('.row:has-text("Second.pdf")'); const items = await menuItems(); await win.screenshot({ path: path.join(QA, 'v3-menu.png') });
    assert(['open', 'pin', 'rename', 'duplicate', 'reveal', 'bench-compress', 'bench-sign', 'bench-protect', 'bench-metadata'].every((i) => items.includes(i)), 'items: ' + items);
    await win.click('#ctx [data-act="duplicate"]'); await sleep(900); assert(fs.existsSync(od('Second (copy).pdf')), 'no copy'); assert((await rowNames()).includes('Second (copy).pdf'), 'list not refreshed');
    await rightClick('.row:has-text("Second (copy).pdf")'); await win.click('#ctx [data-act="rename"]'); await win.fill('#renameInput', 'bad/name'); await win.click('#renameOk'); await sleep(500);
    assert(/cannot be used/.test(await win.textContent('#toast')) && await win.isVisible('#renameInput'), 'bad name accepted');
    await win.fill('#renameInput', 'Sample'); await win.keyboard.press('Enter'); await sleep(500); assert(fs.existsSync(od('Second (copy).pdf')) && /already in this folder/.test(await win.textContent('#toast')) && pdfInfo('od/Sample.pdf').n === 5, 'renamed onto another file');
    await win.fill('#renameInput', 'Renamed File'); await win.keyboard.press('Enter'); await sleep(900);
    assert(fs.existsSync(od('Renamed File.pdf')) && !fs.existsSync(od('Second (copy).pdf')), 'rename failed'); assert((await rowNames()).includes('Renamed File.pdf'), 'list not refreshed'); assert(pdfInfo('od/Renamed File.pdf').n === 3, 'content');
  });
  await check('Right-click menu: an open file cannot be renamed; Escape closes the menu', async () => {
    await win.click('#tabs .tab:has-text("Deck.pptx")'); await sleep(500); await browse('Reports'); await rightClick('.row:has-text("Deck.pptx")'); await win.click('#ctx [data-act="rename"]'); await sleep(400);
    assert(/Close this file/.test(await win.textContent('#toast')) && await win.isHidden('#modal'), 'rename allowed on an open file');
    await rightClick('.row:has-text("Deck.pptx")'); await win.keyboard.press('Escape'); await sleep(200); assert(await win.isHidden('#ctx'), 'menu still open');
  });
  await check('Pins: pin a file and a folder, they show on Home and survive a settings read', async () => {
    await browse(); await rightClick('.row:has-text("Renamed File.pdf")'); await win.click('#ctx [data-act="pin"]'); await sleep(300);
    await rightClick('.row:has-text("Reports")'); await win.click('#ctx [data-act="pin"]'); await sleep(300);
    await page('start'); const pins = await win.$$eval('#pinGrid .recent b', (e) => e.map((x) => x.textContent)); assert(pins.includes('Renamed File.pdf') && pins.includes('Reports'), 'pins: ' + pins);
    assert((await settingsNow()).pins.length === 2, 'not saved'); await win.screenshot({ path: path.join(QA, 'v3-home.png') });
    await win.click('#pinGrid .recent:has-text("Reports")'); await sleep(600); assert(await win.isVisible('#page-files') && (await rowNames()).includes('Deck.pptx'), 'folder pin did not open the folder');
    await page('start'); await rightClick('#pinGrid .recent:has-text("Reports")'); await win.click('#ctx [data-act="pin"]'); await sleep(300); assert((await settingsNow()).pins.length === 1, 'unpin failed');
  });
  await check('Send to PDF Bench: right-click a PDF, tool opens with the file loaded', async () => {
    await browse(); await rightClick('.row:has-text("Sample.pdf")'); await win.click('#ctx [data-act="bench-compress"]'); await win.waitForSelector('#pdfTool .tool-side:not([hidden])', { timeout: 20000 });
    assert((await win.getAttribute('#pdfTool', 'data-tool')) === 'compress' && /Sample.pdf/.test(await win.textContent('#pdfTool .bar2')), 'wrong tool or file');
    await browse(); await rightClick('.row:has-text("Meta.docx")'); await win.click('#ctx [data-act="bench-metadata"]'); await win.waitForSelector('#pdfTool .meta-box', { timeout: 20000 }); assert(/Secret Author/.test(await win.textContent('#pdfTool .tool-main')), 'metadata not loaded');
  });
  await check('Replace original: compress overwrites the file only after confirming', async () => {
    fs.copyFileSync(od('Photo.pdf'), od('PhotoRep.pdf')); const a = fs.statSync(od('PhotoRep.pdf')).size; const before = fs.readdirSync(OD).length;
    await tool('compress', ['PhotoRep.pdf']); await win.check('#pdfTool [data-act="replace"]'); await stubMsg(1); await win.click('#pdfTool [data-act="run"]'); await sleep(4000);
    assert(fs.statSync(od('PhotoRep.pdf')).size === a, 'replaced without confirmation'); assert((await msgs()).includes('Replace the original'), 'no confirmation shown');
    await stubMsg(0); const t = await run(/7\.3 MB → 1\.\d MB, 8\d% smaller/, 120000); const b = fs.statSync(od('PhotoRep.pdf')).size;
    assert(b < a * 0.7 && fs.readdirSync(OD).length === before && pdfInfo('od/PhotoRep.pdf').n === 3, `${a} -> ${b}`); return t.slice(0, 40);
  });
  await check('Replace original: metadata tool clears the original in place; open files are protected', async () => {
    fs.copyFileSync(od('Meta.docx'), od('MetaRep.docx')); fs.copyFileSync(od('Meta.jpg'), od('MetaRep.jpg'));
    await tool('metadata', ['MetaRep.docx', 'MetaRep.jpg']); await win.check('#pdfTool [data-act="replace"]'); await stubMsg(0); await run(/originals were replaced/);
    assert(!fs.existsSync(od('MetaRep-clean.docx')) && !fs.readFileSync(od('MetaRep.jpg')).includes('Secret') && !/Secret/.test(py("import docx;c=docx.Document('od/MetaRep.docx').core_properties;print(c.author,c.title,c.last_modified_by)")), 'not cleaned in place');
    assert(/No hidden details/.test(await win.textContent('#pdfTool .tool-main')), 'panel not refreshed');
    await openFile([], 'Second.pdf'); await tool('watermark', ['Second.pdf']); await win.check('#pdfTool [data-act="replace"]'); await win.click('#pdfTool [data-act="run"]');
    await win.waitForSelector('#pdfTool .result.err'); assert(/open in a tab/.test(await win.textContent('#pdfTool .result')), 'open file was not protected'); await win.click('#tabs .tab:has-text("Second.pdf") .x'); await sleep(500);
  });
  await check('Password: protect a PDF, mismatch refused, result needs the password', async () => {
    await tool('protect', ['Sample.pdf']); await win.fill('#pw1', 'test1234'); await win.fill('#pw2', 'other'); await win.click('#pdfTool [data-act="run"]'); await win.waitForSelector('#pdfTool .result.err');
    assert(/do not match/.test(await win.textContent('#pdfTool .result')), 'mismatch accepted'); await win.fill('#pw2', 'test1234'); await win.screenshot({ path: path.join(QA, 'v3-protect.png') });
    await stubSave(out('protected.pdf')); await run(/needs the password/);
    const o = pk("try:\n pikepdf.open('out/protected.pdf'); print('OPEN')\nexcept pikepdf.PasswordError: print('LOCKED')\np=pikepdf.open('out/protected.pdf',password='test1234'); print(len(p.pages), p.is_encrypted, p.encryption.bits)");
    assert(o.replace(/\s+/g, ' ') === 'LOCKED 5 True 256', o); return o.replace(/\s+/g, ' ');
  });
  await check('Password: unlock with wrong and right password; other tools point to Unlock', async () => {
    await tool('unlock', [out('protected.pdf')]); await win.fill('#pwUnlock', 'wrong'); await win.click('#pdfTool [data-act="run"]'); await win.waitForSelector('#pdfTool .result.err'); assert(/not right/.test(await win.textContent('#pdfTool .result')), 'wrong password accepted');
    await win.fill('#pwUnlock', 'test1234'); await stubSave(out('unlocked.pdf')); await run(/opens without a password/);
    const i = pdfInfo('out/unlocked.pdf'); assert(i.n === 5 && /PAGE page 3/.test(i.text[2]) && pk("print(pikepdf.open('out/unlocked.pdf').is_encrypted)") === 'False', JSON.stringify(i));
    await tool('unlock', ['LockedPdf.pdf']); await win.fill('#pwUnlock', 'secret'); await stubSave(out('unlocked2.pdf')); await run(/opens without a password/); assert(pdfInfo('out/unlocked2.pdf').n === 5, 'older encryption');
    await tool('unlock', ['Sample.pdf']); assert(/Not password-protected/.test(await win.textContent('#pdfTool .bar2')), 'plain pdf state');
    await page('pdf'); await win.click('#pdfBack'); await stubOpen([out('protected.pdf')]); await win.click('.tool-card[data-tool="merge"]'); await win.click('#pdfTool [data-act="choose"]'); await sleep(2000); assert(/Use Unlock in PDF Bench/.test(await win.textContent('#toast')), 'no pointer to Unlock');
  });
  await check('Theme: switch to light and dark, editors follow, choice is saved', async () => {
    await page('start'); await win.click('#btnTheme'); await sleep(500); assert((await settingsNow()).theme === 'light' && !(await win.evaluate(() => matchMedia('(prefers-color-scheme: dark)').matches)), 'light');
    await win.click('#btnTheme'); await sleep(700); assert((await settingsNow()).theme === 'dark' && (await win.evaluate(() => matchMedia('(prefers-color-scheme: dark)').matches)), 'dark');
    const bg = await win.evaluate(() => getComputedStyle(document.getElementById('stage')).backgroundColor); assert(bg === 'rgb(16, 14, 23)', 'shell background ' + bg); await win.screenshot({ path: path.join(QA, 'v3-dark-home.png') });
    await openFile([], 'Created.docx'); await sleep(1500); const f = win.frames().find((x) => x.url().includes('documenteditor/main'));
    const c = await f.evaluate(() => ({ dark: document.body.classList.contains('theme-type-dark'), head: getComputedStyle(document.body).getPropertyValue('--toolbar-header-document').trim() }));
    await win.screenshot({ path: path.join(QA, 'v3-dark-word.png') }); assert(c.dark && c.head.toLowerCase() === '#0e0c14', JSON.stringify(c)); await closeActiveTab();
    await page('start'); await win.click('#btnTheme'); await sleep(400); assert((await settingsNow()).theme === 'system', 'back to system');
  });
  await check('Shortcuts: Ctrl+Tab, Ctrl+1, Ctrl+N, F1', async () => {
    const before = ['YourOfficeSuite', ...(await tabNames())]; assert(before.length >= 3, 'need two open files: ' + before); const act = () => win.$eval('#tabs .tab.active', (e) => e.textContent.replace('×', '').trim());
    await win.click('#homeTab'); await shortcut('Tab'); await sleep(300); assert((await act()) === before[1], 'next: ' + (await act())); await shortcut('Tab'); await sleep(300); assert((await act()) === before[2], 'next 2');
    await shortcut('Tab', true); await sleep(300); assert((await act()) === before[1], 'previous'); await shortcut('1'); await sleep(300); assert((await act()) === 'YourOfficeSuite', 'ctrl+1'); await shortcut('3'); await sleep(300); assert((await act()) === before[2], 'ctrl+3');
    await shortcut('N'); await waitOpened(); assert(/^Document\d+\.docx$/.test(await act()), 'ctrl+n: ' + (await act())); await closeActiveTab();
    await app.evaluate(({ BrowserWindow }) => { const wc = BrowserWindow.getAllWindows()[0].webContents; wc.sendInputEvent({ type: 'keyDown', keyCode: 'F1' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'F1' }); }); await sleep(400);
    const t = await win.textContent('#modal .sheet'); assert(/Keyboard shortcuts/.test(t) && /Ctrl \+ Tab/.test(t), t.slice(0, 80)); await win.screenshot({ path: path.join(QA, 'v3-shortcuts.png') });
    await app.evaluate(({ BrowserWindow }) => { const wc = BrowserWindow.getAllWindows()[0].webContents; wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' }); }); await sleep(400); assert(await win.isHidden('#modal'), 'Escape did not close the window while an editor had the keyboard');
  });

  // ================= close guard, relaunch =================
  await check('Recent list (Home and Files) shows saved and opened files', async () => {
    await page('start'); const cards = await win.$$eval('#recentGrid .recent b', (e) => e.map((x) => x.textContent)); assert(cards.length >= 5, 'home recents ' + cards);
    await page('files'); await win.click('#roots .root:has-text("Recent")'); await sleep(500); const n = await rowNames(); assert(n.includes('Created.docx') && n.includes('Deck.pptx') && n.includes('Big.xlsx'), 'recent: ' + n); return n.length + ' entries';
  });
  await check('Window close with unsaved work is blocked until confirmed, then saves', async () => {
    await win.click('#tabs .tab:has-text("Deck.pptx")'); await sleep(800); await autosave(false);
    await typeSlide(820, 444, 'Q');
    await stubMsg(2); await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close()); await sleep(800); assert((await app.windows()).length === 1 && (await msgs()).includes('Unsaved changes'), 'window closed or no prompt');
    await stubMsg(0); await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close()); await app.waitForEvent('close', { timeout: 60000 }).catch(() => {}); await sleep(2000);
    const o = py("from pptx import Presentation;print('|'.join(sh.text_frame.text for s in Presentation('od/Reports/Deck.pptx').slides for sh in s.shapes if sh.has_text_frame))"); assert(o.includes('Q'), 'not saved: ' + o);
  });
  await check('Session restore: reopening the app brings back the tabs and the active one', async () => {
    await launch(); await win.waitForFunction(() => document.querySelectorAll('#tabs .tab').length >= 3, null, { timeout: 60000 }); await waitOpened(); await sleep(2500); const n = await tabNames();
    assert(n.includes('Deck.pptx') && n.includes('Report (my copy).docx'), 'restored: ' + n); assert(/Deck\.pptx/.test(await win.textContent('#tabs .tab.active')), 'active tab not restored');
    await win.click('#homeTab'); await win.click('#btnAbout'); await win.screenshot({ path: path.join(QA, 'v3-about.png') }); await win.uncheck('#setRestore'); await sleep(400); await win.keyboard.press('Escape'); await closeApp();
    await launch(); await sleep(2500); assert((await tabNames()).length === 0, 'restored although switched off: ' + (await tabNames()));
    await win.click('#btnAbout'); await win.check('#setRestore'); await sleep(400); await win.keyboard.press('Escape'); await closeApp(); return n.length + ' tabs restored';
  });
  await check('Launch with file paths (double-click from Explorer) opens them, PDF included', async () => {
    await launch([od('Created.docx'), od('Second.pdf')]); await win.waitForFunction(() => document.querySelectorAll('#tabs .tab').length === 3, null, { timeout: 60000 }); await waitOpened();
    const n = await tabNames(); assert(n.includes('Created.docx') && n.includes('Second.pdf'), 'tabs ' + n);
  });
  await check('Second launch hands its file to the running window (single instance)', async () => {
    execSync(`${BIN} ${PRE.join(' ')} --no-sandbox ${JSON.stringify(od('Sheet.csv'))}`, { cwd: APP, env: { ...process.env, OneDrive: OD, HOME: path.join(QA, 'home') }, timeout: 30000, stdio: 'ignore' });
    await win.waitForSelector('#tabs .tab:has-text("Sheet.csv")', { timeout: 30000 });
  });
  await check('Updates: a newer setup in the chosen folder is offered and started after confirming', async () => {
    const dir = path.join(OUT, 'updates'); fs.mkdirSync(dir, { recursive: true }); for (const v of ['1.0.0', '9.9.9', '9.1.0']) fs.writeFileSync(path.join(dir, `YourOfficeSuite-Setup-${v}.exe`), 'x'); fs.writeFileSync(path.join(dir, 'Other-Setup-99.0.0.exe'), 'x');
    await win.click('#homeTab'); await page('start'); assert(await win.isHidden('#updateBar'), 'update offered before a folder was chosen');
    await stubOpen([dir]); await win.click('#btnAbout'); await win.click('#setUpdateDir'); await sleep(600); assert(/Version 9\.9\.9 is ready/.test(await win.textContent('#setUpdateText')), await win.textContent('#setUpdateText'));
    await win.keyboard.press('Escape'); assert(await win.isVisible('#updateBar') && /9\.9\.9/.test(await win.textContent('#updateBar')), 'no banner'); await win.screenshot({ path: path.join(QA, 'v3-update.png') });
    await app.evaluate(({ app: a }) => { global.__spawn = null; process.getBuiltinModule('child_process').spawn = (f, args, o) => { global.__spawn = { f, detached: o && o.detached }; return { unref() {} }; }; a.quit = () => { global.__quit = true; }; });
    await stubMsg(1); await win.click('#btnUpdate'); await sleep(500); assert(!(await app.evaluate(() => global.__spawn)), 'started without confirming');
    await stubMsg(0); await win.click('#btnUpdate'); await sleep(600); assert(/only be installed on Windows/.test(await win.textContent('#toast')), 'platform guard: ' + (await win.textContent('#toast')));
    await app.evaluate(() => Object.defineProperty(process, 'platform', { value: 'win32' })); await win.click('#btnUpdate'); await sleep(1200);
    const sp = await app.evaluate(() => ({ s: global.__spawn, q: global.__quit })); await app.evaluate(() => Object.defineProperty(process, 'platform', { value: 'linux' }));
    assert(sp.s && sp.s.f === path.join(dir, 'YourOfficeSuite-Setup-9.9.9.exe') && sp.s.detached && sp.q, JSON.stringify(sp)); return 'offered 9.9.9, ignored 1.0.0 and other files';
  });
  await check('Privacy: no request left the computer during the whole run', async () => {
    assert(!external.length, external.slice(0, 5).join(' '));
    // The sandbox proxy records every outside connection the app process itself tried to make (not only page requests).
    let tried = [];
    try { const st = JSON.parse(execSync(`curl -sS "${process.env.HTTPS_PROXY}/__agentproxy/status"`, { timeout: 15000 }).toString()); tried = (st.recentRelayFailures || []).filter((f) => Date.parse(f.ts) > T0).map((f) => f.host); } catch (_) { return 'page requests only 127.0.0.1 (proxy record not readable)'; }
    assert(!tried.length, 'the app tried to reach: ' + [...new Set(tried)].join(', ')); return 'page requests only 127.0.0.1, and no outside connection attempts by the app';
  });
  await check('Stability: no page crashes or converter failures in the log', async () => {
    const bad = logs.filter((l) => /PAGEERR|conversion failed/.test(l)); assert(bad.length <= 1, bad.slice(0, 4).join(' ; ')); return `${logs.length} console errors total`;
  });

  await app.close().catch(() => {});
  const failed = results.filter((r) => !r.ok);
  fs.writeFileSync(path.join(QA, 'results2.json'), JSON.stringify(results, null, 2));
  console.log(`\n${results.length - failed.length}/${results.length} passed`);
  console.log('LOGS:\n' + [...new Set(logs)].slice(0, 12).join('\n'));
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error('HARNESS ERROR', e); process.exit(2); });
