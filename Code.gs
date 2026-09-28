/**
 * منصة معتمد — خادم البيانات (Google Apps Script)
 * البيانات في Google Sheet «منصة معتمد - قاعدة البيانات»، والملفات في فولدر «منصة معتمد» على Google Drive
 * مرتبة: الفرع ← (أولاً - الوثائق النظامية | ثانياً - المشاريع المنفذة) ← المتطلب / التصنيف ← المشروع.
 * الواجهة على GitHub Pages وتتصل عبر doPost — بدون تسجيل دخول، أي شخص معه الرابط يرفع.
 */
const FOLDER_ID = '1z7e36oisnewWAzeqjpJYwxLxfKFMCUs3';
const FOLDER_NAME = 'منصة معتمد';
const SHEET_ID = '1LnQk-NcqtYKbk5ebDGH2h8ZNmLo3oCCqwt-G22fNelQ';
const SITE_URL = 'https://ebom2030-gif.github.io/mutamad/';

const BRANCHES = {
  makkah: ['01 مكة المكرمة', 'مكة المكرمة'],
  jeddah: ['02 جدة', 'جدة'],
  madinah: ['03 المدينة المنورة', 'المدينة المنورة'],
  baha: ['04 الباحة', 'الباحة'],
  mikhwah: ['05 المخواة', 'المخواة']
};
const DOCS_DIR = 'أولاً - الوثائق النظامية';
const PROJ_DIR = 'ثانياً - المشاريع المنفذة';
const ITEMS = {
  iban: '1 شهادة الآيبان - الحساب البنكي',
  address: '2 العنوان الوطني',
  sce: '3 ترخيص الهيئة السعودية للمهندسين',
  contact: '4 بيانات التواصل الرسمية',
  cr: '5 السجل التجاري والتراخيص الداعمة'
};
const CATS = { res: ['مشاريع سكنية', 'سكني'], com: ['مشاريع تجارية', 'تجاري'], ind: ['مشاريع صناعية', 'صناعي'] };
const STATUS = { ongoing: 'جاري', done: 'منتهي' };

const BCOLS = ['branch', 'phone', 'email', 'person', 'sceExpiry', 'crExpiry', 'updatedAt', 'updatedBy'];
const PCOLS = ['id', 'branch', 'category', 'status', 'name', 'owner', 'location', 'start', 'end', 'value', 'description', 'folderId', 'updatedAt', 'updatedBy'];
const FCOLS = ['id', 'branch', 'item', 'name', 'mime', 'size', 'url', 'at', 'by'];

const ss_ = () => SpreadsheetApp.openById(SHEET_ID);

/* ---------- sheets ---------- */
function sheet_(name, headers) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange('A:Z').setNumberFormat('@');
    s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.setRightToLeft(true);
  }
  const h = s.getRange(1, 1, 1, headers.length).getValues()[0];
  if (h.join('|') !== headers.join('|')) s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  return s;
}
const branchSheet_ = () => sheet_('Branches', BCOLS);
const projSheet_ = () => sheet_('Projects', PCOLS);
const filesSheet_ = () => sheet_('Files', FCOLS);

function rows_(s, cols) {
  const n = s.getLastRow(); if (n < 2) return [];
  return s.getRange(2, 1, n - 1, cols.length).getDisplayValues().filter(r => r[0]).map((r, i) => {
    const o = { _row: i + 2 }; cols.forEach((c, j) => o[c] = r[j]); return o;
  });
}
function strip_(o) { const x = Object.assign({}, o); delete x._row; return x; }
function put_(s, cols, row, o) {
  const vals = [cols.map(c => o[c] === undefined || o[c] === null ? '' : String(o[c]))];
  if (row > 0) s.getRange(row, 1, 1, cols.length).setValues(vals);
  else s.getRange(s.getLastRow() + 1, 1, 1, cols.length).setNumberFormat('@').setValues(vals);
}
function rid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function clean_(s) { return String(s || '').replace(/[\\/:*?"<>|#%]/g, '-').trim().slice(0, 120) || 'بدون اسم'; }

/* ---------- الوصول ----------
 * بدون تسجيل دخول: أي شخص معه الرابط يرفع بيانات فرعه.
 */
const U_ = { name: 'رابط الفرع' };
function needBranch_(bk) { if (!BRANCHES[bk]) throw new Error('فرع غير معروف'); }

/* ---------- drive folders ---------- */
function root_() {
  try { return DriveApp.getFolderById(FOLDER_ID); } catch (e) {}
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}
function sub_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function cachedFolder_(key, make) {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('f_' + key);
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = make(); p.setProperty('f_' + key, f.getId()); return f;
}
function branchFolder_(bk) { return cachedFolder_(bk, () => sub_(root_(), BRANCHES[bk][0])); }
function itemFolder_(bk, ik) {
  if (!ITEMS[ik]) throw new Error('متطلب غير معروف');
  return cachedFolder_(bk + '_' + ik, () => sub_(sub_(branchFolder_(bk), DOCS_DIR), ITEMS[ik]));
}
function catFolder_(bk, ck) {
  if (!CATS[ck]) throw new Error('تصنيف غير معروف');
  return cachedFolder_(bk + '_' + ck, () => sub_(sub_(branchFolder_(bk), PROJ_DIR), CATS[ck][0]));
}
function buildFolders_() {
  Object.keys(BRANCHES).forEach(bk => {
    Object.keys(ITEMS).forEach(ik => itemFolder_(bk, ik));
    Object.keys(CATS).forEach(ck => catFolder_(bk, ck));
  });
}
function writeText_(folder, name, text) {
  const it = folder.getFilesByName(name);
  while (it.hasNext()) it.next().setTrashed(true);
  folder.createFile(Utilities.newBlob('﻿' + text, 'text/plain', name));
}

/* ---------- data ---------- */
function getAll() {
  const vis = () => true;
  const branches = {};
  rows_(branchSheet_(), BCOLS).forEach(r => { branches[r.branch] = strip_(r); });
  const projects = rows_(projSheet_(), PCOLS).map(strip_);
  const files = rows_(filesSheet_(), FCOLS).map(strip_);
  // ملخص كل الفروع (للوحة المتابعة) + التفاصيل لفرع المستخدم فقط
  const summary = {};
  Object.keys(BRANCHES).forEach(bk => {
    const b = branches[bk] || {};
    const docs = {};
    Object.keys(ITEMS).forEach(ik => docs[ik] = ik === 'contact' ? !!(b.phone && b.email && b.person) : files.some(f => f.branch === bk && f.item === ik));
    const ps = projects.filter(p => p.branch === bk).map(p => ({ category: p.category, complete: projMissing_(p, files).length === 0 }));
    summary[bk] = { docs: docs, projects: ps };
  });
  return {
    summary: summary,
    branches: Object.keys(branches).filter(vis).reduce((a, k) => { a[k] = branches[k]; return a; }, {}),
    projects: projects.filter(p => vis(p.branch)),
    files: files.filter(f => vis(f.branch))
  };
}
function projMissing_(p, files) {
  const m = [];
  ['name', 'status', 'owner', 'location', 'start', 'end', 'value', 'description'].forEach(k => { if (!p[k] || (k === 'value' && !Number(p[k]))) m.push(k); });
  if (files.filter(f => f.item === 'p:' + p.id).length < 3) m.push('photos');
  return m;
}

function saveBranch(bk, patch) {
  const u = U_; needBranch_(bk);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = branchSheet_();
    const cur = rows_(s, BCOLS).filter(r => r.branch === bk)[0] || { branch: bk, _row: -1 };
    ['phone', 'email', 'person', 'sceExpiry', 'crExpiry'].forEach(k => { if (patch[k] !== undefined) cur[k] = String(patch[k]).trim(); });
    cur.updatedAt = Date.now(); cur.updatedBy = u.name;
    put_(s, BCOLS, cur._row, cur);
    if (patch.phone !== undefined || patch.email !== undefined || patch.person !== undefined) {
      writeText_(itemFolder_(bk, 'contact'), 'بيانات التواصل.txt',
        'فرع: ' + BRANCHES[bk][1] + '\nرقم الهاتف: ' + (cur.phone || '') + '\nالبريد الإلكتروني: ' + (cur.email || '') + '\nالشخص المسؤول: ' + (cur.person || '') + '\n');
    }
    return strip_(cur);
  } finally { lock.releaseLock(); }
}

function storeFile_(folder, bk, item, name, mime, b64, u) {
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', clean_(name));
  const f = folder.createFile(blob);
  const row = { id: f.getId(), branch: bk, item: item, name: f.getName(), mime: mime, size: f.getSize(), url: f.getUrl(), at: Date.now(), by: u.name };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { put_(filesSheet_(), FCOLS, -1, row); } finally { lock.releaseLock(); }
  return row;
}
function uploadDoc(bk, ik, name, mime, b64) {
  const u = U_; needBranch_(bk);
  if (ik === 'contact') throw new Error('بيانات التواصل تُكتب في حقولها');
  return storeFile_(itemFolder_(bk, ik), bk, ik, name, mime, b64, u);
}
function uploadPhoto(pid, name, mime, b64) {
  const u = U_;
  const p = rows_(projSheet_(), PCOLS).filter(r => r.id === pid)[0];
  if (!p) throw new Error('احفظ المشروع أولاً');
  return storeFile_(DriveApp.getFolderById(p.folderId), p.branch, 'p:' + pid, name, mime, b64, u);
}
function deleteFile(fileId) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = filesSheet_();
    const r = rows_(s, FCOLS).filter(x => x.id === fileId)[0];
    if (!r) return true;
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
    s.deleteRow(r._row);
    return true;
  } finally { lock.releaseLock(); }
}

function saveProject(p) {
  const u = U_; needBranch_(p.branch);
  if (!String(p.name || '').trim()) throw new Error('اكتب اسم المشروع');
  if (!CATS[p.category]) throw new Error('اختر تصنيف المشروع');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = projSheet_();
    const old = p.id ? rows_(s, PCOLS).filter(r => r.id === p.id)[0] : null;
    const target = catFolder_(p.branch, p.category);
    let folder = null;
    if (old && old.folderId) { try { folder = DriveApp.getFolderById(old.folderId); } catch (e) {} }
    if (!folder) folder = target.createFolder(clean_(p.name));
    else {
      if (folder.getName() !== clean_(p.name)) folder.setName(clean_(p.name));
      const par = folder.getParents();
      if (par.hasNext() && par.next().getId() !== target.getId()) folder.moveTo(target);
    }
    const row = {
      id: old ? old.id : rid_(), branch: p.branch, category: p.category, status: STATUS[p.status] ? p.status : 'done',
      name: String(p.name).trim(), owner: p.owner || '', location: p.location || '', start: p.start || '', end: p.end || '',
      value: Number(String(p.value || '').replace(/[^\d.]/g, '')) || 0, description: p.description || '',
      folderId: folder.getId(), updatedAt: Date.now(), updatedBy: u.name
    };
    put_(s, PCOLS, old ? old._row : -1, row);
    writeText_(folder, 'بيانات المشروع.txt', [
      'اسم المشروع: ' + row.name,
      'الفرع: ' + BRANCHES[row.branch][1],
      'تصنيف المشروع: ' + CATS[row.category][1],
      'حالة المشروع: ' + STATUS[row.status],
      'الجهة المالكة / العميل: ' + row.owner,
      'موقع المشروع: ' + row.location,
      'تاريخ بداية المشروع: ' + row.start,
      'تاريخ نهاية المشروع: ' + row.end,
      'قيمة المشروع: ' + Number(row.value).toLocaleString('en-US') + ' ريال',
      'وصف المشروع: ' + row.description
    ].join('\n') + '\n');
    return row;
  } finally { lock.releaseLock(); }
}
function deleteProject(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = projSheet_();
    const p = rows_(s, PCOLS).filter(r => r.id === id)[0];
    if (!p) return true;
    try { DriveApp.getFolderById(p.folderId).setTrashed(true); } catch (e) {}
    s.deleteRow(p._row);
    const fs = filesSheet_();
    rows_(fs, FCOLS).filter(f => f.item === 'p:' + id).reverse().forEach(f => fs.deleteRow(f._row));
    return true;
  } finally { lock.releaseLock(); }
}

/* ---------- API ---------- */
function doPost(e) {
  let out;
  try {
    const b = JSON.parse(e.postData.contents);
    const API = {
      getAll: getAll, saveBranch: saveBranch, uploadDoc: uploadDoc, uploadPhoto: uploadPhoto, deleteFile: deleteFile,
      saveProject: saveProject, deleteProject: deleteProject
    };
    if (!API[b.fn]) throw new Error('طلب غير معروف');
    const data = API[b.fn].apply(null, b.args || []);
    out = { ok: true, data: data };
  } catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) {
  // الرابط نفسه يفتح اللوحة مباشرة — ?b=jeddah يفتح على صفحة الفرع
  const b = (e && e.parameter && e.parameter.b) || '';
  const html = HtmlService.createHtmlOutputFromFile('Index').getContent()
    .replace('<head>', '<head><script>window.__START=' + JSON.stringify(BRANCHES[b] ? b : '') + ';</script>');
  return HtmlService.createHtmlOutput(html).setTitle('منصة معتمد')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- تشغيل مرة واحدة من المحرر ---------- */
function init() {
  branchSheet_(); projSheet_(); filesSheet_();
  const d = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('ورقة1');
  if (d && ss_().getSheets().length > 1) ss_().deleteSheet(d);
  buildFolders_();
  setupWeeklyBackup();
  return 'تم';
}
function weeklyBackup() {
  const name = 'نسخة احتياطية - منصة معتمد - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  DriveApp.getFileById(SHEET_ID).makeCopy(name, root_());
}
function setupWeeklyBackup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'weeklyBackup').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(16).create();
  return 'تم';
}
