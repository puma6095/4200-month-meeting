/*******************************************************
 * 奇美醫院 教學部｜各項教學業務報告系統  後端主程式
 * 部署方式：Apps Script 專案 → 部署為網頁應用程式
 *   執行身分：我（承辦人帳號）
 *   具有存取權的使用者：所有人
 * 前端以人事號辨識身分，不需 Google 帳號登入。
 *******************************************************/

const CFG = {
  SHEET_ID   : '1CuaEKZXHKRJHSlAOKWyDPx_d8iLOUruwQNdvwbOTX7k',
  SH_STAFF   : '人員對照表',
  SH_REPORT  : '月報內容',
  SH_UNIT    : '單位對照表',
  SH_TRACK   : '追蹤事項',
  SH_REQ     : '新增業務申請',      // 首次執行自動建立
  SH_LOG     : '寄送紀錄',          // 首次執行自動建立
  FOLDER     : '教學部月報附件',    // 雲端硬碟資料夾，自動建立
  MAIL_DOMAIN: 'chimei.org.tw',
  MEETING_DAY: 15,                  // 每月部務會議日（可於指令碼屬性覆寫）
  REMIND_DAYS: [5, 2],              // 會議前幾天寄催繳
  TITLE      : '教學部各項教學業務報告'
};

/* 組別呈現順序：與說明分頁一致 */
const UNIT_ORDER = [
  '醫學教育委員會','教學部','專責教學型醫事人員','教師培育中心','住院醫師訓練組','實習醫學生訓練組',
  '畢業後一般醫學訓練','一般醫學內科示範中心','一般醫學內科示範中心-社區訓練及師資培育',
  '臨床技能中心','OSCE小組','醫事人員教育訓練小組','臨床醫事人員培訓計畫',
  '台灣醫事聯合臨床技能發展學會','奇美伊婆婆系統開發進度','病歷審查小組',
  '視聽中心','圖書館'
];

/* ===================================================================
 * 一、基礎工具
 * =================================================================*/

/* 單次執行內的快取：同一張表只讀一次，大幅減少往返 */
var _M = {};
function ss_() {
  if (!_M.ss) _M.ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  return _M.ss;
}
function invalidate_(name) {
  delete _M['t:' + name];
  delete _M['h:' + name];
  delete _M.people;
  delete _M.quarter;
  delete _M.units;
}

function sheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh && headers) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#0d5c60').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('找不到工作表：' + name);
  return sh;
}

/** 讀整張表，回傳 [{欄名: 值}]，並附 _row 實際列號 */
function readTable_(name) {
  if (_M['t:' + name]) return _M['t:' + name];
  const sh = sheet_(name);
  const rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  const head = rng[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < rng.length; i++) {
    const row = rng[i];
    if (row.every(c => c === '' || c === null)) continue;
    const o = { _row: i + 1 };
    head.forEach((h, j) => { if (h) o[h] = row[j]; });
    out.push(o);
  }
  _M['t:' + name] = out;
  return out;
}

function headerIndex_(name) {
  if (_M['h:' + name]) return _M['h:' + name];
  const sh = sheet_(name);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  head.forEach((h, i) => { const k = String(h).trim(); if (k) map[k] = i + 1; });
  _M['h:' + name] = map;
  return map;
}

function fmtDate_(d) {
  const p = n => ('0' + n).slice(-2);
  let y = d.getFullYear();
  if (y > 1911) y = y - 1911;      // 西元 → 民國；若已是民國年（被試算表誤判）則原樣保留
  return y + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate())
       + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
}
function s_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return fmtDate_(v);
  return String(v).trim();
}
function upper_(v) { return s_(v).toUpperCase(); }

/** 負責人員欄可寫多位，以 、,，/ 或空白分隔 */
function owners_(v) {
  return s_(v).split(/[、,，\/\s　]+/).map(s_).filter(String);
}
function isOwner_(v, name) { return owners_(v).indexOf(s_(name)) >= 0; }

/** 回覆欄格式：「姓名：內容」以 ｜ 分段；舊資料（無姓名）視為共用段 */
function parseReply_(v) {
  const txt = s_(v);
  if (!txt) return [];
  return txt.split('｜').map(seg => {
    const m = seg.match(/^\s*([^：:\n]{1,10})[：:]([\s\S]*)$/);
    if (m) return { who: s_(m[1]), text: s_(m[2]) };
    return { who: '', text: s_(seg) };
  }).filter(x => x.text);
}
function buildReply_(list) {
  return list.filter(x => s_(x.text))
             .map(x => (x.who ? x.who + '：' : '') + s_(x.text))
             .join('｜');
}
function replyOf_(v, name) {
  const list = parseReply_(v);
  for (let i = 0; i < list.length; i++) {
    if (list[i].who === s_(name)) return list[i].text;
  }
  if (list.length === 1 && !list[0].who) return list[0].text;
  return '';
}
function setReply_(v, name, text) {
  const list = parseReply_(v);
  let hit = false;
  for (let i = 0; i < list.length; i++) {
    if (list[i].who === s_(name)) { list[i].text = s_(text); hit = true; break; }
    if (!list[i].who && list.length === 1) { list[i].who = s_(name); list[i].text = s_(text); hit = true; break; }
  }
  if (!hit) list.push({ who: s_(name), text: s_(text) });
  return buildReply_(list);
}

function ym_(d) {
  d = d || new Date();
  return (d.getFullYear() - 1911) + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
}

function nowStr_() {
  const d = new Date();
  const p = n => ('0' + n).slice(-2);
  return (d.getFullYear() - 1911) + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate())
       + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
}

function meetingDay_() {
  const v = PropertiesService.getScriptProperties().getProperty('MEETING_DAY');
  return v ? Number(v) : CFG.MEETING_DAY;
}

/* ===================================================================
 * 二、名冊與權限
 * =================================================================*/

/** 人員對照表：首次使用時自動補上「業務狀態」欄 */
function ensureStaffCols_() {
  const sh = sheet_(CFG.SH_STAFF);
  let idx = headerIndex_(CFG.SH_STAFF);
  if (!idx['業務狀態']) {
    const col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue('業務狀態')
      .setBackground('#0d5c60').setFontColor('#ffffff').setFontWeight('bold');
    invalidate_(CFG.SH_STAFF);
    idx = headerIndex_(CFG.SH_STAFF);
  }
  return idx;
}
function ended_(r) { return s_(r['業務狀態']) === '已結束'; }

/** 回傳 { 人事號: {hr,name,mail,role,own:[],view:[],quarterUnits:[]} } */
function loadPeople_() {
  if (_M.people) return _M.people;
  const people = {};

  readTable_(CFG.SH_STAFF).forEach(r => {
    const hr = upper_(r['人事號']);
    if (!hr) return;
    if (ended_(r)) return;
    const role = s_(r['權限']) || '一般';
    const unit = s_(r['組別']);
    if (!people[hr]) {
      people[hr] = { hr: hr, name: s_(r['姓名']), mail: s_(r['電子郵件']),
                     role: role, own: [], view: [] };
    }
    const p = people[hr];
    if (!p.mail && s_(r['電子郵件'])) p.mail = s_(r['電子郵件']);
    if (role === '主管') { if (unit && p.view.indexOf(unit) < 0) p.view.push(unit); }
    else if (unit && p.own.indexOf(unit) < 0) p.own.push(unit);
    if (role === '管理者') p.role = '管理者';
  });

  // 每季報告單位（視聽中心、圖書館）併入同一份名冊
  readTable_(CFG.SH_UNIT).forEach(r => {
    const hr = upper_(r['報告主管人事號']);
    const unit = s_(r['單位名稱']);
    if (!hr || !unit) return;
    if (!people[hr]) {
      people[hr] = { hr: hr, name: s_(r['姓名']) || unit, mail: s_(r['電子郵件']),
                     role: '一般', own: [], view: [] };
    }
    if (people[hr].own.indexOf(unit) < 0) people[hr].own.push(unit);
  });

  Object.keys(people).forEach(hr => {
    const p = people[hr];
    if (!p.mail) p.mail = hr + '@' + CFG.MAIL_DOMAIN;   // 未填時以人事號自動組出
    p.own.sort((a, b) => unitIdx_(a) - unitIdx_(b));
  });
  _M.people = people;
  return people;
}

/** 每季報告單位設定 { 單位名稱: [月份...] } */
function loadQuarter_() {
  if (_M.quarter) return _M.quarter;
  const q = {};
  readTable_(CFG.SH_UNIT).forEach(r => {
    const unit = s_(r['單位名稱']);
    if (!unit) return;
    q[unit] = s_(r['報告月份']).replace(/，/g, ',').split(',')
                .map(x => Number(x.trim())).filter(n => n >= 1 && n <= 12);
  });
  _M.quarter = q;
  return q;
}

/** 本月是否需要報告 */
function unitActive_(unit, quarter, month) {
  if (!quarter[unit] || !quarter[unit].length) return true;
  return quarter[unit].indexOf(month) >= 0;
}

/** 完整組別清單：內建順序 + 名冊中新核准的組別（接在最後） */
function allUnits_() {
  if (_M.units) return _M.units;
  const out = UNIT_ORDER.slice();
  const add = u => { u = s_(u); if (u && out.indexOf(u) < 0) out.push(u); };
  try { readTable_(CFG.SH_STAFF).forEach(r => { if (!ended_(r)) add(r['組別']); }); } catch (e) {}
  try { readTable_(CFG.SH_UNIT).forEach(r => add(r['單位名稱'])); } catch (e) {}
  _M.units = out;
  return out;
}
function unitIdx_(u) {
  const i = allUnits_().indexOf(u);
  return i < 0 ? 9999 : i;
}

/** 依名冊算出各組主責（維持呈現順序） */
function ownersByUnit_(people) {
  const map = {};
  allUnits_().forEach(u => { map[u] = []; });
  Object.keys(people).forEach(hr => {
    people[hr].own.forEach(u => {
      if (!map[u]) map[u] = [];
      if (map[u].indexOf(hr) < 0) map[u].push(hr);
    });
  });
  Object.keys(map).forEach(u => {
    map[u].sort((a, b) => people[a].name.localeCompare(people[b].name, 'zh-Hant'));
  });
  return map;
}

/* ===================================================================
 * 三、月報內容存取
 * =================================================================*/

function loadReports_(ym) {
  return loadReportsMulti_([ym])[ym] || {};
}

/** 一次讀取多個年月，避免重複掃描整張表 */
function loadReportsMulti_(yms) {
  const want = {};
  yms.forEach(y => { want[y] = {}; });
  readTable_(CFG.SH_REPORT).forEach(r => {
    const y = s_(r['年月']);
    if (!want[y]) return;
    const key = s_(r['組別']) + '|' + upper_(r['主責人事號']);
    want[y][key] = {
      row     : r._row,
      unit    : s_(r['組別']),
      hr      : upper_(r['主責人事號']),
      name    : s_(r['主責姓名']),
      item    : s_(r['業務項目']),
      text    : s_(r['本月業務內容']),
      status  : s_(r['狀態']) || '草稿',
      fileName: s_(r['附件檔名']),
      fileUrl : s_(r['附件連結']),
      co      : s_(r['協同人事號']).split(',').map(upper_).filter(String),
      coName  : s_(r['協同人員']),
      at      : s_(r['最後更新'])
    };
  });
  return want;
}

function writeReport_(ym, unit, hr, patch) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh  = sheet_(CFG.SH_REPORT);
    const idx = headerIndex_(CFG.SH_REPORT);
    const cur = loadReports_(ym);
    const key = unit + '|' + upper_(hr);
    const people = loadPeople_();
    const person = people[upper_(hr)] || { name: '' };
    const old = cur[key] || {};

    const rec = {
      '年月'          : ym,
      '組別'          : unit,
      '主責人事號'    : upper_(hr),
      '主責姓名'      : person.name,
      '業務項目'      : patch.item     !== undefined ? patch.item     : (old.item || ''),
      '本月業務內容'  : patch.text     !== undefined ? patch.text     : (old.text || ''),
      '狀態'          : patch.status   !== undefined ? patch.status   : (old.status || '草稿'),
      '附件檔名'      : patch.fileName !== undefined ? patch.fileName : (old.fileName || ''),
      '附件連結'      : patch.fileUrl  !== undefined ? patch.fileUrl  : (old.fileUrl || ''),
      '協同人事號'    : patch.co       !== undefined ? patch.co.join(',') : (old.co || []).join(','),
      '協同人員'      : '',
      '填報人事號'    : upper_(patch.by || hr),
      '最後更新'      : nowStr_()
    };
    rec['協同人員'] = rec['協同人事號'].split(',').filter(String)
        .map(id => (people[upper_(id)] || {}).name || id).join(',');

    // 整列一次寫入，避免逐格呼叫造成延遲
    const width = sh.getLastColumn();
    const row = old.row || (sh.getLastRow() + 1);
    let line;
    if (old.row) {
      line = sh.getRange(row, 1, 1, width).getValues()[0];
    } else {
      line = new Array(width).fill('');
    }
    Object.keys(rec).forEach(k => { if (idx[k]) line[idx[k] - 1] = rec[k]; });
    if (idx['最後更新']) sh.getRange(row, idx['最後更新']).setNumberFormat('@');
    if (idx['年月'])     sh.getRange(row, idx['年月']).setNumberFormat('@');
    sh.getRange(row, 1, 1, width).setValues([line]);
    invalidate_(CFG.SH_REPORT);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/* ===================================================================
 * 四、前端進入點
 * =================================================================*/

function doGet(e) {
  return HtmlService.createHtmlOutput(PAGE_HTML())
    .setTitle(CFG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================================================================
 * 五、對前端開放的 API（前端以 google.script.run 呼叫）
 * =================================================================*/

/** 登入並取回該人所需的全部資料 */
function api_bootstrap(hrRaw) {
  const hr = upper_(hrRaw);
  const people = loadPeople_();
  if (!people[hr]) return { ok: false, msg: '查無此人事號，請確認後再輸入。' };

  const me      = people[hr];
  const month   = new Date().getMonth() + 1;
  const ym      = ym_();
  const prevYm  = prevYm_(ym);
  const quarter = loadQuarter_();
  const owners  = ownersByUnit_(people);
  const both    = loadReportsMulti_([ym, prevYm]);
  const cur     = both[ym]     || {};
  const prev    = both[prevYm] || {};

  // 我主責的區塊
  const mine = me.own.map(u => {
    const k = u + '|' + hr;
    const p = prev[k];
    return {
      unit    : u,
      active  : unitActive_(u, quarter, month),
      months  : quarter[u] || null,
      multi   : (owners[u] || []).length > 1,
      block   : cur[k]  ? stripRow_(cur[k])  : null,
      prevText: p ? p.text : '',
      prevCo  : p ? p.co : []
    };
  });

  // 我被列為協同、可檢視的區塊
  const asCo = [];
  Object.keys(cur).forEach(k => {
    const b = cur[k];
    if (b.hr === hr) return;
    if (b.co.indexOf(hr) < 0) return;
    if (!unitActive_(b.unit, quarter, month)) return;
    if (b.status === '草稿') return;              // 草稿不外流
    asCo.push(stripRow_(b));
  });

  // 指派給我的追蹤事項
  const ymKey = ym.replace('-', '/');
  const tracks = readTable_(CFG.SH_TRACK)
    .filter(r => {
      if (!isOwner_(r['負責人員'], me.name)) return false;
      if (s_(r['狀態']) !== '已結案') return true;
      return s_(r['結案時間']).indexOf(ymKey) === 0;   // 本月結案的仍可補回覆
    })
    .map(r => ({ id: r['編號'], from: s_(r['來源會議']), q: s_(r['事項內容']),
                 owners: owners_(r['負責人員']),
                 reply: replyOf_(r['回覆執行進度'], me.name),
                 others: parseReply_(r['回覆執行進度']).filter(x => x.who && x.who !== me.name),
                 files: trackFiles_(r),
                 status: s_(r['狀態']) || '待回覆' }));

  return {
    ok: true,
    me: { hr: hr, name: me.name, role: me.role, own: me.own, view: me.view },
    ym: ym, month: month,
    units: allUnits_().filter(u => unitActive_(u, quarter, month)),
    mine: mine, asCo: asCo, tracks: tracks,
    myRequests: myRequests_(hr),
    people: Object.keys(people).map(function(k){
      return { hr: k, name: people[k].name, role: people[k].role };
    })
  };
}

function stripRow_(b) { const c = Object.assign({}, b); delete c.row; return c; }

/** 組出單一區塊在前端所需的完整資料 */
function mineEntry_(hr, unit) {
  const month   = new Date().getMonth() + 1;
  const ym      = ym_();
  const both    = loadReportsMulti_([ym, prevYm_(ym)]);
  const cur     = both[ym] || {};
  const prev    = both[prevYm_(ym)] || {};
  const quarter = loadQuarter_();
  const owners  = ownersByUnit_(loadPeople_());
  const k = unit + '|' + upper_(hr);
  const p = prev[k];
  return {
    unit    : unit,
    active  : unitActive_(unit, quarter, month),
    months  : quarter[unit] || null,
    multi   : (owners[unit] || []).length > 1,
    block   : cur[k] ? stripRow_(cur[k]) : null,
    prevText: p ? p.text : '',
    prevCo  : p ? p.co : []
  };
}

function prevYm_(ym) {
  const p = ym.split('-');
  let y = Number(p[0]), m = Number(p[1]) - 1;
  if (m === 0) { m = 12; y -= 1; }
  return y + '-' + ('0' + m).slice(-2);
}

/** 暫存或送出 */
function api_saveBlock(hrRaw, unit, item, text, co, submit) {
  const hr = upper_(hrRaw);
  const people = loadPeople_();
  if (!people[hr] || people[hr].own.indexOf(unit) < 0) {
    return { ok: false, msg: '你沒有這個組別的填報權限。' };
  }
  if (!s_(text)) return { ok: false, msg: '請先填寫本月業務內容。' };
  writeReport_(ym_(), unit, hr, {
    item: s_(item), text: s_(text), co: (co || []).map(upper_),
    status: submit ? '已送出' : '草稿', by: hr
  });
  return { ok: true, entry: mineEntry_(hr, unit),
           msg: submit ? '已送出，主管與投影畫面都看得到了。' : '已暫存，只有你看得到。' };
}

/** 標記本月無業務 */
function api_markNone(hrRaw, unit) {
  const hr = upper_(hrRaw);
  const people = loadPeople_();
  if (!people[hr] || people[hr].own.indexOf(unit) < 0) return { ok: false, msg: '權限不足。' };
  writeReport_(ym_(), unit, hr, { item: '', text: '', status: '本月無業務', by: hr });
  return { ok: true, entry: mineEntry_(hr, unit), msg: '已標記本月無業務，不會列入催繳。' };
}

/** 收回（已送出 → 草稿） */
function api_unsend(hrRaw, unit) {
  const hr = upper_(hrRaw);
  writeReport_(ym_(), unit, hr, { status: '草稿', by: hr });
  return { ok: true, entry: mineEntry_(hr, unit), msg: '已收回，目前是暫存狀態，主管看不到。' };
}

/** 上傳附件（前端傳 base64） */
function api_upload(hrRaw, unit, fileName, mimeType, b64) {
  const hr = upper_(hrRaw);
  const people = loadPeople_();
  if (!people[hr] || people[hr].own.indexOf(unit) < 0) return { ok: false, msg: '權限不足。' };

  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, fileName);
  const folder = subFolder_(ym_() + '／' + unit);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const cur = loadReports_(ym_())[unit + '|' + hr] || {};
  const names = s_(cur.fileName) ? cur.fileName.split('|') : [];
  const urls  = s_(cur.fileUrl)  ? cur.fileUrl.split('|')  : [];
  if (names.length >= 5) { file.setTrashed(true); return { ok: false, msg: '每個區塊最多 5 個附件。' }; }
  names.push(fileName);
  urls.push(file.getUrl());
  writeReport_(ym_(), unit, hr, { fileName: names.join('|'), fileUrl: urls.join('|'), by: hr });
  return { ok: true, entry: mineEntry_(hr, unit) };
}

function api_deleteFile(hrRaw, unit, index) {
  const hr = upper_(hrRaw);
  const cur = loadReports_(ym_())[unit + '|' + hr];
  if (!cur) return { ok: false, msg: '找不到資料。' };
  const names = s_(cur.fileName) ? cur.fileName.split('|') : [];
  const urls  = s_(cur.fileUrl)  ? cur.fileUrl.split('|')  : [];
  names.splice(index, 1); urls.splice(index, 1);
  writeReport_(ym_(), unit, hr, { fileName: names.join('|'), fileUrl: urls.join('|'), by: hr });
  return { ok: true, entry: mineEntry_(hr, unit) };
}

function subFolder_(path) {
  let root;
  const it = DriveApp.getFoldersByName(CFG.FOLDER);
  root = it.hasNext() ? it.next() : DriveApp.createFolder(CFG.FOLDER);
  path.split('／').forEach(seg => {
    const f = root.getFoldersByName(seg);
    root = f.hasNext() ? f.next() : root.createFolder(seg);
  });
  return root;
}

/** 回覆追蹤事項 */
function api_saveTrack(hrRaw, id, reply) {
  const hr = upper_(hrRaw);
  const people = loadPeople_();
  const me = people[hr];
  if (!me) return { ok: false, msg: '權限不足。' };
  if (!s_(reply)) return { ok: false, msg: '請填寫回覆執行進度後再儲存。' };

  const sh = sheet_(CFG.SH_TRACK);
  const idx = headerIndex_(CFG.SH_TRACK);
  const rows = readTable_(CFG.SH_TRACK);
  const t = rows.filter(r => String(r['編號']) === String(id))[0];
  if (!t) return { ok: false, msg: '找不到該事項。' };
  if (!isOwner_(t['負責人員'], me.name) && me.role !== '管理者') {
    return { ok: false, msg: '這不是指派給你的事項。' };
  }
  sh.getRange(t._row, idx['回覆執行進度'])
    .setValue(setReply_(t['回覆執行進度'], me.name, reply));
  invalidate_(CFG.SH_TRACK);
  return { ok: true, msg: '已儲存回覆。' };
}

/** 結案／重新開啟：僅管理者 */
function api_closeTrack(hrRaw, id, close) {
  if (!isAdmin_(hrRaw)) return { ok: false, msg: '權限不足。' };
  const sh = sheet_(CFG.SH_TRACK);
  const idx = headerIndex_(CFG.SH_TRACK);
  const t = readTable_(CFG.SH_TRACK).filter(r => String(r['編號']) === String(id))[0];
  if (!t) return { ok: false, msg: '找不到該事項。' };
  if (close) {
    const miss = owners_(t['負責人員']).filter(w => !replyOf_(t['回覆執行進度'], w));
    if (miss.length === owners_(t['負責人員']).length) {
      return { ok: false, msg: '這一項還沒有回覆內容，結案後主管會看不到執行進度。請先請負責人回覆，或直接在試算表補上。' };
    }
  }
  sh.getRange(t._row, idx['狀態']).setValue(close ? '已結案' : '待回覆');
  if (idx['結案時間']) {
    sh.getRange(t._row, idx['結案時間']).setNumberFormat('@').setValue(close ? nowStr_() : '');
  }
  invalidate_(CFG.SH_TRACK);
  return { ok: true, msg: close ? '已結案。' : '已重新開啟，將持續追蹤至結案。' };
}

/* ===================================================================
 * 六、主管檢視
 * =================================================================*/

function api_overview(hrRaw) {
  const hr = upper_(hrRaw);
  const people = loadPeople_();
  const me = people[hr];
  if (!me) return { ok: false, msg: '權限不足。' };
  const isBoss = (me.role === '主管' || me.role === '管理者');

  const month = new Date().getMonth() + 1;
  const quarter = loadQuarter_();
  const owners = ownersByUnit_(people);
  const cur = loadReports_(ym_());

  let scope;
  if (!isBoss) {
    scope = allUnits_().slice();                  // 一般同仁：投影內容全部可看
  } else if (me.role === '管理者' || !me.view.length) {
    scope = allUnits_().slice();
  } else {
    scope = me.view.slice();
  }
  scope = scope.filter(u => unitActive_(u, quarter, month));

  const data = scope.map(u => ({
    unit: u,
    owners: (owners[u] || []).map(id => {
      const b = cur[u + '|' + id];
      const sent = b && b.status === '已送出';
      const none = b && b.status === '本月無業務';
      const coNames = (sent && b.co && b.co.length)
        ? b.co.map(cid => (people[upper_(cid)] || {}).name || cid).filter(String)
        : [];
      return {
        hr: id, name: people[id].name,
        status: sent ? '已送出' : (none ? '本月無業務' : '尚未送出'),
        item: sent ? b.item : '', text: sent ? b.text : '',
        co: coNames,
        files: sent ? fileList_(b) : [], at: b ? b.at : ''
      };
    })
  }));
  const ymNow = ym_();
  const tracks = readTable_(CFG.SH_TRACK).filter(r => {
    const st = s_(r['狀態']);
    if (st !== '已結案') return true;
    return s_(r['結案時間']).indexOf(ymNow.replace('-', '/')) === 0;   // 本月才結案的仍列出
  }).map(r => ({
    q: s_(r['事項內容']), owner: owners_(r['負責人員']).join('、'),
    replies: parseReply_(r['回覆執行進度']),
    files: trackFiles_(r),
    reply: s_(r['回覆執行進度']), status: s_(r['狀態']) || '待回覆',
    from: s_(r['來源會議'])
  }));

  return { ok: true, ym: ymNow, data: data, tracks: tracks, preview: false };
}

function fileList_(b) {
  const names = s_(b.fileName) ? b.fileName.split('|') : [];
  const urls  = s_(b.fileUrl)  ? b.fileUrl.split('|')  : [];
  return names.map((n, i) => ({ n: n, url: urls[i] || '' }));
}

/* ===================================================================
 * 七、管理者：進度、催繳、申請審核
 * =================================================================*/

function api_admin(hrRaw) {
  const hr = upper_(hrRaw);
  const people = loadPeople_();
  if (!people[hr] || people[hr].role !== '管理者') return { ok: false, msg: '權限不足。' };

  const month = new Date().getMonth() + 1;
  const quarter = loadQuarter_();
  const owners = ownersByUnit_(people);
  const cur = loadReports_(ym_());
  const noted = notedMap_(ym_());

  let total = 0, fin = 0;
  const rows = [];
  allUnits_().forEach(u => {
    if (!unitActive_(u, quarter, month)) {
      rows.push({ unit: u, inactive: true, months: quarter[u] || [],
                  owners: (owners[u] || []).map(id => people[id].name) });
      return;
    }
    (owners[u] || []).forEach(id => {
      const b = cur[u + '|' + id];
      const st = b ? b.status : '未填報';
      total++;
      if (st === '已送出' || st === '本月無業務') fin++;
      rows.push({ unit: u, hr: id, name: people[id].name, item: b ? b.item : '',
                  status: st === '草稿' ? '暫存中' : st, at: b ? b.at : '',
                  noted: noted[u + '|' + id] || '' });
    });
  });

  return {
    ok: true, ym: ym_(), total: total, fin: fin,
    rows: rows,
    tracks: readTable_(CFG.SH_TRACK).map(r => ({
      id: r['編號'], q: s_(r['事項內容']), owner: owners_(r['負責人員']).join('、'),
      replies: parseReply_(r['回覆執行進度']),
      files: trackFiles_(r),
      reply: s_(r['回覆執行進度']), status: s_(r['狀態']) })),
    requests: allRequests_()
  };
}

/* --- 催繳 --- */

function api_remindOne(hrRaw, unit, targetHr) {
  if (!isAdmin_(hrRaw)) return { ok: false, msg: '權限不足。' };
  const people = loadPeople_();
  const t = people[upper_(targetHr)];
  if (!t) return { ok: false, msg: '找不到該同仁。' };
  const tkAll = pendingTracks_();
  sendReminder_(t, [unit], tkAll[t.name] || []);
  logSend_(ym_(), unit, t.hr, '個別催繳');
  return { ok: true, msg: '已寄出催繳信給 ' + t.name + '（' + unit + '）。' };
}

function api_remindAll(hrRaw) {
  if (!isAdmin_(hrRaw)) return { ok: false, msg: '權限不足。' };
  const n = runReminder_('手動催繳');
  return { ok: n >= 0, msg: n ? ('已寄出催繳信給 ' + n + ' 位同仁。') : '目前所有區塊都已完成，沒有需要催繳的對象。' };
}

/** 找出未完成者並寄信；回傳寄出人數 */
function runReminder_(tag) {
  const people = loadPeople_();
  const quarter = loadQuarter_();
  const owners = ownersByUnit_(people);
  const month = new Date().getMonth() + 1;
  const cur = loadReports_(ym_());

  const pending = {};   // hr -> [unit...]
  allUnits_().forEach(u => {
    if (!unitActive_(u, quarter, month)) return;
    (owners[u] || []).forEach(id => {
      const b = cur[u + '|' + id];
      const st = b ? b.status : '';
      if (st !== '已送出' && st !== '本月無業務') {
        (pending[id] = pending[id] || []).push(u);
      }
    });
  });

  // 尚未回覆的追蹤事項也納入催繳對象
  const tk = pendingTracks_();
  const byName = {};
  Object.keys(people).forEach(id => { byName[people[id].name] = id; });
  const tkOf = {};
  Object.keys(tk).forEach(nm2 => {
    const id = byName[nm2];
    if (id) tkOf[id] = tk[nm2];
  });

  const ids = [];
  Object.keys(pending).forEach(id => { if (ids.indexOf(id) < 0) ids.push(id); });
  Object.keys(tkOf).forEach(id => { if (ids.indexOf(id) < 0) ids.push(id); });

  ids.forEach(id => {
    sendReminder_(people[id], pending[id] || [], tkOf[id] || []);
    (pending[id] || []).forEach(u => logSend_(ym_(), u, id, tag));
    if (!pending[id]) logSend_(ym_(), '（追蹤事項）', id, tag);
  });
  return ids.length;
}

/** 尚未回覆的追蹤事項（未結案且回覆為空），以姓名為鍵 */
function pendingTracks_() {
  const map = {};
  readTable_(CFG.SH_TRACK).forEach(r => {
    if (s_(r['狀態']) === '已結案') return;
    owners_(r['負責人員']).forEach(who => {
      if (replyOf_(r['回覆執行進度'], who)) return;
      (map[who] = map[who] || []).push(s_(r['事項內容']));
    });
  });
  return map;
}

function sendReminder_(person, units, tracks) {
  const url = ScriptApp.getService().getUrl();
  units  = units  || [];
  tracks = tracks || [];
  let body = person.name + ' 您好：\n\n本月（' + ym_() + '）部務會議尚有以下項目未完成：\n';

  if (units.length) {
    body += '\n【業務報告尚未送出】\n' + units.map(u => '　・' + u).join('\n') + '\n';
  }
  if (tracks.length) {
    body += '\n【上次會議追蹤事項尚未回覆】\n' +
            tracks.map(q => '　・' + (q.length > 60 ? q.substring(0, 60) + '…' : q)).join('\n') + '\n';
  }

  body += '\n請於部務會議前進入系統完成：\n' + url + '\n\n' +
          '業務報告若本月確實沒有可報事項，請按「本月無業務」；\n' +
          '追蹤事項請填寫「回覆執行進度」後按「儲存回覆」，已完成者可一併勾選結案。\n\n' +
          '— 教學部各項教學業務報告系統（系統自動發送，請勿直接回覆）';

  const subj = units.length && tracks.length ? ' 業務報告與追蹤事項尚未完成'
             : (units.length ? ' 業務報告尚未送出' : ' 追蹤事項尚未回覆');
  MailApp.sendEmail({ to: person.mail, subject: '【教學部】' + ym_() + subj, body: body });
}

/** 通知主管本月已彙整完成 */
function api_notifyManagers(hrRaw) {
  if (!isAdmin_(hrRaw)) return { ok: false, msg: '權限不足。' };
  const people = loadPeople_();
  const url = ScriptApp.getService().getUrl();
  const mgrs = Object.keys(people).filter(k => people[k].role === '主管');
  mgrs.forEach(k => {
    MailApp.sendEmail({
      to: people[k].mail,
      subject: '【教學部】' + ym_() + ' 各項教學業務報告已彙整完成',
      body: people[k].name + ' 您好：\n\n本月各組業務報告已彙整完成，可直接進入系統檢視：\n' + url +
            '\n\n— 教學部各項教學業務報告系統（系統自動發送，請勿直接回覆）'
    });
  });
  return { ok: true, msg: '已通知 ' + mgrs.length + ' 位主管。' };
}

function isAdmin_(hrRaw) {
  const p = loadPeople_()[upper_(hrRaw)];
  return !!(p && p.role === '管理者');
}

function logSend_(ym, unit, hr, kind) {
  const sh = sheet_(CFG.SH_LOG, ['年月', '組別', '人事號', '類型', '時間']);
  sh.appendRow([ym, unit, hr, kind, nowStr_()]);
  sh.getRange(sh.getLastRow(), 5).setNumberFormat('@').setValue(nowStr_());
  invalidate_(CFG.SH_LOG);
}

function notedMap_(ym) {
  const out = {};
  try {
    readTable_(CFG.SH_LOG).forEach(r => {
      if (s_(r['年月']) !== ym) return;
      out[s_(r['組別']) + '|' + upper_(r['人事號'])] = s_(r['時間']);
    });
  } catch (e) {}
  return out;
}

/* --- 新增業務申請 --- */

function reqSheet_() {
  return sheet_(CFG.SH_REQ, ['編號', '送出時間', '申請人事號', '申請人', '組別', '是否新組別', '說明', '狀態', '處理時間']);
}

function api_submitRequest(hrRaw, unit, isNew, why) {
  const hr = upper_(hrRaw);
  const people = loadPeople_();
  const me = people[hr];
  if (!me) return { ok: false, msg: '權限不足。' };
  unit = s_(unit);
  if (!unit) return { ok: false, msg: '請填寫要新增的業務名稱。' };
  if (me.own.indexOf(unit) >= 0) return { ok: false, msg: '你已經有「' + unit + '」這項業務了。' };

  const st = sheet_(CFG.SH_STAFF);
  const si = ensureStaffCols_();

  // 若先前結束過同一項，直接復原該列
  const exist = readTable_(CFG.SH_STAFF).filter(x =>
    upper_(x['人事號']) === hr && s_(x['組別']) === unit && ended_(x))[0];
  if (exist) {
    st.getRange(exist._row, si['業務狀態']).setValue('');
    invalidate_(CFG.SH_STAFF);
    return { ok: true, msg: '已重新啟用「' + unit + '」，可以繼續填報。' };
  }

  const row = [];
  row[si['人事號'] - 1] = hr;
  row[si['姓名'] - 1]   = me.name;
  if (si['電子郵件']) row[si['電子郵件'] - 1] = me.mail;
  row[si['權限'] - 1]   = me.role === '管理者' ? '管理者' : '一般';
  row[si['組別'] - 1]   = unit;
  if (si['備註']) row[si['備註'] - 1] = '同仁自行新增 ' + nowStr_();
  st.appendRow(row);
  invalidate_(CFG.SH_STAFF);

  // 留存紀錄，供管理者檢視或移除
  const sh = reqSheet_();
  const id = sh.getLastRow();
  sh.appendRow([id, nowStr_(), hr, me.name, unit, isNew ? '是' : '否', s_(why), '已新增', nowStr_()]);
  const ri = headerIndex_(CFG.SH_REQ);
  const rr = sh.getLastRow();
  if (ri['送出時間']) sh.getRange(rr, ri['送出時間']).setNumberFormat('@').setValue(nowStr_());
  if (ri['處理時間']) sh.getRange(rr, ri['處理時間']).setNumberFormat('@').setValue(nowStr_());
  invalidate_(CFG.SH_REQ);

  // 通知管理者（僅告知，不需審核）
  const url = ScriptApp.getService().getUrl();
  Object.keys(people).filter(k => people[k].role === '管理者').forEach(k => {
    MailApp.sendEmail({
      to: people[k].mail,
      subject: '【教學部】' + me.name + ' 新增業務－' + unit,
      body: me.name + '（' + hr + '）已自行新增業務：' + unit +
            (s_(why) ? '\n說明：' + s_(why) : '') +
            '\n\n此項已立即生效並納入本月填報與部務會議報告。\n' +
            '若需移除，請至系統「彙整與催繳」頁的新增業務紀錄操作。\n' + url +
            '\n\n— 教學部各項教學業務報告系統'
    });
  });
  return { ok: true, msg: '已新增「' + unit + '」，可以直接開始填報。' };
}

/** 追蹤事項附件：首次使用時自動補上欄位 */
function ensureTrackCols_() {
  const sh = sheet_(CFG.SH_TRACK);
  let idx = headerIndex_(CFG.SH_TRACK);
  const need = ['附件檔名', '附件連結'];
  let col = sh.getLastColumn();
  let added = false;
  need.forEach(h => {
    if (!idx[h]) {
      col++;
      sh.getRange(1, col, 1, 1).setValue(h)
        .setBackground('#0d5c60').setFontColor('#ffffff').setFontWeight('bold');
      added = true;
    }
  });
  if (added) { invalidate_(CFG.SH_TRACK); idx = headerIndex_(CFG.SH_TRACK); }
  return idx;
}

function trackFiles_(r) {
  const names = s_(r['附件檔名']) ? s_(r['附件檔名']).split('|') : [];
  const urls  = s_(r['附件連結']) ? s_(r['附件連結']).split('|') : [];
  return names.map((n, i) => ({ n: n, url: urls[i] || '' }));
}

function trackRow_(id) {
  return readTable_(CFG.SH_TRACK).filter(r => String(r['編號']) === String(id))[0];
}

/** 上傳追蹤事項附件（負責人或管理者） */
function api_uploadTrackFile(hrRaw, id, fileName, mimeType, b64) {
  const hr = upper_(hrRaw);
  const me = loadPeople_()[hr];
  if (!me) return { ok: false, msg: '權限不足。' };
  const idx = ensureTrackCols_();
  const t = trackRow_(id);
  if (!t) return { ok: false, msg: '找不到該事項。' };
  if (!isOwner_(t['負責人員'], me.name) && me.role !== '管理者') {
    return { ok: false, msg: '這不是指派給你的事項。' };
  }
  const cur = trackFiles_(t);
  if (cur.length >= 5) return { ok: false, msg: '每筆事項最多 5 個附件。' };

  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, fileName);
  const file = subFolder_(ym_() + '／追蹤事項').createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const sh = sheet_(CFG.SH_TRACK);
  const names = cur.map(f => f.n).concat([fileName]);
  const urls  = cur.map(f => f.url).concat([file.getUrl()]);
  sh.getRange(t._row, idx['附件檔名']).setValue(names.join('|'));
  sh.getRange(t._row, idx['附件連結']).setValue(urls.join('|'));
  invalidate_(CFG.SH_TRACK);
  return { ok: true, files: names.map((n, i) => ({ n: n, url: urls[i] })) };
}

function api_deleteTrackFile(hrRaw, id, index) {
  const hr = upper_(hrRaw);
  const me = loadPeople_()[hr];
  if (!me) return { ok: false, msg: '權限不足。' };
  const idx = ensureTrackCols_();
  const t = trackRow_(id);
  if (!t) return { ok: false, msg: '找不到該事項。' };
  if (!isOwner_(t['負責人員'], me.name) && me.role !== '管理者') {
    return { ok: false, msg: '這不是指派給你的事項。' };
  }
  const cur = trackFiles_(t);
  cur.splice(index, 1);
  const sh = sheet_(CFG.SH_TRACK);
  sh.getRange(t._row, idx['附件檔名']).setValue(cur.map(f => f.n).join('|'));
  sh.getRange(t._row, idx['附件連結']).setValue(cur.map(f => f.url).join('|'));
  invalidate_(CFG.SH_TRACK);
  return { ok: true, files: cur };
}

/** 刪除一筆申請紀錄（僅管理者） */
function api_deleteRequest(hrRaw, id) {
  if (!isAdmin_(hrRaw)) return { ok: false, msg: '權限不足。' };
  const sh = reqSheet_();
  const rows = readTable_(CFG.SH_REQ);
  const r = rows.filter(x => String(x['編號']) === String(id))[0];
  if (!r) return { ok: false, msg: '找不到該申請。' };
  sh.deleteRow(r._row);
  invalidate_(CFG.SH_REQ);
  return { ok: true, msg: '已刪除該筆紀錄。' };
}

/** 個人歷年紀錄：含已結束的業務，依年月由新到舊 */
function api_myHistory(hrRaw, year) {
  const hr = upper_(hrRaw);
  const me = loadPeople_()[hr];
  if (!me) return { ok: false, msg: '權限不足。' };

  const mine = readTable_(CFG.SH_REPORT)
    .filter(r => upper_(r['主責人事號']) === hr && s_(r['狀態']) === '已送出' && s_(r['本月業務內容']));

  const years = [];
  mine.forEach(r => {
    const y = s_(r['年月']).split('-')[0];
    if (y && years.indexOf(y) < 0) years.push(y);
  });
  years.sort().reverse();

  const pick = s_(year) || years[0] || String(new Date().getFullYear() - 1911);
  const rows = mine.filter(r => s_(r['年月']).indexOf(pick + '-') === 0);
  rows.sort((a, b) => s_(b['年月']).localeCompare(s_(a['年月'])));

  // 依年月分組
  const byYm = {}, order = [];
  rows.forEach(r => {
    const ym = s_(r['年月']);
    if (!byYm[ym]) { byYm[ym] = []; order.push(ym); }
    byYm[ym].push({
      unit: s_(r['組別']),
      text: s_(r['本月業務內容']),
      files: fileList_({ fileName: r['附件檔名'], fileUrl: r['附件連結'] }),
      co: s_(r['協同人員']),
      at: s_(r['最後更新'])
    });
  });

  // 該年度做過的業務清單（含已結束）
  const units = [];
  rows.forEach(r => {
    const u = s_(r['組別']);
    if (u && units.indexOf(u) < 0) units.push(u);
  });

  // 該年度指派給他的追蹤事項
  const tracks = readTable_(CFG.SH_TRACK)
    .filter(r => isOwner_(r['負責人員'], me.name))
    .map(r => ({ q: s_(r['事項內容']), from: s_(r['來源會議']),
                 reply: replyOf_(r['回覆執行進度'], me.name),
                 status: s_(r['狀態']) || '待回覆' }))
    .filter(t => t.reply);

  return { ok: true, name: me.name, years: years, year: pick,
           months: order.map(ym => ({ ym: ym, items: byYm[ym] })),
           units: units, tracks: tracks,
           active: me.own.slice() };
}

/** 結束一項業務：不再出現在填報、投影與催繳，歷史紀錄保留 */
function api_endUnit(hrRaw, unit) {
  const hr = upper_(hrRaw);
  const me = loadPeople_()[hr];
  if (!me) return { ok: false, msg: '權限不足。' };
  if (me.own.indexOf(unit) < 0 && me.role !== '管理者') {
    return { ok: false, msg: '這不是你負責的業務。' };
  }
  const idx = ensureStaffCols_();
  const sh = sheet_(CFG.SH_STAFF);
  const rows = readTable_(CFG.SH_STAFF);
  let hit = false;
  for (let i = 0; i < rows.length; i++) {
    if (upper_(rows[i]['人事號']) === hr && s_(rows[i]['組別']) === unit && !ended_(rows[i])) {
      sh.getRange(rows[i]._row, idx['業務狀態']).setValue('已結束');
      hit = true;
      break;
    }
  }
  invalidate_(CFG.SH_STAFF);
  if (!hit) return { ok: false, msg: '找不到這項業務。' };
  return { ok: true, msg: '已結束「' + unit + '」，不會再出現在填報與部務會議報告中。歷史紀錄仍保留。' };
}

/** 管理者移除同仁自行新增的業務（一併刪除名冊那一列） */
function api_removeUnit(hrRaw, id) {
  if (!isAdmin_(hrRaw)) return { ok: false, msg: '權限不足。' };
  const rows = readTable_(CFG.SH_REQ);
  const r = rows.filter(x => String(x['編號']) === String(id))[0];
  if (!r) return { ok: false, msg: '找不到該紀錄。' };

  const targetHr = upper_(r['申請人事號']);
  const unit = s_(r['組別']);

  // 若已有填報內容，先擋下來避免誤刪
  const used = readTable_(CFG.SH_REPORT).filter(x =>
    s_(x['組別']) === unit && upper_(x['主責人事號']) === targetHr && s_(x['本月業務內容']));
  if (used.length) {
    return { ok: false, msg: '「' + unit + '」已經有填報內容，無法直接移除。如確定要移除，請先到試算表刪除相關月報列。' };
  }

  const st = sheet_(CFG.SH_STAFF);
  const staffRows = readTable_(CFG.SH_STAFF);
  for (let i = staffRows.length - 1; i >= 0; i--) {
    if (upper_(staffRows[i]['人事號']) === targetHr && s_(staffRows[i]['組別']) === unit) {
      st.deleteRow(staffRows[i]._row);
      break;
    }
  }
  const sh = reqSheet_();
  sh.deleteRow(r._row);
  invalidate_(CFG.SH_STAFF);
  invalidate_(CFG.SH_REQ);
  return { ok: true, msg: '已移除「' + unit + '」。' };
}

function myRequests_(hr) {
  try {
    return readTable_(CFG.SH_REQ)
      .filter(r => upper_(r['申請人事號']) === upper_(hr))
      .map(r => ({ id: r['編號'], unit: s_(r['組別']), why: s_(r['說明']),
                   status: s_(r['狀態']), at: s_(r['送出時間']) }));
  } catch (e) { return []; }
}

function allRequests_() {
  try {
    return readTable_(CFG.SH_REQ).map(r => ({
      id: r['編號'], at: s_(r['送出時間']), name: s_(r['申請人']),
      unit: s_(r['組別']), isNew: s_(r['是否新組別']) === '是',
      why: s_(r['說明']), status: s_(r['狀態']) }));
  } catch (e) { return []; }
}

/* ===================================================================
 * 八、匯出 Word 會議報告
 * =================================================================*/

function api_export(hrRaw) {
  if (!isAdmin_(hrRaw)) return { ok: false, msg: '權限不足。' };

  const ym = ym_();
  const month = new Date().getMonth() + 1;
  const people = loadPeople_();
  const quarter = loadQuarter_();
  const owners = ownersByUnit_(people);
  const cur = loadReports_(ym);

  const doc = DocumentApp.create(ym + ' ' + CFG.TITLE);
  const body = doc.getBody();
  body.appendParagraph(CFG.TITLE).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(ym + '　部務會議').setHeading(DocumentApp.ParagraphHeading.HEADING3);

  // 壹、上次會議追蹤事項
  body.appendParagraph('壹、上次會議追蹤事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const tr = [['事　項', '回覆執行進度', '負責人員']];
  readTable_(CFG.SH_TRACK).filter(r => s_(r['狀態']) !== '已結案')
    .forEach(r => {
      const ps = parseReply_(r['回覆執行進度']);
      let txt = ps.length
        ? ps.map(x => (x.who ? '【' + x.who + '】\n' : '') + x.text).join('\n')
        : '';
      const fs = trackFiles_(r);
      if (fs.length) txt += (txt ? '\n' : '') + fs.map(f => '附件：' + f.n).join('\n');
      tr.push([s_(r['事項內容']), txt, owners_(r['負責人員']).join('、')]);
    });
  if (tr.length === 1) tr.push(['（無）', '', '']);
  styleTable_(body.appendTable(tr));

  // 貳、業務報告
  body.appendParagraph('貳、各項教學業務報告').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const rt = [['組別／教學業務', '近期進行業務', '負責人員']];
  allUnits_().forEach(u => {
    if (!unitActive_(u, quarter, month)) return;
    const blocks = (owners[u] || [])
      .map(id => cur[u + '|' + id])
      .filter(b => b && b.status === '已送出' && b.text);
    if (!blocks.length) { rt.push([u, '', (owners[u] || []).map(i => people[i].name).join('、')]); return; }
    const text = blocks.map(b => {
      const head = b.item ? ('【' + b.item + '】（' + b.name + '）\n') : ('【' + b.name + '】\n');
      const files = fileList_(b).map(f => '　附件：' + f.n).join('\n');
      return head + b.text + (files ? '\n' + files : '');
    }).join('\n\n');
    rt.push([u, text, blocks.map(b => b.name).join('、')]);
  });
  styleTable_(body.appendTable(rt));

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  subFolder_(ym).addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const docx = 'https://docs.google.com/document/d/' + doc.getId() + '/export?format=docx';
  return { ok: true, url: file.getUrl(), docx: docx };
}

function styleTable_(t) {
  t.setBorderWidth(1);
  const head = t.getRow(0);
  for (let i = 0; i < head.getNumCells(); i++) {
    head.getCell(i).setBackgroundColor('#e3efee').editAsText().setBold(true);
  }
  for (let r = 0; r < t.getNumRows(); r++) {
    for (let c = 0; c < t.getRow(r).getNumCells(); c++) {
      t.getRow(r).getCell(c).editAsText().setFontFamily('標楷體').setFontSize(12);
    }
  }
}

/* ===================================================================
 * 九、定時觸發：會議前自動催繳
 * =================================================================*/

/** 建立每日觸發器（手動執行一次即可） */
function 建立觸發器() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '每日檢查催繳') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('每日檢查催繳').timeBased().atHour(9).everyDays(1).create();
  return '已建立每日 09:00 的催繳檢查觸發器。';
}

function 每日檢查催繳() {
  const today = new Date();
  const day = today.getDate();
  const mDay = meetingDay_();
  const diff = mDay - day;
  if (CFG.REMIND_DAYS.indexOf(diff) < 0) return;
  runReminder_('自動催繳（會議前 ' + diff + ' 天）');
}

/* ===================================================================
 * 十、安裝檢查
 * =================================================================*/

/** 部署前先執行這一支，確認欄位與資料都對得起來 */
function 檢查設定() {
  const need = {};
  need[CFG.SH_STAFF]  = ['人事號', '姓名', '電子郵件', '權限', '組別'];
  need[CFG.SH_REPORT] = ['年月', '組別', '主責人事號', '主責姓名', '業務項目', '本月業務內容',
                         '狀態', '附件檔名', '附件連結', '協同人事號', '協同人員', '填報人事號', '最後更新'];
  need[CFG.SH_UNIT]   = ['單位名稱', '報告主管人事號', '姓名', '電子郵件', '報告月份'];
  need[CFG.SH_TRACK]  = ['編號', '來源會議', '事項內容', '負責人員', '回覆執行進度', '狀態'];

  const msg = [];
  Object.keys(need).forEach(name => {
    let idx;
    try { idx = headerIndex_(name); }
    catch (e) { msg.push('✗ 缺少工作表：' + name); return; }
    const miss = need[name].filter(h => !idx[h]);
    msg.push(miss.length ? ('✗ ' + name + ' 缺少欄位：' + miss.join('、')) : ('✓ ' + name + ' 欄位正確'));
  });

  const people = loadPeople_();
  const owners = ownersByUnit_(people);
  msg.push('－－－');
  msg.push('名冊共 ' + Object.keys(people).length + ' 人');
  const noOwner = allUnits_().filter(u => !(owners[u] || []).length);
  msg.push(noOwner.length ? ('✗ 尚未指定主責的組別：' + noOwner.join('、')) : '✓ 所有組別都有主責');

  const admins = Object.keys(people).filter(k => people[k].role === '管理者');
  msg.push(admins.length ? ('✓ 管理者：' + admins.map(k => people[k].name).join('、')) : '✗ 沒有任何管理者');

  const unknown = [];
  Object.keys(people).forEach(k => {
    people[k].own.concat(people[k].view).forEach(u => {
      if (allUnits_().indexOf(u) < 0 && unknown.indexOf(u) < 0) unknown.push(u);
    });
  });
  if (unknown.length) msg.push('✗ 名冊出現不在組別清單中的名稱：' + unknown.join('、'));

  const out = msg.join('\n');
  Logger.log(out);
  return out;
}


/* ===================================================================
 * 十一、前端頁面（與後端同檔，不需另建 index.html）
 * =================================================================*/

function PAGE_HTML() {
  return String.raw`<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<base target="_top">
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=Noto+Serif+TC:wght@500;700&display=swap" rel="stylesheet">
<style>
:root{--pri:#0d5c60;--pri-d:#083f42;--pri-t:#e3efee;--sand:#b5722a;--sand-t:#faf0e0;
--bg:#f4f2ed;--card:#fff;--line:#e2ded5;--ink:#1b2b2c;--soft:#5f7071;--faint:#8b9a9a;
--ok:#1d8064;--ok-t:#e2f2ec;--warn:#a8442f;--warn-t:#fae9e4;--r:12px}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Noto Sans TC",system-ui,"Microsoft JhengHei",sans-serif;background:var(--bg);color:var(--ink);line-height:1.65;-webkit-font-smoothing:antialiased}
button,input,textarea,select{font:inherit;color:inherit}
.hide{display:none!important}
header{background:var(--pri-d);color:#fff;padding:14px 22px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
header .lg{font-family:"Noto Serif TC",serif;font-weight:700;font-size:19px;letter-spacing:.5px}
header .sub{font-size:12.5px;opacity:.72}
header .who{margin-left:auto;display:flex;align-items:center;gap:10px;font-size:13.5px}
header .who .tag{background:rgba(255,255,255,.16);padding:4px 12px;border-radius:99px}
header button.lo{background:transparent;border:1px solid rgba(255,255,255,.4);color:#fff;padding:5px 13px;border-radius:99px;font-size:13px;cursor:pointer}
.wrap{max-width:1060px;margin:0 auto;padding:26px 22px 70px}
.tabs{display:flex;gap:8px;margin-bottom:22px;flex-wrap:wrap}
.tabs button{background:var(--card);border:1px solid var(--line);color:var(--soft);padding:8px 18px;border-radius:99px;cursor:pointer;font-size:14px}
.tabs button.on{background:var(--pri);border-color:var(--pri);color:#fff}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:22px;margin-bottom:18px}
h2{font-family:"Noto Serif TC",serif;font-size:19px;font-weight:700;margin-bottom:4px}
.lead{font-size:13.5px;color:var(--soft);margin-bottom:18px}
.eyebrow{font-size:12px;letter-spacing:2px;color:var(--sand);font-weight:500;margin-bottom:6px}
.login{max-width:420px;margin:60px auto;text-align:center}
.login .card{padding:34px 30px}
.login input{width:100%;padding:13px 16px;border:1.5px solid var(--line);border-radius:10px;font-size:17px;letter-spacing:2px;text-align:center;margin:14px 0 6px}
.login input:focus{outline:none;border-color:var(--pri)}
.btn{background:var(--pri);color:#fff;border:0;padding:12px 22px;border-radius:10px;cursor:pointer;font-size:15px;width:100%}
.btn:disabled{opacity:.5;cursor:default}
.btn.sm{width:auto;padding:8px 16px;font-size:13.5px}
.btn.gh{background:transparent;color:var(--pri);border:1px solid var(--pri)}
.err{color:var(--warn);font-size:13.5px;min-height:20px;margin-top:8px}
.unitbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px}
.unitbar button{background:var(--card);border:1px solid var(--line);padding:7px 15px;border-radius:99px;font-size:13.5px;cursor:pointer;color:var(--soft)}
.unitbar button.on{background:var(--pri-t);border-color:var(--pri);color:var(--pri-d)}
.prev{background:var(--bg);border-left:3px solid var(--faint);border-radius:0 8px 8px 0;padding:12px 15px}
.prev .tx{font-size:13.5px;color:var(--soft);white-space:pre-wrap}
.prevwrap{margin-top:12px}
.prevwrap summary{font-size:13px;color:var(--soft);cursor:pointer;padding:6px 0;list-style:none;display:inline-flex;align-items:center;gap:6px}
.prevwrap summary::-webkit-details-marker{display:none}
.prevwrap summary::before{content:"▸";font-size:11px}
.prevwrap[open] summary::before{content:"▾"}
textarea{width:100%;border:1.5px solid var(--line);border-radius:10px;padding:13px 15px;font-size:14.5px;line-height:1.85;resize:vertical;min-height:130px}
#ta{font-family:"Menlo","Consolas","Noto Sans Mono CJK TC","Noto Sans TC",monospace;font-variant-numeric:tabular-nums;tab-size:4}
textarea:focus{outline:none;border-color:var(--pri)}
textarea.short{min-height:76px}
.tin{width:100%;max-width:420px;border:1.5px solid var(--line);border-radius:10px;padding:10px 14px;font-size:14.5px}
.tin:focus{outline:none;border-color:var(--pri)}
.rowend{display:flex;align-items:center;gap:12px;margin-top:14px;flex-wrap:wrap}
.saved{font-size:13px;color:var(--ok)}
.saved.bad{color:var(--warn)}
.colabel{font-size:13px;color:var(--soft);margin:18px 0 8px}
.talabel{display:flex;align-items:center;gap:12px;font-size:13px;color:var(--soft);margin:18px 0 8px}
.mini{background:transparent;border:1px solid var(--line);color:var(--soft);padding:3px 12px;border-radius:99px;font-size:12.5px;cursor:pointer}
.mini:hover:not(:disabled){border-color:var(--pri);color:var(--pri)}
.mini:disabled{opacity:.4;cursor:default}
.mini.danger{border-color:var(--warn);color:var(--warn)}
.chips{display:flex;gap:8px;flex-wrap:wrap}
.chip{background:var(--card);border:1px solid var(--line);color:var(--soft);padding:6px 14px;border-radius:99px;font-size:13.5px;cursor:pointer}
.chip.on{background:var(--sand-t);border-color:var(--sand);color:var(--sand);font-weight:500}
.chip.on::before{content:"✓ "}
.files{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}
.fchip{display:inline-flex;align-items:center;gap:7px;background:var(--bg);border:1px solid var(--line);border-radius:99px;padding:5px 12px;font-size:13px;color:var(--soft)}
.fchip a{color:var(--pri);text-decoration:none}
.fchip .fx{background:none;border:0;color:var(--faint);cursor:pointer;font-size:15px;line-height:1;padding:0 2px}
.state{display:flex;align-items:center;gap:12px;flex-wrap:wrap;border-radius:10px;padding:9px 14px;margin-top:18px;font-size:13px}
.state.ok{background:var(--ok-t);color:var(--ok);border:1px solid #cfe6dc}
.state.dr{background:var(--sand-t);color:var(--sand);border:1px solid #ecdcc0}
.off{background:var(--pri-t);border:1px solid #c9e0dd;border-radius:10px;padding:11px 15px;margin:6px 0 14px;font-size:13.5px;color:var(--pri)}
.none{display:flex;align-items:center;gap:10px;flex-wrap:wrap;background:var(--pri-t);border:1px solid #c9e0dd;border-radius:10px;padding:10px 14px;margin-top:18px;font-size:13px;color:var(--pri)}
.cfm{display:flex;align-items:center;gap:10px;flex-wrap:wrap;background:var(--warn-t);border:1px solid #edcfc6;border-radius:10px;padding:10px 14px;margin-bottom:10px;font-size:13px;color:var(--warn)}
.track{border:1px solid var(--line);border-radius:10px;padding:15px 17px;margin-bottom:12px}
.track .q{font-size:14.5px;margin-bottom:10px}
.track .meta{font-size:12px;color:var(--faint);margin-bottom:8px;display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.track textarea{min-height:78px}
.pill{font-size:11.5px;padding:2px 10px;border-radius:99px;font-weight:500}
.pill.o{background:var(--sand-t);color:var(--sand)}
.pill.c{background:var(--ok-t);color:var(--ok)}
.pill.n{background:var(--warn-t);color:var(--warn)}
.pill.v{background:var(--pri-t);color:var(--pri)}
.pill.g{background:#eeeae2;color:#7d8583}
label.ck{display:flex;align-items:center;gap:7px;font-size:13.5px;color:var(--soft);cursor:pointer;margin-top:9px}
.ro{border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin-bottom:12px}
.ro .rh{display:flex;align-items:center;gap:10px;margin-bottom:9px;flex-wrap:wrap}
.ro .rn{font-weight:500;font-size:15px}
.ro .rb{font-size:14px;color:var(--soft)}
.ro .rm{font-size:12px;color:var(--faint);margin-top:9px}
.cow{font-size:12.5px;color:var(--sand);background:var(--sand-t);padding:2px 10px;border-radius:99px}
.hsum{background:var(--bg);border-radius:10px;padding:14px 16px;margin-bottom:20px}
.hsum .hl{font-size:12.5px;color:var(--faint);margin-bottom:8px}
.hsum .hu{display:inline-block;background:var(--card);border:1px solid var(--line);
  border-radius:99px;padding:5px 14px;font-size:13.5px;margin:0 6px 6px 0}
.hsum .hu.off{color:var(--faint);border-style:dashed}
.hm{margin-bottom:26px}
.hm .hmy{font-family:"Noto Serif TC",serif;font-size:17px;font-weight:700;
  padding-bottom:8px;border-bottom:1px solid var(--line);margin-bottom:12px}
.hi2{margin-bottom:16px}
.hi2 .hiu{font-weight:500;font-size:14.5px;color:var(--pri-d);margin-bottom:5px}
.hi2 .hib{font-size:14px;color:var(--soft)}
.hi2 .him{font-size:12px;color:var(--faint);margin-top:5px}
.ohead{border-bottom:1px solid var(--line);padding-bottom:12px;margin-bottom:16px}
.ohead .ot{font-family:"Noto Serif TC",serif;font-size:22px;font-weight:700}
.ohead .os{font-size:13px;color:var(--soft);margin-top:2px}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:16px 18px}
.stat .n{font-family:"Noto Serif TC",serif;font-size:30px;font-weight:700;line-height:1.1;color:var(--pri)}
.stat .l{font-size:12.5px;color:var(--soft);margin-top:3px}
.stat.w .n{color:var(--warn)}
table{width:100%;border-collapse:collapse;font-size:14px}
th{text-align:left;font-weight:500;font-size:12.5px;color:var(--faint);padding:0 10px 9px;border-bottom:1px solid var(--line)}
td{padding:11px 10px;border-bottom:1px solid var(--line);vertical-align:top}
td.u{font-weight:500;white-space:nowrap}
td.p{color:var(--soft);font-size:13.5px}
.dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:7px;vertical-align:middle}
.dot.y{background:var(--ok)}.dot.n{background:#d8d3c8}.dot.g{background:#b9b3a7}.dot.d{background:var(--sand)}
.cnt{font-size:11.5px;color:var(--sand);background:var(--sand-t);padding:1px 8px;border-radius:99px}
.noted{font-size:11.5px;color:var(--faint);margin-top:4px}
.acts{white-space:nowrap}
.acts .mini{margin:0 4px 4px 0}
.reqrow{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:6px}
.reqform{border:1px solid var(--line);border-radius:12px;padding:18px 20px;margin:8px 0 4px;background:var(--bg)}
.reqform .rft{font-weight:500;font-size:15px}
.reqform .rfd{font-size:12.5px;color:var(--faint);margin-top:3px;line-height:1.7}
.rq{display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-size:13px;color:var(--soft);margin:4px 0}
.rq .rqu{font-weight:500;color:var(--ink)}
.proj{background:var(--pri-d);color:#fff;min-height:100vh;padding:36px 44px;display:flex;flex-direction:column}
.proj .sheet{display:flex;flex-direction:column;flex:1;min-height:0}
.proj .top{display:flex;align-items:baseline;gap:16px;border-bottom:1px solid rgba(255,255,255,.18);padding-bottom:16px;margin-bottom:32px}
.proj .top .t{font-family:"Noto Serif TC",serif;font-size:26px;font-weight:700}
.proj .top .m{font-size:15px;opacity:.6;margin-left:auto}
.proj .unit{font-family:"Noto Serif TC",serif;font-size:44px;font-weight:700;line-height:1.25;margin-bottom:6px}
.proj .owner{font-size:17px;color:#8fd0c4;margin-bottom:26px}
.proj .body{font-size:23px;line-height:1.85;flex:1;overflow-y:auto;padding-right:8px}
.proj .body.empty{opacity:.4;font-size:19px}
.proj .pblk{margin-bottom:24px;padding-bottom:24px;border-bottom:1px solid rgba(255,255,255,.16)}
.proj .pblk:last-child{border-bottom:0;padding-bottom:0}
.proj .phrow{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:12px}
.proj .ph{display:inline-block;font-size:19px;font-weight:700;color:#083f42;background:#8fd0c4;padding:3px 16px;border-radius:99px}
.proj .pco{font-size:16px;color:#8fd0c4;opacity:.85}
.proj .pt{max-width:1600px;padding-left:2px}
.ln{white-space:pre-wrap;word-break:break-word}
.ln.plain{font-family:"Menlo","Consolas","Noto Sans Mono CJK TC",monospace;font-variant-numeric:tabular-nums}
.ln.hi{color:var(--sand);font-weight:500}
.proj .ln.hi{color:#f5c96b;font-weight:700}
.proj .pt .ln.hi{background:rgba(245,201,107,.10);box-shadow:inset 4px 0 0 #f5c96b;
  border-radius:0 6px 6px 0;padding-top:3px;padding-bottom:3px;padding-right:12px}
.dtab.hi th,.dtab.hi td{border-color:rgba(245,201,107,.5)}
.proj .dtab.hi th{background:rgba(245,201,107,.18);color:#f5c96b}
.proj .dtab.hi td{color:#f5c96b}
.dtab{border-collapse:collapse;margin:10px 0;font-variant-numeric:tabular-nums}
.dtab th,.dtab td{border:1px solid rgba(127,127,127,.35);padding:5px 16px;text-align:left;white-space:nowrap}
.dtab th{font-weight:700;background:rgba(127,127,127,.14)}
.proj .dtab{font-size:.92em}
.proj .dtab th,.proj .dtab td{border-color:rgba(143,208,196,.45);padding:7px 22px}
.proj .dtab th{background:rgba(143,208,196,.18);color:#d6ebe6}
.ro .dtab{font-size:13px}
.ln.blank{height:.7em}
.proj .pt .ln{margin-bottom:2px}
.proj .pf{font-size:15px;color:#8fd0c4;margin-top:10px}
.proj .tk{padding-bottom:20px;margin-bottom:20px;border-bottom:1px solid rgba(255,255,255,.16)}
.proj .tk:last-child{border-bottom:0;margin-bottom:0}
.proj .tkh{display:flex;align-items:center;gap:12px;margin-bottom:8px;flex-wrap:wrap}
.proj .tkn{display:inline-block;font-size:16px;font-weight:700;color:#083f42;background:#8fd0c4;padding:2px 14px;border-radius:99px}
.proj .tkn.done{background:#6fae9a}
.proj .tkq{font-size:20px;line-height:1.7;max-width:1600px}
.proj .tkw{display:inline-block;font-size:15px;font-weight:700;color:#083f42;
  background:#8fd0c4;padding:2px 14px;border-radius:99px;margin-bottom:8px}
.tkfiles:not(:empty){margin-top:10px}
.tfs{margin-top:6px;font-size:12.5px}
.tfs a{color:var(--pri)}
.othr{background:var(--bg);border-radius:8px;padding:10px 14px;margin-bottom:10px;font-size:13.5px;color:var(--soft)}
.othr .ow{font-size:12px;color:var(--faint);margin-bottom:4px}
.proj .tkr{font-size:19px;line-height:1.7;margin-top:14px;padding-left:20px;
  border-left:4px solid rgba(143,208,196,.55);color:#e3f2ee;max-width:1600px}
.proj .tkr.empty{color:rgba(255,255,255,.4)}
.proj .pf a{color:#8fd0c4}
.proj.pt9{align-items:center;background:#052e30}
.proj.pt9 .top,.proj.pt9 .head,.proj.pt9 .body,.proj.pt9 .nav{width:100%}
.proj.pt9 .sheet{width:min(78vh,100%);max-width:820px;background:var(--pri-d);
  border-radius:14px;padding:38px 40px 30px;display:flex;flex-direction:column;
  flex:1;box-shadow:0 10px 40px rgba(0,0,0,.35);overflow:hidden}
.proj.pt9 .unit{font-size:34px}
.proj.pt9 .owner{font-size:15px;margin-bottom:20px}
.proj.pt9 .body{font-size:19px}
.proj.pt9 .top .t{font-size:21px}
.proj.pt9 .dtab th,.proj.pt9 .dtab td{padding:5px 14px}
.proj .nav{display:flex;align-items:center;gap:16px;padding-top:22px;border-top:1px solid rgba(255,255,255,.18);font-size:14px}
.proj .nav button{background:rgba(255,255,255,.12);border:0;color:#fff;padding:9px 18px;border-radius:99px;cursor:pointer;font-size:14px}
.proj .nav .pg{margin-left:auto;opacity:.6}
.proj .kb{opacity:.45;font-size:13px}
#busy{position:fixed;top:0;left:0;right:0;height:3px;background:var(--sand);transform:scaleX(0);transform-origin:left;transition:transform .3s;z-index:99}
#busy.on{transform:scaleX(.8)}
@media(max-width:820px){.stats{grid-template-columns:repeat(2,1fr)}.proj{padding:28px 24px}.proj .unit{font-size:32px}.proj .body{font-size:19px}}
</style>
</head>
<body>
<div id="busy"></div>

<div id="v-login" class="login">
  <div class="card">
    <div class="eyebrow">奇美醫院 教學部</div>
    <h2>各項教學業務報告</h2>
    <p class="lead">請輸入人事號進入本月填報</p>
    <input id="hrid" placeholder="人事號" maxlength="10" autocomplete="off">
    <div class="err" id="loginerr"></div>
    <button class="btn" id="loginbtn" onclick="doLogin()">進入填報</button>
  </div>
</div>

<div id="v-app" class="hide">
<header>
  <div>
    <div class="lg">教學部各項教學業務報告</div>
    <div class="sub" id="ymtag"></div>
  </div>
  <div class="who">
    <button class="lo" id="refbtn" onclick="refresh()">↻ 重新整理</button>
    <span class="tag" id="whotag"></span>
    <button class="lo" onclick="logout()">登出</button>
  </div>
</header>

<div class="wrap">
  <div class="tabs" id="tabs"></div>

  <section id="p-report">
    <div class="card hide" id="trackcard">
      <div class="eyebrow">01</div>
      <h2>上次會議追蹤事項</h2>
      <p class="lead">指派給你的事項會自動出現在這裡，請填寫執行進度並儲存。是否結案由部務會議決定，未結案下個月會再出現。</p>
      <div id="tracklist"></div>
    </div>

    <div class="card" id="editcard">
      <div class="eyebrow">02</div>
      <h2>本月業務報告</h2>
      <p class="lead">直接填寫本月內容；若要沿用上月，按「帶入上月內容」再修改。</p>
      <div class="unitbar" id="unitbar"></div>
      <div class="off hide" id="offbar"><span id="offtxt"></span></div>
      <div class="reqrow">
        <button class="mini" id="reqbtn" onclick="openReq()">＋ 新增我的業務</button>
        <div id="reqlist"></div>
      </div>
      <div class="reqform hide" id="reqform">
        <div class="rft">新增我的業務</div>
        <div class="rfd">新增後立即生效，可直接開始填報，並會通知承辦人。若只是既有業務的一部分，建議直接寫在該項內容裡，不必另外新增。</div>
        <div class="colabel">要新增的業務</div>
        <select id="requnit" class="tin" onchange="reqUnitChange()"></select>
        <div id="reqnewwrap" class="hide">
          <div class="colabel">業務名稱</div>
          <input id="reqnew" class="tin" placeholder="請填寫完整名稱" autocomplete="off">
        </div>
        <div class="colabel">業務內容與來源（選填）</div>
        <textarea id="reqwhy" class="short" placeholder="例如：9/5 部務會議主任交辦，負責醫學教育委員會會議紀錄與追蹤"></textarea>
        <div class="rowend">
          <button class="btn sm" onclick="sendReq()">新增</button>
          <button class="btn sm gh" onclick="closeReq()">取消</button>
          <span class="saved bad" id="reqmsg"></span>
        </div>
      </div>

      <div id="editbody">
        <div id="itemwrap" class="hide">
          <input id="item" class="tin" autocomplete="off">
        </div>
        <div class="talabel"><span>本月業務內容</span><button class="mini" id="copybtn" onclick="copyPrev()">帶入上月內容</button></div>
        <div class="cfm hide" id="confirmbar">
          <span>目前欄位已有內容，帶入上月內容會整段覆蓋掉。</span>
          <button class="mini danger" onclick="doCopy()">確定覆蓋</button>
          <button class="mini" onclick="cancelCopy()">取消</button>
        </div>
        <textarea id="ta" placeholder="條列本月進行中的業務、專案與活動"></textarea>
        <div class="colabel">附件（選填，簡報、海報等，最多 5 個；報告時可直接點開）</div>
        <div><button class="mini" id="upbtn" onclick="document.getElementById('fileinput').click()">＋ 上傳檔案</button>
          <input type="file" id="fileinput" multiple class="hide" onchange="addFiles(this)"></div>
        <div id="filelist"></div>
        <details class="prevwrap hide" id="prevwrap"><summary>查看上月原始內容</summary><div class="prev" id="prevbox"></div></details>
        <div class="colabel">本月協同人員（點選加入或移除，可複選。已帶入上月選擇；被選到的人可登入檢視本組內容，但不能修改）</div>
        <div class="chips" id="cochips"></div>
        <div class="none hide" id="nonebar">
          <span>本月已標記為「無業務」，不會列入催繳，投影時也不會呈現這一塊。</span>
          <button class="mini" onclick="unmarkNone()">取消標記</button>
        </div>
        <div class="state hide" id="statebar"><span id="statetxt"></span><button class="mini hide" id="unsendbtn" onclick="unsend()">收回修改</button></div>
        <div class="rowend">
          <button class="btn sm" onclick="submitReport()">送出報告</button>
          <button class="btn sm gh" onclick="saveDraft()">暫存</button>
          <button class="btn sm gh" onclick="markNone()">本月無業務</button>
          <button class="mini" id="endbtn" onclick="endUnit()">結束此業務</button>
          <span class="saved" id="savedmsg"></span>
        </div>
      </div>
    </div>

    <div class="card hide" id="viewcard">
      <div class="eyebrow">03</div>
      <h2>我協同的業務</h2>
      <p class="lead">以下是本月將你列為協同人員的組別，內容由該組主責填寫，此處僅供檢視。</p>
      <div id="viewlist"></div>
    </div>
  </section>

  <section id="p-hist" class="hide">
    <div class="card">
      <h2>我的歷年紀錄</h2>
      <p class="lead" id="histnote"></p>
      <div class="unitbar" id="histyears"></div>
      <div id="histunits"></div>
      <div id="histlist"></div>
    </div>
  </section>

  <section id="p-over" class="hide">
    <div class="card">
      <h2>全體業務報告</h2>
      <p class="lead" id="overnote"></p>
      <div class="unitbar" id="overbar"></div>
      <div class="ohead"><div class="ot" id="overtitle"></div><div class="os" id="oversub"></div></div>
      <div id="overlist"></div>
    </div>
  </section>

  <section id="p-admin" class="hide">
    <div class="stats" id="stats"></div>
    <div class="card">
      <h2>各組填報進度</h2>
      <p class="lead">一組有多位主責時逐位列出。系統會在會議前自動催繳，也可以在這裡個別或整批寄送。</p>
      <table><thead><tr><th style="width:24%">組別／教學業務</th><th style="width:10%">主責</th><th style="width:18%">業務項目</th><th style="width:12%">狀態</th><th style="width:14%">最後更新</th><th>催繳</th></tr></thead><tbody id="admtb"></tbody></table>
      <div class="rowend">
        <button class="btn sm" onclick="remindAll()">一次通知所有未完成者</button>
        <button class="btn sm gh" onclick="notifyMgr()">通知主管本月已彙整完成</button>
        <button class="btn sm gh" onclick="expo()">匯出 Word 會議報告</button>
        <span class="saved" id="admmsg"></span>
      </div>
    </div>
    <div class="card">
      <h2>新增業務紀錄</h2>
      <p class="lead" id="reqcount"></p>
      <table><thead><tr><th style="width:13%">時間</th><th style="width:9%">同仁</th><th style="width:24%">業務</th><th style="width:30%">說明</th><th style="width:10%">狀態</th><th>處理</th></tr></thead><tbody id="reqtb"></tbody></table>
      <div class="rowend"><span class="saved" id="reqadmmsg"></span></div>
    </div>
    <div class="card">
      <h2>追蹤事項</h2>
      <p class="lead">會議中主任指示請直接在 Google 試算表的「追蹤事項」分頁建檔並指定負責人。同仁只負責回覆，結案與否在這裡決定。</p>
      <table><thead><tr><th style="width:40%">事項</th><th style="width:10%">負責人員</th><th style="width:11%">狀態</th><th style="width:26%">回覆</th><th>處理</th></tr></thead><tbody id="trktb"></tbody></table>
      <div class="rowend"><span class="saved" id="trkmsg"></span></div>
    </div>
  </section>
</div>
</div>

<input type="file" id="trackfileinput" class="hide" onchange="uploadTrackFile(this)">

<div id="v-proj" class="hide">
  <div class="proj" id="projwrap">
    <div class="sheet" id="projsheet">
    <div class="top"><div class="t">教學部部務會議</div><div class="m" id="projym"></div></div>
    <div><div class="unit" id="pu"></div><div class="owner" id="po"></div></div>
    <div class="body" id="pb"></div>
    </div>
    <div class="nav">
      <button onclick="pmove(-1)">← 上一組</button>
      <button onclick="pmove(1)">下一組 →</button>
      <button id="orbtn" onclick="toggleOrient()">切換直式</button>
      <span class="kb">左右鍵切換頁面，V 切換直橫式，Esc 離開</span>
      <span class="pg" id="ppg"></span>
      <button onclick="exitProj()">離開投影</button>
    </div>
  </div>
</div>

<script>
var BOOT=null, ME=null, curUnit=null, ADMIN=null, OVER=null, oIdx=0, pIdx=0;
function $(i){return document.getElementById(i);}
function esc(t){return String(t==null?"":t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
function _w(s){var n=0;for(var i=0;i<s.length;i++){n+=(s.charCodeAt(i)>255?1:0.5);}return n;}
function _cells(ln){
  var toks=ln.trim().split(/[\s　]+/).filter(Boolean);
  var out=[];
  toks.forEach(function(t){
    if(/^\d{1,2}\/\d{1,2}$/.test(t)||/^\d{1,4}\/\d{1,2}\/\d{1,2}$/.test(t)){out.push(t);return;}
    t=t.replace(/^\/+|\/+$/g,"");
    if(t.indexOf("/")>=0&&!/\d\/\d/.test(t)){t.split("/").forEach(function(p){if(p)out.push(p);});}
    else if(t){out.push(t);}
  });
  return out;
}
function _isTable(rows){
  if(rows.length<2)return false;
  var n=rows[0].length;
  if(n<2)return false;
  for(var i=0;i<rows.length;i++){
    if(rows[i].length!==n)return false;
    for(var j=0;j<n;j++){if(_w(rows[i][j])>14)return false;}
  }
  return true;
}
function _emitTable(rows,pad,out,hi){
  var html='<table class="dtab'+(hi?" hi":"")+'" style="margin-left:'+pad.toFixed(2)+'em"><tbody>';
  rows.forEach(function(r,i){
    html+='<tr>'+r.map(function(c){
      return (i===0?'<th>':'<td>')+esc(c)+(i===0?'</th>':'</td>');
    }).join("")+'</tr>';
  });
  out.push(html+'</tbody></table>');
}
function _plain(ln,pad,out,hi){
  var p=pad+(hi?0.7:0);
  out.push('<div class="ln plain'+(hi?" hi":"")+'" style="padding-left:'+p.toFixed(2)+'em">'+esc(ln)+'</div>');
}
function _flush(buf,pad,out,hi){
  if(!buf.length)return;
  var cells=buf.map(_cells);
  var i=0;
  while(i<buf.length){
    var n=cells[i].length, j=i;
    if(n>=2){
      while(j+1<buf.length&&cells[j+1].length===n)j++;
      var run=cells.slice(i,j+1);
      if(_isTable(run)){_emitTable(run,pad,out,hi);i=j+1;continue;}
    }
    _plain(buf[i],pad,out,hi);
    i++;
  }
  buf.length=0;
}
function fmtBody(text){
  var lines=String(text==null?"":text).split(/\r?\n/);
  var reNum=/^(\d{1,2}\.\d{1,2}[.、)]|\d{1,2}[.、):])\s*/;
  var reParen=/^([(（][^)）]{1,6}[)）])\s*/;
  var reBul=/^([・‧•·※◆▪])\s*/;
  var STEP=1.7, base=0, out=[], buf=[], hiLv=-1, curHi=false;
  lines.forEach(function(raw){
    var ln=raw.replace(/^[\s　]+/,"").replace(/[\s　]+$/,"");
    if(!ln){_flush(buf,base,out,curHi);out.push('<div class="ln blank"></div>');return;}

    var star=false;
    if(/^[*＊]/.test(ln)){star=true;ln=ln.replace(/^[*＊][\s　]*/,"");}

    var m,lvl=-1,mk="";
    if((m=ln.match(reNum))){lvl=0;mk=m[1]+" ";}
    else if((m=ln.match(reParen))){lvl=1;mk=m[1]+" ";}
    else if((m=ln.match(reBul))){lvl=2;mk=m[1]+" ";}

    var body=(lvl<0)?ln:ln.slice(m[0].length);
    if(/^[*＊]/.test(body)){star=true;body=body.replace(/^[*＊][\s　]*/,"");}

    if(lvl<0){
      if(star){_flush(buf,base,out,curHi);curHi=true;hiLv=99;
        _plain(body,base,out,true);return;}
      buf.push(body);return;
    }

    _flush(buf,base,out,curHi);
    if(star){curHi=true;hiLv=lvl;}
    else if(curHi&&lvl<=hiLv){curHi=false;hiLv=-1;}

    var mw=_w(mk), left=lvl*STEP;
    base=left+mw;
    var pl=base+(curHi?0.7:0);
    out.push('<div class="ln'+(curHi?" hi":"")+'" style="padding-left:'+pl.toFixed(2)+
             'em;text-indent:-'+mw.toFixed(2)+'em">'+esc(mk+body)+'</div>');
  });
  _flush(buf,base,out,curHi);
  return out.join("");
}
function tkLabel(status,reply){
  if(status==="已結案")return "已結案";
  return (reply&&String(reply).trim())?"執行進度":"待回覆";
}
function busy(on){$("busy").classList.toggle("on",!!on);}

function call(fn){
  var args=Array.prototype.slice.call(arguments,1);
  busy(true);
  return new Promise(function(res,rej){
    google.script.run
      .withSuccessHandler(function(r){busy(false);res(r);})
      .withFailureHandler(function(e){busy(false);rej(e);})
      [fn].apply(null,args);
  });
}
function fail(e){alert("系統忙線或發生錯誤，請稍後再試。\n\n"+(e&&e.message?e.message:e));}

/* ---------- 登入 ---------- */
function doLogin(){
  var hr=$("hrid").value.trim().toUpperCase();
  if(!hr){$("loginerr").textContent="請輸入人事號。";return;}
  $("loginbtn").disabled=true;
  call("api_bootstrap",hr).then(function(r){
    $("loginbtn").disabled=false;
    if(!r||!r.ok){$("loginerr").textContent=(r&&r.msg)||"登入失敗。";return;}
    BOOT=r; ME=r.me; curUnit=null;
    $("loginerr").textContent="";
    $("v-login").classList.add("hide");
    $("v-app").classList.remove("hide");
    $("ymtag").textContent=r.ym+"　部務會議";
    $("projym").textContent=r.ym;
    $("whotag").textContent=ME.name+"（"+ME.hr+"）"+(ME.role==="管理者"?" · 管理者":(ME.role==="主管"?" · 主管":""));
    buildTabs();
  })["catch"](function(e){$("loginbtn").disabled=false;fail(e);});
}
function logout(){
  BOOT=null;ME=null;curUnit=null;ADMIN=null;OVER=null;oIdx=0;pIdx=0;
  $("v-proj").classList.add("hide");
  $("v-app").classList.add("hide");
  $("p-report").classList.remove("hide");
  $("p-over").classList.add("hide");
  $("p-admin").classList.add("hide");
  $("tabs").innerHTML="";
  $("hrid").value="";
  $("loginerr").textContent="";
  $("loginbtn").disabled=false;
  $("v-login").classList.remove("hide");
  window.scrollTo(0,0);
  $("hrid").focus();
}
function refresh(){
  var b=$("refbtn");
  b.disabled=true;b.textContent="更新中…";
  var done=function(){b.disabled=false;b.textContent="↻ 重新整理";};
  var on=document.querySelector("#tabs button.on");
  var label=on?on.textContent:"";
  if(label.indexOf("歷年")>=0){loadHist(HIST?HIST.year:"").then(done,done);}
  else if(label.indexOf("彙整")>=0){loadAdmin().then(done,done);}
  else if(label.indexOf("全體")>=0){loadOver().then(done,done);}
  else{reload().then(done,done);}
}
$("hrid").addEventListener("keydown",function(e){if(e.key==="Enter")doLogin();});

function reload(){
  return call("api_bootstrap",ME.hr).then(function(r){
    if(r&&r.ok){BOOT=r;renderReport();}
  })["catch"](fail);
}

/* ---------- 分頁 ---------- */
function buildTabs(){
  var h="";
  var boss=(ME.role==="主管"||ME.role==="管理者");
  if(BOOT.mine.length){h+='<button onclick="tab(\'report\',this)">我的填報</button>';}
  h+='<button onclick="tab(\'hist\',this)">我的歷年紀錄</button>';
  if(boss){h+='<button onclick="tab(\'over\',this)">全體業務報告</button>';}
  if(ME.role==="管理者"){h+='<button onclick="tab(\'admin\',this)">彙整與催繳</button>';}
  if(boss){h+='<button onclick="enterProj()">會議投影模式</button>';}
  else{h+='<button onclick="enterProj()">會議投影模式</button>';}
  $("tabs").innerHTML=h;
  var first=document.querySelector("#tabs button");
  if(first){first.classList.add("on");}
  if(BOOT.mine.length){renderReport();}
  else if(first){tab("over",first);}
}
function tab(k,el){
  var bs=document.querySelectorAll("#tabs button");
  for(var i=0;i<bs.length;i++){bs[i].classList.remove("on");}
  if(el){el.classList.add("on");}
  $("p-report").classList.toggle("hide",k!=="report");
  $("p-hist").classList.toggle("hide",k!=="hist");
  $("p-over").classList.toggle("hide",k!=="over");
  $("p-admin").classList.toggle("hide",k!=="admin");
  if(k==="hist"){loadHist();}
  else if(k==="admin"){loadAdmin();}
  else if(k==="over"){loadOver();}
  else{renderReport();}
}

/* ---------- 我的填報 ---------- */
function mineOf(u){
  for(var i=0;i<BOOT.mine.length;i++){if(BOOT.mine[i].unit===u)return BOOT.mine[i];}
  return null;
}
function renderReport(){
  var t=BOOT.tracks||[];
  $("trackcard").classList.toggle("hide",t.length===0);
  $("tracklist").innerHTML=t.map(function(x){
    var done=(x.status==="已結案");
    var lab=tkLabel(x.status,x.reply);
    var mates=(x.owners||[]).filter(function(n){return n!==ME.name;});
    var others=(x.others||[]).filter(function(o){return o.text;});
    return '<div class="track">'+
      '<div class="meta"><span class="pill '+(done?"c":(lab==="執行進度"?"v":"o"))+'" id="tp'+x.id+'">'+esc(lab)+'</span>'+
      '<span>'+esc(x.from)+'</span>'+
      (mates.length?'<span>共同負責：'+esc(mates.join("、"))+'</span>':'')+'</div>'+
      '<div class="q">'+esc(x.q)+'</div>'+
      (others.length?others.map(function(o){
        return '<div class="othr"><div class="ow">'+esc(o.who)+' 的回覆</div>'+fmtBody(o.text)+'</div>';
      }).join(""):"")+
      '<textarea id="tr'+x.id+'" placeholder="'+(mates.length?"請填寫你負責部分的執行進度":"請填寫執行進度或處理情形")+'" oninput="trackDirty('+x.id+')">'+esc(x.reply)+'</textarea>'+
      '<div class="tkfiles" id="tf'+x.id+'">'+fileChips(x.files,true,x.id)+'</div>'+
      '<div class="rowend"><button class="btn sm gh" onclick="saveTrack('+x.id+')">儲存回覆</button>'+
      '<button class="mini" id="tub'+x.id+'" onclick="pickTrackFile('+x.id+')">＋ 上傳檔案</button>'+
      '<span class="saved" id="tm'+x.id+'"></span></div></div>';
  }).join("");

  $("editcard").classList.toggle("hide",BOOT.mine.length===0);
  if(!BOOT.mine.length){renderCo();return;}
  if(!curUnit||!mineOf(curUnit)){curUnit=BOOT.mine[0].unit;}
  var m=mineOf(curUnit);

  $("unitbar").innerHTML=BOOT.mine.map(function(x){
    return '<button class="'+(x.unit===curUnit?"on":"")+'" onclick="pickUnit(\''+x.unit+'\')">'+esc(x.unit)+'</button>';
  }).join("");
  $("unitbar").classList.toggle("hide",BOOT.mine.length<2);

  $("offbar").classList.toggle("hide",m.active);
  $("editbody").classList.toggle("hide",!m.active);
  if(!m.active){
    $("offtxt").textContent="貴單位為每季報告（"+(m.months||[]).join("、")+" 月），本月免填報。";
    return;
  }

  var b=m.block;
  $("item").value=b?b.item:"";
  $("ta").value=(b&&b.status!=="本月無業務")?b.text:"";
  $("prevwrap").classList.toggle("hide",!m.prevText);
  $("prevwrap").open=false;
  $("prevbox").innerHTML='<div class="tx">'+esc(m.prevText)+'</div>';
  $("copybtn").disabled=!m.prevText;
  $("confirmbar").classList.add("hide");

  var co=b?b.co:(m.prevCo||[]);
  $("cochips").innerHTML=(BOOT.people||[]).filter(function(p){
    return p.hr!==ME.hr&&p.role!=="管理者";
  }).map(function(p){
    return '<button class="chip '+(co.indexOf(p.hr)>=0?"on":"")+'" onclick="toggleCo(\''+p.hr+'\')">'+esc(p.name)+'</button>';
  }).join("");

  var stt=b?b.status:"";
  $("nonebar").classList.toggle("hide",stt!=="本月無業務");
  $("statebar").className="state "+(stt==="已送出"?"ok":(stt==="草稿"?"dr":"hide"));
  $("statetxt").textContent=stt==="已送出"?("已於 "+b.at+" 送出")
    :(stt==="草稿"?("暫存中，尚未送出（"+b.at+"）"):"");
  $("unsendbtn").classList.toggle("hide",stt!=="已送出");
  $("filelist").innerHTML=fileChips(b?filesOf(b):[],true);
  $("savedmsg").textContent="";
  $("savedmsg").classList.remove("bad");
  renderCo();
  renderMyReqs();
}
function filesOf(b){
  var n=b.fileName?String(b.fileName).split("|"):[];
  var u=b.fileUrl?String(b.fileUrl).split("|"):[];
  return n.map(function(x,i){return {n:x,url:u[i]||""};});
}
function fileChips(list,removable,tid){
  if(!list||!list.length)return "";
  return '<div class="files">'+list.map(function(f,i){
    var rm=(tid!==undefined)?'delTrackFile('+tid+','+i+')':'delFile('+i+')';
    return '<span class="fchip">📎 <a href="'+esc(f.url)+'" target="_blank">'+esc(f.n)+'</a>'+
      (removable?'<button class="fx" onclick="'+rm+'">×</button>':'')+'</span>';
  }).join("")+'</div>';
}
function renderCo(){
  var a=BOOT.asCo||[];
  $("viewcard").classList.toggle("hide",a.length===0);
  $("viewlist").innerHTML=a.map(function(b){
    var none=b.status==="本月無業務";
    return '<div class="ro"><div class="rh"><span class="rn">'+esc(b.unit)+
      (b.item?'　'+esc(b.item):'')+'</span><span class="pill '+(none?"g":"v")+'">'+
      (none?"本月無業務":"僅供檢視")+'</span></div>'+
      '<div class="rb">'+(none?"該主責已回報本月無業務。":fmtBody(b.text))+'</div>'+
      (none?"":fileChips(filesOf(b),false))+
      '<div class="rm">主責 '+esc(b.name)+'　最後更新 '+esc(b.at)+'</div></div>';
  }).join("");
}
function pickUnit(u){
  curUnit=u;
  _endArm=false;
  var b=$("endbtn");
  if(b){b.classList.remove("danger");b.textContent="結束此業務";}
  renderReport();
}
function copyPrev(){
  var m=mineOf(curUnit);
  if(!m||!m.prevText)return;
  if($("ta").value.trim()){$("confirmbar").classList.remove("hide");return;}
  doCopy();
}
function doCopy(){
  $("confirmbar").classList.add("hide");
  $("ta").value=mineOf(curUnit).prevText;
  $("ta").focus();
}
function cancelCopy(){$("confirmbar").classList.add("hide");}

function currentCo(){
  var out=[];
  var cs=$("cochips").querySelectorAll(".chip.on");
  var ps=(BOOT.people||[]).filter(function(p){return p.hr!==ME.hr&&p.role!=="管理者";});
  var all=$("cochips").querySelectorAll(".chip");
  for(var i=0;i<all.length;i++){if(all[i].classList.contains("on")&&ps[i])out.push(ps[i].hr);}
  return out;
}
function toggleCo(hr){
  var ps=(BOOT.people||[]).filter(function(p){return p.hr!==ME.hr&&p.role!=="管理者";});
  var all=$("cochips").querySelectorAll(".chip");
  for(var i=0;i<all.length;i++){if(ps[i]&&ps[i].hr===hr){all[i].classList.toggle("on");}}
}

function applyEntry(entry){
  if(!entry)return;
  for(var i=0;i<BOOT.mine.length;i++){
    if(BOOT.mine[i].unit===entry.unit){BOOT.mine[i]=entry;break;}
  }
  renderReport();
}
function saveBlock(submit){
  var msg=$("savedmsg");
  msg.classList.remove("bad");
  var text=$("ta").value.trim();
  if(!text){msg.classList.add("bad");msg.textContent="請先填寫本月業務內容。";return;}
  msg.textContent="儲存中…";
  call("api_saveBlock",ME.hr,curUnit,$("item").value.trim(),text,currentCo(),!!submit)
    .then(function(r){
      if(!r.ok){msg.classList.add("bad");msg.textContent=r.msg;return;}
      applyEntry(r.entry);
      $("savedmsg").textContent=r.msg;
    })["catch"](fail);
}
function saveDraft(){saveBlock(false);}
function submitReport(){saveBlock(true);}
var _endArm=false;
function endUnit(){
  var b=$("endbtn"),m=$("savedmsg");
  if(!_endArm){
    _endArm=true;
    b.classList.add("danger");
    b.textContent="確定結束？此項將不再出現";
    setTimeout(function(){if(_endArm){_endArm=false;b.classList.remove("danger");b.textContent="結束此業務";}},6000);
    return;
  }
  _endArm=false;b.classList.remove("danger");b.textContent="結束此業務";
  m.classList.remove("bad");m.textContent="處理中…";
  call("api_endUnit",ME.hr,curUnit).then(function(r){
    if(!r.ok){m.classList.add("bad");m.textContent=r.msg;return;}
    curUnit=null;
    return reload().then(function(){$("savedmsg").textContent=r.msg;});
  })["catch"](fail);
}
function markNone(){
  $("savedmsg").textContent="處理中…";
  call("api_markNone",ME.hr,curUnit).then(function(r){
    if(!r.ok){alert(r.msg);return;}
    applyEntry(r.entry);$("savedmsg").textContent=r.msg;
  })["catch"](fail);
}
function unmarkNone(){
  $("savedmsg").textContent="處理中…";
  call("api_unsend",ME.hr,curUnit).then(function(r){
    applyEntry(r.entry);$("savedmsg").textContent="已取消標記，可以繼續填寫。";
  })["catch"](fail);
}
function unsend(){
  $("savedmsg").textContent="處理中…";
  call("api_unsend",ME.hr,curUnit).then(function(r){
    applyEntry(r.entry);$("savedmsg").textContent=r.msg;
  })["catch"](fail);
}

/* ---------- 附件 ---------- */
function addFiles(inp){
  var files=Array.prototype.slice.call(inp.files);
  inp.value="";
  if(!files.length)return;
  $("upbtn").disabled=true;
  var seq=Promise.resolve();
  files.forEach(function(f){
    seq=seq.then(function(){return upload1(f);});
  });
  seq.then(function(){$("upbtn").disabled=false;})
     ["catch"](function(e){$("upbtn").disabled=false;fail(e);});
}
function upload1(f){
  return new Promise(function(res,rej){
    if(f.size>20*1024*1024){alert("「"+f.name+"」超過 20MB，請壓縮後再上傳。");return res();}
    var fr=new FileReader();
    fr.onload=function(){
      var b64=fr.result.split(",")[1];
      call("api_upload",ME.hr,curUnit,f.name,f.type||"application/octet-stream",b64)
        .then(function(r){if(!r.ok){alert(r.msg);}else{applyEntry(r.entry);}res();})["catch"](rej);
    };
    fr.onerror=function(){rej(new Error("讀取檔案失敗"));};
    fr.readAsDataURL(f);
  });
}
function delFile(i){
  call("api_deleteFile",ME.hr,curUnit,i).then(function(r){applyEntry(r.entry);})["catch"](fail);
}

/* ---------- 追蹤事項 ---------- */
function trackDirty(id){
  var m=$("tm"+id);
  m.classList.add("bad");
  m.textContent="尚未儲存，請按「儲存回覆」";
}
var TKID=null;
function pickTrackFile(id){TKID=id;$("trackfileinput").click();}
function uploadTrackFile(inp){
  var f=inp.files&&inp.files[0];
  inp.value="";
  if(!f||TKID===null)return;
  if(f.size>20*1024*1024){alert("「"+f.name+"」超過 20MB，請壓縮後再上傳。");return;}
  var id=TKID,btn=$("tub"+id),m=$("tm"+id);
  btn.disabled=true;m.classList.remove("bad");m.textContent="上傳中…";
  var fr=new FileReader();
  fr.onload=function(){
    call("api_uploadTrackFile",ME.hr,id,f.name,f.type||"application/octet-stream",
         fr.result.split(",")[1])
      .then(function(r){
        btn.disabled=false;
        if(!r.ok){m.classList.add("bad");m.textContent=r.msg;return;}
        m.textContent="附件已上傳。";
        $("tf"+id).innerHTML=fileChips(r.files,true,id);
        for(var i=0;i<BOOT.tracks.length;i++){
          if(String(BOOT.tracks[i].id)===String(id))BOOT.tracks[i].files=r.files;
        }
      })["catch"](function(e){btn.disabled=false;fail(e);});
  };
  fr.onerror=function(){btn.disabled=false;m.classList.add("bad");m.textContent="讀取檔案失敗。";};
  fr.readAsDataURL(f);
}
function delTrackFile(id,i){
  call("api_deleteTrackFile",ME.hr,id,i).then(function(r){
    if(!r.ok){alert(r.msg);return;}
    $("tf"+id).innerHTML=fileChips(r.files,true,id);
    for(var k=0;k<BOOT.tracks.length;k++){
      if(String(BOOT.tracks[k].id)===String(id))BOOT.tracks[k].files=r.files;
    }
  })["catch"](fail);
}
function saveTrack(id){
  var m=$("tm"+id);
  var reply=$("tr"+id).value.trim();
  if(!reply){
    m.classList.add("bad");
    m.textContent="請填寫回覆執行進度後再儲存。";
    $("tr"+id).focus();
    return;
  }
  m.classList.remove("bad");
  m.textContent="儲存中…";
  call("api_saveTrack",ME.hr,id,reply)
    .then(function(r){
      if(!r.ok){m.classList.add("bad");m.textContent=r.msg;return;}
      m.classList.remove("bad");
      m.textContent=r.msg||"已儲存";
      for(var i=0;i<BOOT.tracks.length;i++){
        if(String(BOOT.tracks[i].id)===String(id)){BOOT.tracks[i].reply=reply;}
      }
      var pill=$("tp"+id);
      if(pill){pill.textContent="執行進度";pill.className="pill v";}
    })["catch"](fail);
}

/* ---------- 申請新增業務 ---------- */
function openReq(){
  $("reqform").classList.remove("hide");
  $("reqbtn").classList.add("hide");
  var mineUnits=BOOT.mine.map(function(x){return x.unit;});
  var opts=(BOOT.units||[]).filter(function(u){return mineUnits.indexOf(u)<0;})
    .map(function(u){return '<option value="'+esc(u)+'">'+esc(u)+'</option>';}).join("");
  $("requnit").innerHTML=opts+'<option value="__new__">其他（清單中沒有，自行輸入）</option>';
  $("reqnewwrap").classList.add("hide");
  $("reqnew").value="";$("reqwhy").value="";$("reqmsg").textContent="";
}
function closeReq(){$("reqform").classList.add("hide");$("reqbtn").classList.remove("hide");}
function reqUnitChange(){$("reqnewwrap").classList.toggle("hide",$("requnit").value!=="__new__");}
function sendReq(){
  var sel=$("requnit").value, isNew=(sel==="__new__");
  var unit=isNew?$("reqnew").value.trim():sel;
  $("reqmsg").classList.remove("bad");
  $("reqmsg").textContent="處理中…";
  call("api_submitRequest",ME.hr,unit,isNew,$("reqwhy").value.trim()).then(function(r){
    if(!r.ok){$("reqmsg").classList.add("bad");$("reqmsg").textContent=r.msg;return;}
    closeReq();
    return reload().then(function(){
      curUnit=unit;renderReport();
      $("savedmsg").textContent=r.msg;
    });
  })["catch"](fail);
}
function renderMyReqs(){
  var rq=BOOT.myRequests||[];
  $("reqlist").innerHTML=rq.map(function(r){
    var cls=(r.status==="已新增"||r.status==="已核准")?"c":(r.status==="未通過"?"n":"o");
    return '<div class="rq"><span class="pill '+cls+'">'+esc(r.status)+'</span>'+
      '<span class="rqu">'+esc(r.unit)+'</span><span>'+esc(r.why)+'</span></div>';
  }).join("");
}

/* ---------- 主管檢視 ---------- */
var HIST=null;
function loadHist(y){
  return call("api_myHistory",ME.hr,y||"").then(function(r){
    if(!r.ok){alert(r.msg);return;}
    HIST=r;renderHist();
  })["catch"](fail);
}
function renderHist(){
  var r=HIST;
  if(!r.years.length){
    $("histnote").textContent="目前還沒有已送出的紀錄。";
    $("histyears").innerHTML="";$("histunits").innerHTML="";$("histlist").innerHTML="";
    return;
  }
  $("histnote").textContent="這裡整理你歷年送出的業務報告，含已結束的業務，可作為年度回顧使用。";
  $("histyears").innerHTML=r.years.map(function(y){
    return '<button class="'+(y===r.year?"on":"")+'" onclick="loadHist(\''+y+'\')">'+esc(y)+' 年</button>';
  }).join("");
  $("histyears").classList.toggle("hide",r.years.length<2);

  $("histunits").innerHTML=r.units.length
    ? '<div class="hsum"><div class="hl">'+esc(r.year)+' 年度負責過的業務（虛線為已結束或當年度才有）</div>'+
      r.units.map(function(u){
        var on=(r.active.indexOf(u)>=0);
        return '<span class="hu'+(on?"":" off")+'">'+esc(u)+'</span>';
      }).join("")+'</div>'
    : "";

  $("histlist").innerHTML=r.months.map(function(m){
    return '<div class="hm"><div class="hmy">'+esc(m.ym)+'</div>'+
      m.items.map(function(it){
        return '<div class="hi2"><div class="hiu">'+esc(it.unit)+'</div>'+
          '<div class="hib">'+fmtBody(it.text)+'</div>'+
          (it.files&&it.files.length?fileChips(it.files,false):"")+
          (it.co?'<div class="him">協同　'+esc(it.co)+'</div>':'')+
          '</div>';
      }).join("")+'</div>';
  }).join("");
}
function loadOver(){
  return call("api_overview",ME.hr).then(function(r){
    if(!r.ok){alert(r.msg);return;}
    OVER=r;oIdx=0;renderOver();
  })["catch"](fail);
}
function renderOver(){
  var d=OVER.data||[];
  $("overnote").textContent="可檢視 "+d.length+" 組已送出的業務報告。";
  $("overbar").innerHTML=d.map(function(x,i){
    return '<button class="'+(i===oIdx?"on":"")+'" onclick="pickOver('+i+')">'+esc(x.unit)+'</button>';
  }).join("");
  $("overbar").classList.toggle("hide",d.length<2);
  if(!d.length){$("overtitle").textContent="";$("overlist").innerHTML="";return;}
  var u=d[oIdx];
  $("overtitle").textContent=u.unit;
  $("oversub").textContent=u.owners.length?("主責　"+u.owners.map(function(o){return o.name;}).join("、")):"尚未指定主責";
  $("overlist").innerHTML=u.owners.map(function(o){
    var sent=o.status==="已送出";
    return '<div class="ro"><div class="rh"><span class="rn">'+
      esc(o.name)+'</span>'+
      '<span class="pill '+(sent?"v":(o.status==="本月無業務"?"g":"n"))+'">'+esc(o.status)+'</span>'+
      ((sent&&o.co&&o.co.length)?'<span class="cow">協同　'+esc(o.co.join("、"))+'</span>':'')+'</div>'+
      '<div class="rb">'+(sent?fmtBody(o.text):(o.status==="本月無業務"?"本月無業務。":"本月尚未送出。"))+'</div>'+
      (sent?fileChips(o.files,false):"")+
      (sent?'<div class="rm">最後更新 '+esc(o.at)+'</div>':'')+'</div>';
  }).join("");
}
function pickOver(i){oIdx=i;renderOver();}

/* ---------- 管理者 ---------- */
function loadAdmin(){
  return call("api_admin",ME.hr).then(function(r){
    if(!r.ok){alert(r.msg);return;}
    ADMIN=r;renderAdmin();
  })["catch"](fail);
}
function renderAdmin(){
  var r=ADMIN;
  $("stats").innerHTML=
    '<div class="stat"><div class="n">'+r.total+'</div><div class="l">應填報區塊</div></div>'+
    '<div class="stat"><div class="n">'+r.fin+'</div><div class="l">已完成</div></div>'+
    '<div class="stat w"><div class="n">'+(r.total-r.fin)+'</div><div class="l">尚未完成</div></div>'+
    '<div class="stat w"><div class="n">'+r.tracks.filter(function(t){return t.status!=="已結案";}).length+'</div><div class="l">未結案追蹤事項</div></div>';

  var last="";
  $("admtb").innerHTML=r.rows.map(function(x){
    if(x.inactive){
      return '<tr><td class="u">'+esc(x.unit)+'</td><td class="p">'+esc(x.owners.join("、")||"—")+
        '</td><td class="p">—</td><td class="p">每季報告（'+x.months.join("、")+' 月）</td><td class="p">—</td><td class="p">—</td></tr>';
    }
    var show=(x.unit!==last);last=x.unit;
    var d=x.status==="已送出"?"y":(x.status==="本月無業務"?"g":(x.status==="暫存中"?"d":"n"));
    return '<tr><td class="u">'+(show?esc(x.unit):"")+'</td>'+
      '<td class="p">'+esc(x.name)+'</td>'+
      '<td class="p">'+(esc(x.item)||"—")+'</td>'+
      '<td><span class="dot '+d+'"></span>'+esc(x.status)+'</td>'+
      '<td class="p">'+(esc(x.at)||"—")+'</td>'+
      '<td>'+((x.status==="已送出"||x.status==="本月無業務")?"—":
        '<button class="mini" onclick="remindOne(\''+x.unit+'\',\''+x.hr+'\')">寄送催繳</button>'+
        (x.noted?'<div class="noted">已於 '+esc(x.noted)+' 通知</div>':''))+'</td></tr>';
  }).join("");

  $("reqcount").textContent=r.requests.length
    ? ("同仁自行新增的業務共 "+r.requests.length+" 筆；若有誤可在此移除。")
    : "目前沒有同仁自行新增的業務。";
  $("reqtb").innerHTML=r.requests.length?r.requests.map(function(q){
    var cls=(q.status==="已新增"||q.status==="已核准")?"c":(q.status==="未通過"?"n":"o");
    return '<tr><td class="p">'+esc(q.at)+'</td><td class="u">'+esc(q.name)+'</td>'+
      '<td>'+esc(q.unit)+(q.isNew?' <span class="cnt">新業務</span>':'')+'</td>'+
      '<td class="p">'+(esc(q.why)||"—")+'</td><td><span class="pill '+cls+'">'+esc(q.status)+'</span></td>'+
      '<td class="acts"><button class="mini danger" id="rq'+q.id+'" onclick="removeUnit('+q.id+')">移除</button></td></tr>';
  }).join(""):'<tr><td class="p" colspan="6">尚無紀錄。</td></tr>';
  $("reqadmmsg").textContent="";

  $("trktb").innerHTML=r.tracks.map(function(t){
    var done=(t.status==="已結案");
    var lab=tkLabel(t.status,t.reply);
    return '<tr><td>'+esc(t.q)+'</td><td class="p">'+esc(t.owner)+'</td>'+
      '<td><span class="pill '+(done?"c":(lab==="執行進度"?"v":"n"))+'">'+esc(lab)+'</span></td>'+
      '<td class="p">'+((t.replies&&t.replies.length)
        ? t.replies.map(function(p){return (p.who?'<b>'+esc(p.who)+'</b>：':'')+esc(p.text);}).join("<br>")
        : "—")+
      ((t.files&&t.files.length)?'<div class="tfs">'+t.files.map(function(f){
        return '📎 <a href="'+esc(f.url)+'" target="_blank">'+esc(f.n)+'</a>';}).join('<br>')+'</div>':'')+
      '</td>'+
      '<td class="acts">'+(done
        ? '<button class="mini" onclick="closeTrack('+t.id+',false)">重新開啟</button>'
        : '<button class="mini" onclick="closeTrack('+t.id+',true)">結案</button>')+'</td></tr>';
  }).join("");
  $("trkmsg").textContent="";
  $("admmsg").textContent="";
}
function remindOne(unit,hr){
  call("api_remindOne",ME.hr,unit,hr).then(function(r){
    $("admmsg").textContent=r.msg;return loadAdmin();
  })["catch"](fail);
}
function remindAll(){
  call("api_remindAll",ME.hr).then(function(r){
    $("admmsg").textContent=r.msg;return loadAdmin();
  })["catch"](fail);
}
function notifyMgr(){
  call("api_notifyManagers",ME.hr).then(function(r){$("admmsg").textContent=r.msg;})["catch"](fail);
}
var _rmArm=null;
function removeUnit(id){
  var b=$("rq"+id),m=$("reqadmmsg");
  if(_rmArm!==id){
    if(_rmArm!==null&&$("rq"+_rmArm)){$("rq"+_rmArm).textContent="移除";}
    _rmArm=id;b.textContent="確定移除？";
    setTimeout(function(){if(_rmArm===id&&$("rq"+id)){$("rq"+id).textContent="移除";_rmArm=null;}},5000);
    return;
  }
  _rmArm=null;
  m.classList.remove("bad");m.textContent="處理中…";
  call("api_removeUnit",ME.hr,id).then(function(r){
    if(!r.ok){m.classList.add("bad");m.textContent=r.msg;return;}
    return loadAdmin().then(function(){$("reqadmmsg").textContent=r.msg;});
  })["catch"](fail);
}
var _delArm=null;
function delReq(id){
  var b=$("dq"+id);
  if(_delArm!==id){
    if(_delArm!==null&&$("dq"+_delArm)){$("dq"+_delArm).textContent="刪除";}
    _delArm=id;b.textContent="確定刪除？";
    setTimeout(function(){if(_delArm===id&&$("dq"+id)){$("dq"+id).textContent="刪除";_delArm=null;}},5000);
    return;
  }
  _delArm=null;
  call("api_deleteRequest",ME.hr,id).then(function(r){
    $("admmsg").textContent=r.msg;
    if(r.ok)return loadAdmin();
  })["catch"](fail);
}
function closeTrack(id,close){
  var m=$("trkmsg");
  m.classList.remove("bad");
  m.textContent="處理中…";
  call("api_closeTrack",ME.hr,id,close).then(function(r){
    if(!r.ok){m.classList.add("bad");m.textContent=r.msg;return;}
    m.textContent=r.msg;
    return loadAdmin().then(function(){$("trkmsg").textContent=r.msg;});
  })["catch"](fail);
}
function expo(){
  $("admmsg").textContent="產生中，請稍候…";
  call("api_export",ME.hr).then(function(r){
    if(!r.ok){$("admmsg").textContent=r.msg;return;}
    $("admmsg").innerHTML='已產生：<a href="'+r.docx+'" target="_blank">下載 Word 檔</a>　'+
      '<a href="'+r.url+'" target="_blank">線上檢視</a>';
  })["catch"](fail);
}

/* ---------- 投影 ---------- */
function enterProj(){
  call("api_overview",ME.hr).then(function(r){
    if(!r.ok){alert(r.msg);return;}
    OVER=r;pIdx=0;
    $("v-app").classList.add("hide");
    $("v-proj").classList.remove("hide");
    $("projym").textContent=OVER.ym;
    pdraw();
  })["catch"](fail);
}
var PORTRAIT=false;
function toggleOrient(){
  PORTRAIT=!PORTRAIT;
  $("projwrap").classList.toggle("pt9",PORTRAIT);
  $("orbtn").textContent=PORTRAIT?"切換橫式":"切換直式";
}
function exitProj(){$("v-proj").classList.add("hide");$("v-app").classList.remove("hide");}
function projUnits(){
  if(!OVER||!OVER.data)return [];
  return OVER.data.filter(function(u){
    return u.owners.some(function(o){return o.status==="已送出"&&o.text;});
  });
}
function pageCount(){return 1+projUnits().length;}
function pmove(d){
  var n=pageCount();if(!n)return;
  pIdx=(pIdx+d+n)%n;pdraw();
}
function pdraw(){
  if(pIdx===0){drawTracks();$("ppg").textContent="1 / "+pageCount();return;}
  var list=projUnits();
  var u=list[pIdx-1];
  if(!u){
    $("pu").textContent="本月沒有已送出的業務報告";
    $("po").textContent="";
    $("pb").textContent="";
    $("pb").className="body empty";
    $("ppg").textContent=(pIdx+1)+" / "+pageCount();
    return;
  }
  $("pu").textContent=u.unit;
  $("po").textContent=u.owners.length?("主責　"+u.owners.map(function(o){return o.name;}).join("、")):"";
  var blocks=u.owners.filter(function(o){return o.status==="已送出"&&o.text;});
  var allco=[];
  blocks.forEach(function(o){
    (o.co||[]).forEach(function(c){if(allco.indexOf(c)<0)allco.push(c);});
  });
  $("po").textContent=blocks.length
    ?("主責　"+blocks.map(function(o){return o.name;}).join("、")+
      (allco.length?"　　協同　"+allco.join("、"):""))
    :"";
  if(!blocks.length){
    $("pb").textContent="";
    $("pb").className="body empty";
  }else{
    $("pb").innerHTML=blocks.map(function(o){
      return '<div class="pblk"><div class="phrow"><span class="ph">'+esc(o.name)+'</span>'+
        ((o.co&&o.co.length)?'<span class="pco">協同　'+esc(o.co.join("、"))+'</span>':'')+'</div>'+
        '<div class="pt">'+fmtBody(o.text)+'</div>'+
        (o.files&&o.files.length?'<div class="pf">'+o.files.map(function(f){
          return '📎 <a href="'+esc(f.url)+'" target="_blank">'+esc(f.n)+'</a>';}).join('　')+'</div>':'')+
        '</div>';
    }).join("");
    $("pb").className="body";
  }
  $("ppg").textContent=(pIdx+1)+" / "+pageCount();
}
function drawTracks(){
  var t=(OVER&&OVER.tracks)?OVER.tracks:[];
  $("pu").textContent="上次會議追蹤事項";
  $("po").textContent=t.length?("共 "+t.length+" 項"):"";
  if(!t.length){
    $("pb").textContent="目前沒有待追蹤的事項";
    $("pb").className="body empty";
    return;
  }
  $("pb").innerHTML=t.map(function(x){
    var done=(x.status==="已結案");
    var lab=tkLabel(x.status,x.reply);
    return '<div class="tk"><div class="tkh">'+
      '<span class="tkn'+(done?" done":"")+'">'+esc(x.owner||"—")+'</span>'+
      '<span class="tkn'+(done?" done":"")+'">'+esc(lab)+'</span></div>'+
      '<div class="tkq">'+esc(x.q)+'</div>'+
      ((x.replies&&x.replies.length)
        ? x.replies.map(function(p){
            var showWho=(x.replies.length>1&&p.who);
            return '<div class="tkr">'+(showWho?'<div class="tkw">'+esc(p.who)+'</div>':'')+fmtBody(p.text)+'</div>';
          }).join("")
        : '<div class="tkr empty">尚未回覆</div>')+
      ((x.files&&x.files.length)?'<div class="pf">'+x.files.map(function(f){
        return '📎 <a href="'+esc(f.url)+'" target="_blank">'+esc(f.n)+'</a>';}).join('　')+'</div>':'')+
      '</div>';
  }).join("");
  $("pb").className="body";
}
document.addEventListener("keydown",function(e){
  if($("v-proj").classList.contains("hide"))return;
  if(e.key==="ArrowRight")pmove(1);
  if(e.key==="ArrowLeft")pmove(-1);
  if(e.key==="Escape")exitProj();
});
</script>
</body>
</html>
`;
}
