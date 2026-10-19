/**
 * 訓練記錄 API（Google Apps Script）
 *
 * 這支腳本綁定在「訓練追蹤」試算表上，提供手機 PWA 讀寫用的 JSON API。
 *   GET  ?action=bootstrap&day=Push%20A&date=2026-10-08&token=…  取得該訓練日每個動作的上次紀錄、今日已記錄的組與建議
 *   GET  ?action=ping&token=…                                    測試連線
 *   POST {action:'set', token, set_id, …}                       寫入一組（有氧／計時也走這裡）
 *   POST {action:'undo', token, set_id}                         清除某一組
 *   POST {action:'update', token, set_id, …}                    修改某一組
 *   POST {action:'bodyweight', token, date, kg}                 記錄體重
 *
 * 寫入時只填「輸入欄」，預估1RM、組合鍵、有效組等公式欄本來就預填到第 3004 列，會自己計算。
 */

const SHEET_LOG = '訓練紀錄';
const SHEET_EX = '動作清單';
const SHEET_BW = '體重記錄';
const SHEET_PLAN = '訓練菜單'; // 每個訓練日排哪些動作（A 訓練日、B 動作、C 正式組數、D 每組次數、E 備註），資料從第 5 列
const FIRST_ROW = 5;        // 資料從第 5 列開始
const LAST_ROW = 3004;      // 公式預填到第 3004 列
const NCOL = 41;            // A:AO
const COL_SETID = 42;       // AP：set_id（程式寫入，用來撤銷與防重複）
const COL_WRITTEN = 43;     // AQ：寫入時間
const COL_SOURCE = 44;      // AR：來源

// 輸入欄（1-based 欄號）
const C = {
  date: 1, day: 2, ex: 3, setNo: 4, setType: 5, kg: 6, reps: 7, rir: 8,
  note: 11, type: 12,
  z2min: 13, z2km: 14, z2inc: 15,
  sprSec: 16, sprM: 17, sprInc: 18, jogSec: 19, jogM: 20, jogInc: 21, rounds: 22,
  timedSec: 23,
  e1rm: 9, effSet: 32,
};

// ───────────────────────── 進入點 ─────────────────────────

function doGet(e) {
  return handle_(function () {
    const p = (e && e.parameter) || {};
    auth_(p.token);
    switch (p.action) {
      case 'ping': return { ok: true, time: new Date().toISOString() };
      case 'bootstrap': return bootstrap_(p.day, p.date);
      default: throw new ApiError('unknown_action', '未知的 action：' + p.action);
    }
  });
}

function doPost(e) {
  return handle_(function () {
    let body;
    try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
    catch (err) { throw new ApiError('bad_json', 'JSON 格式錯誤'); }
    auth_(body.token);
    switch (body.action) {
      case 'set': return withLock_(function () { return writeSet_(body); });
      case 'undo': return withLock_(function () { return undoSet_(body.set_id); });
      case 'update': return withLock_(function () { return updateSet_(body); });
      case 'bodyweight': return withLock_(function () { return writeBodyweight_(body); });
      case 'addExercise': return withLock_(function () { return addExercise_(body); });
      default: throw new ApiError('unknown_action', '未知的 action：' + body.action);
    }
  });
}

// ───────────────────────── 共用 ─────────────────────────

function ApiError(code, message) { this.code = code; this.message = message; }

function handle_(fn) {
  let out;
  try { out = fn(); }
  catch (err) {
    out = { ok: false, error: err.code || 'server_error', message: err.message || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function auth_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!expected) throw new ApiError('not_setup', '尚未執行 setup()');
  if (!token || token !== expected) throw new ApiError('unauthorized', '權杖錯誤');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new ApiError('busy', '伺服器忙碌，請重試');
  try { return fn(); }
  finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
let TZ_ = null; // 每次執行只問一次時區（逐列呼叫會慢好幾秒）
function tz_() { return TZ_ || (TZ_ = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei'); }

/** 'yyyy-MM-dd' → 試算表時區當天 00:00 的 Date */
function parseDate_(s) {
  if (s instanceof Date) return s;
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(s || ''));
  if (!m) throw new ApiError('bad_date', '日期格式應為 yyyy-MM-dd');
  // Utilities.parseDate 依指定時區解析，避免跨日
  return Utilities.parseDate(m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]) + ' 00:00:00', tz_(), 'yyyy-MM-dd HH:mm:ss');
}
function fmtDate_(d) { return d instanceof Date ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : String(d || ''); }
function pad2_(x) { x = String(x); return x.length < 2 ? '0' + x : x; }
function num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isFinite(n) ? n : '';
}

// ───────────────────────── 讀取 ─────────────────────────

/**
 * 讀整張訓練紀錄很慢（3000 列約 4 秒），所以解析後的結果放在 CacheService：
 *  - 第一次讀取時從試算表讀，之後直接用快取
 *  - App 每次寫入／修改／撤銷，只重讀那一列並更新快取
 *  - 你在試算表手動編輯「訓練紀錄」時，onEdit 會清掉快取，下次重新讀
 */
const CACHE_KEY = 'log_v2';
const CACHE_TTL = 21600; // 6 小時（CacheService 上限）
const CHUNK = 90000;     // 單一快取值上限 100KB

function readLog_() {
  const cached = cacheGet_();
  if (cached) return cached.rows;
  return loadLogFromSheet_().rows;
}

function loadLogFromSheet_() {
  const sh = ss_().getSheetByName(SHEET_LOG);
  const values = sh.getRange(FIRST_ROW, 1, LAST_ROW - FIRST_ROW + 1, COL_SOURCE).getValues();
  const rows = [];
  let lastRow = FIRST_ROW - 1;
  for (let i = 0; i < values.length; i++) {
    if (values[i][C.ex - 1] !== '') lastRow = FIRST_ROW + i;
    const o = parseRow_(values[i], FIRST_ROW + i);
    if (o) rows.push(o);
  }
  const data = { rows: rows, lastRow: lastRow };
  cachePut_(data);
  return data;
}

function cacheGet_() {
  try {
    const cache = CacheService.getScriptCache();
    const head = cache.get(CACHE_KEY);
    if (!head) return null;
    const n = Number(head);
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
    const parts = cache.getAll(keys);
    let s = '';
    for (let i = 0; i < n; i++) { if (parts[keys[i]] == null) return null; s += parts[keys[i]]; }
    return JSON.parse(s);
  } catch (err) { return null; }
}

function cachePut_(data) {
  try {
    const s = JSON.stringify(data);
    const obj = {};
    let n = 0;
    for (let i = 0; i < s.length; i += CHUNK) obj[CACHE_KEY + '_' + (n++)] = s.slice(i, i + CHUNK);
    obj[CACHE_KEY] = String(n);
    CacheService.getScriptCache().putAll(obj, CACHE_TTL);
  } catch (err) { cacheClear_(); }
}

function cacheClear_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (err) {}
}

/** 寫入後只重讀那一列，更新快取（沒有快取就不用管，下次讀會重建） */
function cachePatchRow_(sh, row) {
  const data = cacheGet_();
  if (!data) return;
  const o = parseRow_(sh.getRange(row, 1, 1, COL_SOURCE).getValues()[0], row);
  data.rows = data.rows.filter(function (r) { return r.row !== row; });
  if (o) {
    data.rows.push(o);
    data.rows.sort(function (a, b) { return a.row - b.row; });
    if (row > data.lastRow) data.lastRow = row;
  } else if (row === data.lastRow) {
    data.lastRow = data.rows.length ? data.rows[data.rows.length - 1].row : FIRST_ROW - 1;
  }
  cachePut_(data);
}

/** 試算表手動編輯時清快取（simple trigger，不需另外設定） */
function onEdit(e) {
  try {
    const name = e && e.range ? e.range.getSheet().getName() : '';
    if (!name || name === SHEET_LOG) cacheClear_();
  } catch (err) { cacheClear_(); }
}

function parseRow_(r, rowNum) {
  if (r[C.ex - 1] === '' || !(r[C.date - 1] instanceof Date)) return null;
  return {
      row: rowNum,
      date: fmtDate_(r[C.date - 1]),
      day: String(r[C.day - 1] || '').trim() || '（未指定）',
      ex: String(r[C.ex - 1]).trim(),
      setNo: num_(r[C.setNo - 1]),
      setType: String(r[C.setType - 1] || '').trim() || '正式',
      kg: num_(r[C.kg - 1]), reps: num_(r[C.reps - 1]), rir: num_(r[C.rir - 1]),
      note: String(r[C.note - 1] || ''),
      type: String(r[C.type - 1] || '重量型'),
      z2min: num_(r[C.z2min - 1]), z2km: num_(r[C.z2km - 1]), z2inc: num_(r[C.z2inc - 1]),
      sprSec: num_(r[C.sprSec - 1]), sprM: num_(r[C.sprM - 1]), sprInc: num_(r[C.sprInc - 1]),
      jogSec: num_(r[C.jogSec - 1]), jogM: num_(r[C.jogM - 1]), jogInc: num_(r[C.jogInc - 1]),
      rounds: num_(r[C.rounds - 1]), timedSec: num_(r[C.timedSec - 1]),
      e1rm: num_(r[C.e1rm - 1]),
      setId: String(r[COL_SETID - 1] || ''),
  };
}

function readExercises_() {
  const sh = ss_().getSheetByName(SHEET_EX);
  const v = sh.getRange(5, 2, 50, 2).getValues();
  const list = [];
  v.forEach(function (r) {
    const name = String(r[0] || '').trim();
    if (name) list.push({ name: name, type: String(r[1] || '').trim() || '重量型' });
  });
  return list;
}

function readDays_() {
  const sh = ss_().getSheetByName(SHEET_EX);
  return sh.getRange(13, 10, 20, 1).getValues()
    .map(function (r) { return String(r[0] || '').trim(); })
    .filter(function (s) { return s && s.indexOf('（') !== 0; });
}

/** 訓練菜單：回傳某訓練日的動作（依列順序）。沒有這個分頁就回傳空陣列 */
function readPlan_(day) {
  const sh = ss_().getSheetByName(SHEET_PLAN);
  if (!sh || sh.getMaxRows() < 5) return [];
  const v = sh.getRange(5, 1, sh.getMaxRows() - 4, 5).getDisplayValues();
  const out = [];
  v.forEach(function (r) {
    const d = String(r[0]).trim(), ex = String(r[1]).trim();
    if (!ex || d !== day) return;
    if (out.some(function (x) { return x.ex === ex; })) return;
    out.push({
      ex: ex,
      sets: Number(r[2]) > 0 ? Number(r[2]) : null,
      reps: String(r[3]).split(/[,，、]/).map(function (x) { return x.trim(); }).filter(String),
      note: String(r[4]).trim(),
    });
  });
  return out;
}

function latestBodyweight_() {
  const sh = ss_().getSheetByName(SHEET_BW);
  const v = sh.getRange(5, 1, sh.getMaxRows() - 4, 2).getValues();
  let best = null;
  v.forEach(function (r) {
    if (r[0] instanceof Date && typeof r[1] === 'number' && (!best || r[0] > best.d)) best = { d: r[0], kg: r[1] };
  });
  return best ? { date: fmtDate_(best.d), kg: best.kg } : null;
}

// ───────────────────────── bootstrap ─────────────────────────

function bootstrap_(day, dateStr) {
  const today = dateStr || fmtDate_(new Date());
  const rows = readLog_();
  const exList = readExercises_();
  const typeOf = {};
  exList.forEach(function (x) { typeOf[x.name] = x.type; });
  const result = {
    ok: true, today: today, days: readDays_(), exercises: exList,
    bodyweight: latestBodyweight_(), lastDayDone: lastDayDates_(rows), muscles: readMuscles_(),
  };
  if (!day) return result;

  const dayRows = rows.filter(function (r) { return r.day === day; });
  const plan = readPlan_(day);
  // 依「最近一次該訓練日」的順序排動作，再補上更早出現過的動作
  const pastDates = unique_(dayRows.filter(function (r) { return r.date < today; })
    .map(function (r) { return r.date; })).sort().reverse();
  // 先照「訓練菜單」的順序，再補上菜單外、最近做過的動作
  const order = plan.map(function (p) { return p.ex; });
  pastDates.slice(0, 4).forEach(function (d) {
    dayRows.filter(function (r) { return r.date === d; })
      .sort(function (a, b) { return a.row - b.row; })
      .forEach(function (r) { if (order.indexOf(r.ex) < 0) order.push(r.ex); });
  });
  // 今天臨時加做的動作
  dayRows.filter(function (r) { return r.date === today; })
    .forEach(function (r) { if (order.indexOf(r.ex) < 0) order.push(r.ex); });

  result.day = day;
  result.items = order.map(function (ex) {
    const it = buildItem_(ex, typeOf[ex] || '重量型', dayRows.filter(function (r) { return r.ex === ex; }), today);
    const p = plan.filter(function (x) { return x.ex === ex; })[0];
    it.plan = p ? { sets: p.sets, reps: p.reps, note: p.note } : null;
    return it;
  });
  return result;
}

/** 每個訓練日最近一次的日期，給前端推算「下一個訓練日」 */
function lastDayDates_(rows) {
  const out = {};
  rows.forEach(function (r) { if (!out[r.day] || r.date > out[r.day]) out[r.day] = r.date; });
  return out;
}

function buildItem_(ex, type, exRows, today) {
  const prevDates = unique_(exRows.filter(function (r) { return r.date < today; }).map(function (r) { return r.date; })).sort();
  const lastDate = prevDates.length ? prevDates[prevDates.length - 1] : null;
  const pick = function (d) {
    return exRows.filter(function (r) { return r.date === d; })
      .sort(function (a, b) { return a.row - b.row; })
      .map(slimSet_);
  };
  const last = lastDate ? pick(lastDate) : [];
  const prev2Date = prevDates.length > 1 ? prevDates[prevDates.length - 2] : null;
  return {
    ex: ex, type: type,
    lastDate: lastDate, last: last,
    prev2Date: prev2Date, prev2: prev2Date ? pick(prev2Date) : [],
    todaySets: pick(today),
    best: bestE1rm_(exRows, today),
    suggestion: suggest_(type, last),
  };
}

function slimSet_(r) {
  const o = { setId: r.setId, row: r.row, setNo: r.setNo, setType: r.setType };
  ['kg', 'reps', 'rir', 'z2min', 'z2km', 'z2inc', 'sprSec', 'sprM', 'sprInc', 'jogSec', 'jogM', 'jogInc', 'rounds', 'timedSec', 'e1rm']
    .forEach(function (k) { if (r[k] !== '') o[k] = r[k]; });
  if (r.note) o.note = r.note;
  return o;
}

function bestE1rm_(exRows, today) {
  let best = 0;
  exRows.forEach(function (r) { if (r.date < today && typeof r.e1rm === 'number' && r.e1rm > best) best = r.e1rm; });
  return best || null;
}

/**
 * 進步建議（只是建議，畫面上可以改）：
 *  - 上次所有正式組 RIR ≥ 3 → 加重（< 15 kg 加 1 kg，其餘加 2.5 kg；自重型加在負重上）
 *  - 上次有正式組 RIR ≤ 1 且次數比第一組少 2 下以上 → 維持重量、維持次數
 *  - 其他 → 維持重量，每組目標 +1 下（雙重進步）
 *  計時型：最長一組 +5 秒；有氧：沿用上次。
 */
function suggest_(type, last) {
  if (!last.length) return null;
  if (type === '計時型') {
    const mx = Math.max.apply(null, last.map(function (s) { return Number(s.timedSec) || 0; }));
    return { kind: 'time', timedSec: mx + 5, text: '每組目標 ' + (mx + 5) + ' 秒' };
  }
  if (type.indexOf('有氧') === 0) return { kind: 'cardio', text: '沿用上次設定' };

  const work = last.filter(function (s) { return s.setType === '正式' && s.reps; });
  if (!work.length) return null;
  const kg = Number(work[0].kg) || 0;
  const withRir = work.filter(function (s) { return s.rir !== undefined; });
  const allEasy = withRir.length === work.length && withRir.every(function (s) { return s.rir >= 3; });
  const hardDrop = work.some(function (s) { return s.rir !== undefined && s.rir <= 1 && s.reps <= work[0].reps - 2; });

  if (allEasy) {
    const inc = (type !== '自重型' && kg < 15) ? 1 : 2.5;
    return { kind: 'add_weight', kg: round_(kg + inc), repsDelta: -1,
      text: '上次每組都還剩 3 下以上，加重到 ' + round_(kg + inc) + ' kg' };
  }
  if (hardDrop) {
    return { kind: 'hold', kg: kg, repsDelta: 0, text: '上次後段掉次數較多，維持 ' + kg + ' kg 先把次數做穩' };
  }
  return { kind: 'add_rep', kg: kg, repsDelta: 1, text: '維持 ' + kg + ' kg，每組目標多 1 下' };
}

function round_(x) { return Math.round(x * 100) / 100; }
function unique_(arr) { const s = {}; return arr.filter(function (x) { return s[x] ? false : (s[x] = true); }); }

// ───────────────────────── 寫入 ─────────────────────────

/** 找已存在的 set_id（只看程式寫入的列），回傳列號或 0 */
function findSetRow_(sh, setId) {
  if (!setId) return 0;
  const data = cacheGet_();
  if (data) {
    for (let i = data.rows.length - 1; i >= 0; i--) if (data.rows[i].setId === String(setId)) return data.rows[i].row;
    return 0;
  }
  const ids = sh.getRange(FIRST_ROW, COL_SETID, LAST_ROW - FIRST_ROW + 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(setId)) return FIRST_ROW + i;
  return 0;
}

/** 最後一筆資料的下一列（中間被撤銷留下的空列不回填，維持時間順序） */
function nextEmptyRow_(sh) {
  const data = cacheGet_();
  if (data) return data.lastRow + 1;
  const col = sh.getRange(FIRST_ROW, C.ex, LAST_ROW - FIRST_ROW + 1, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) if (col[i][0] !== '') return FIRST_ROW + i + 1;
  return FIRST_ROW;
}

function writeSet_(b) {
  if (!b.set_id) throw new ApiError('missing_set_id', '缺少 set_id');
  if (!b.ex) throw new ApiError('missing_ex', '缺少動作名稱');
  const sh = ss_().getSheetByName(SHEET_LOG);
  const exists = findSetRow_(sh, b.set_id);
  if (exists) return { ok: true, duplicate: true, row: exists, set_id: b.set_id };

  const row = nextEmptyRow_(sh);
  if (row > LAST_ROW) throw new ApiError('full', '訓練紀錄已滿 3000 列，請先擴充公式範圍');
  writeInputs_(sh, row, b);
  sh.getRange(row, COL_SETID, 1, 3).setValues([[String(b.set_id), new Date(), b.source || 'pwa']]);
  cachePatchRow_(sh, row);
  return { ok: true, duplicate: false, row: row, set_id: b.set_id };
}

function writeInputs_(sh, row, b) {
  const date = parseDate_(b.date);
  const rir = num_(b.rir);
  if (rir !== '' && (rir < 0 || rir > 5)) throw new ApiError('bad_rir', 'RIR 必須是 0–5');
  // A:H
  sh.getRange(row, 1, 1, 8).setValues([[
    date, String(b.day || ''), String(b.ex), num_(b.setNo), String(b.setType || '正式'),
    num_(b.kg), num_(b.reps), rir,
  ]]);
  // K 備註
  sh.getRange(row, C.note).setValue(String(b.note || ''));
  // M:W 有氧／HIIT／計時
  sh.getRange(row, C.z2min, 1, 11).setValues([[
    num_(b.z2min), num_(b.z2km), num_(b.z2inc),
    num_(b.sprSec), num_(b.sprM), num_(b.sprInc),
    num_(b.jogSec), num_(b.jogM), num_(b.jogInc),
    num_(b.rounds), num_(b.timedSec),
  ]]);
}

function undoSet_(setId) {
  const sh = ss_().getSheetByName(SHEET_LOG);
  const row = findSetRow_(sh, setId);
  if (!row) return { ok: true, notFound: true, set_id: setId };
  clearInputs_(sh, row);
  sh.getRange(row, COL_SETID, 1, 3).clearContent();
  cachePatchRow_(sh, row);
  return { ok: true, row: row, set_id: setId };
}

function clearInputs_(sh, row) {
  sh.getRange(row, 1, 1, 8).clearContent();
  sh.getRange(row, C.note).clearContent();
  sh.getRange(row, C.z2min, 1, 11).clearContent();
}

function updateSet_(b) {
  const sh = ss_().getSheetByName(SHEET_LOG);
  const row = findSetRow_(sh, b.set_id);
  if (!row) throw new ApiError('not_found', '找不到這一組（可能還沒同步）');
  writeInputs_(sh, row, b);
  sh.getRange(row, COL_WRITTEN).setValue(new Date());
  cachePatchRow_(sh, row);
  return { ok: true, row: row, set_id: b.set_id };
}

function writeBodyweight_(b) {
  const kg = num_(b.kg);
  if (kg === '' || kg <= 0) throw new ApiError('bad_kg', '體重格式錯誤');
  const date = parseDate_(b.date);
  const sh = ss_().getSheetByName(SHEET_BW);
  const v = sh.getRange(5, 1, sh.getMaxRows() - 4, 1).getValues();
  let target = 0, lastFilled = 4;
  for (let i = 0; i < v.length; i++) {
    if (v[i][0] instanceof Date) {
      lastFilled = 5 + i;
      if (fmtDate_(v[i][0]) === fmtDate_(date)) target = 5 + i;   // 同一天重記就覆蓋
    }
  }
  const row = target || lastFilled + 1;
  if (row > 1004) throw new ApiError('full', '體重記錄已滿（公式只看到第 1004 列）');
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
  sh.getRange(row, 1, 1, 2).setValues([[date, kg]]);
  return { ok: true, row: row, replaced: !!target };
}

// ───────────────────────── 設定 ─────────────────────────

/**
 * 第一次使用時在編輯器執行一次：
 *  1. 產生權杖並存起來（已存在就沿用）
 *  2. 在「訓練紀錄」最右邊加上 set_id／寫入時間／來源三欄
 *  3. 把試算表時區設成台北
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty('TOKEN', token);
  }
  const ss = ss_();
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  TZ_ = null;
  cacheClear_();
  const sh = ss.getSheetByName(SHEET_LOG);
  if (sh.getMaxColumns() < COL_SOURCE) sh.insertColumnsAfter(sh.getMaxColumns(), COL_SOURCE - sh.getMaxColumns());
  sh.getRange(4, COL_SETID, 1, 3).setValues([['set_id', '寫入時間', '來源']]);
  sh.getRange(3, COL_SETID).setValue('▼ 手機 App 寫入（請勿修改）');
  Logger.log('權杖：' + token);
  return token;
}

/** 部署後執行：在記錄裡印出要貼到手機 App 的設定碼（網址|權杖） */
function showSetupCode() {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN') || setup();
  const url = ScriptApp.getService().getUrl();
  Logger.log('設定碼（整行複製到手機 App）：\n' + url + '|' + token);
}

/** 不經過網路，在編輯器裡測試讀取 */
function testBootstrap() {
  const r = bootstrap_('Push A', fmtDate_(new Date()));
  Logger.log(JSON.stringify(r.items.map(function (i) {
    return i.ex + ' 上次 ' + i.lastDate + '：' + i.last.map(function (s) { return s.kg + '×' + s.reps + '@' + s.rir; }).join(', ') + ' → ' + (i.suggestion && i.suggestion.text);
  }), null, 1));
}


// ─────────────── 新增動作（2026/10/19） ───────────────

const EX_TYPES = ['重量型', '自重型', '計時型', '有氧-Zone2', '有氧-HIIT', '有氧-耐力'];

/** 動作清單 O 欄的肌群選項（下拉選單來源） */
function readMuscles_() {
  const sh = ss_().getSheetByName(SHEET_EX);
  return sh.getRange(5, 15, 20, 1).getValues()
    .map(function (r) { return String(r[0] || '').trim(); })
    .filter(function (s) { return s && s.indexOf('（') !== 0; });
}

/** 寫進動作清單第一個空白名稱列；同名（不分大小寫）視為已存在，回傳成功 */
function addExercise_(b) {
  const name = String(b.name || '').trim().replace(/\s+/g, ' ');
  if (!name || name.length > 60) throw new ApiError('bad_name', '動作名稱不可空白或超過 60 字');
  const type = String(b.type || '重量型').trim();
  if (EX_TYPES.indexOf(type) < 0) throw new ApiError('bad_type', '動作類型不正確：' + type);
  const muscles = readMuscles_();
  const pick = function (m) {
    m = String(m || '').trim();
    if (m && muscles.indexOf(m) < 0) throw new ApiError('bad_muscle', '肌群不在選項內：' + m);
    return m;
  };
  const cardio = type.indexOf('有氧') === 0;
  const primary = cardio ? '' : pick(b.primary);
  const syn = cardio ? [] : (b.synergists || []).slice(0, 3).map(pick).filter(Boolean);

  const sh = ss_().getSheetByName(SHEET_EX);
  const v = sh.getRange(5, 2, 50, 1).getValues();
  for (let i = 0; i < v.length; i++) {
    const n = String(v[i][0] || '').trim();
    if (n && n.toLowerCase() === name.toLowerCase()) return { ok: true, duplicate: true, name: n };
  }
  let row = 0;
  for (let i = 0; i < v.length; i++) if (String(v[i][0] || '').trim() === '') { row = 5 + i; break; }
  if (!row) throw new ApiError('list_full', '動作清單已滿 50 個');
  sh.getRange(row, 2, 1, 6).setValues([[name, type, primary, syn[0] || '', syn[1] || '', syn[2] || '']]);
  if (b.note) sh.getRange(row, 8).setValue(String(b.note));
  return { ok: true, duplicate: false, name: name, type: type, row: row };
}

/** 用 API 或外部工具直接改了「訓練紀錄」後，在編輯器手動執行這個清掉快取 */
function clearCache() { cacheClear_(); Logger.log('快取已清除'); }
