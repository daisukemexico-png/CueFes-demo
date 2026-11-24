/**
 * CueFes ⇄ Google スプレッドシート連携スクリプト
 *
 * 使い方（約5分）
 *  1. 新しいGoogleスプレッドシートを作る
 *  2. メニュー「拡張機能」→「Apps Script」を開き、このコードを全部貼り付けて保存
 *  3. 下の SHARED_KEY（合言葉）を好きな文字列に変える
 *  4. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       次のユーザーとして実行：自分
 *       アクセスできるユーザー：全員
 *     →「デプロイ」→ 表示された「ウェブアプリのURL」をコピー
 *  5. CueFes の「スプレッドシート共有」画面に URL と合言葉を入れて「接続テスト」
 *
 * 作られるシート
 *  - 概要            : イベント名・最終更新者・更新日時
 *  - 出演者一覧      : 1行1組。ここを直接編集 →「シートから読み込む」でアプリに反映できます
 *  - ライブ DAY1 など : 時間×ステージの見やすいタイテ（書き出すたびに作り直し。編集しても反映されません）
 *  - 特典会 DAY1 など : 同上
 *  - _cuefes_data    : アプリ用のデータ（非表示。触らないでください）
 */

// ▼ 合言葉。CueFes側にも同じものを入れます。空にすると誰でも読み書きできてしまうので必ず設定してください。
const SHARED_KEY = 'ここを合言葉に変える';

const DATA_SHEET = '_cuefes_data';
const ARTIST_SHEET = '出演者一覧';
const SUMMARY_SHEET = '概要';
const LIVE_PREFIX = 'ライブ ';
const TOKUTEN_PREFIX = '特典会 ';
const CHUNK = 40000; // 1セル50,000文字制限より少し小さく
const SLOT_MIN = 5;  // タイテシートの1行あたりの分数

const ARTIST_HEADERS = ['ID', 'DAY', 'アーティスト名', 'ステージ', 'ライブ開始', 'ライブ終了',
  '特典会場', '特典会開始', '特典会終了', '楽屋', '入り時間', 'メモ'];

function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_json' }); }
  return handle_(body);
}

function handle_(p) {
  if (!SHARED_KEY || SHARED_KEY === 'ここを合言葉に変える') return out_({ ok: false, error: 'key_not_set' });
  if (String(p.key || '') !== SHARED_KEY) return out_({ ok: false, error: 'bad_key' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (p.action) {
      case 'ping': return out_(Object.assign({ ok: true }, meta_()));
      case 'pull': return out_(pull_());
      case 'push': return out_(push_(p));
      default: return out_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return out_({ ok: false, error: 'server_error', message: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function props_() { return PropertiesService.getScriptProperties(); }
function meta_() {
  const pr = props_();
  return {
    rev: Number(pr.getProperty('rev') || 0),
    updatedAt: pr.getProperty('updatedAt') || '',
    updatedBy: pr.getProperty('updatedBy') || '',
    sheetEdited: artistSheetHash_() !== (pr.getProperty('artistHash') || ''),
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl()
  };
}

// ---------- pull ----------
function pull_() {
  const m = meta_();
  const state = readData_();
  return Object.assign({ ok: true, state: state, rows: readArtistRows_() }, m);
}

// ---------- push ----------
function push_(p) {
  const m = meta_();
  if (!p.force) {
    // まだ一度も読み込んでいない端末（baseRevなし）は、シートが空のときだけ書き出せる
    const stale = (p.baseRev == null || p.baseRev === '') ? m.rev > 0 : Number(p.baseRev) !== m.rev;
    if (stale) return Object.assign({ ok: false, conflict: 'newer_rev' }, m);
    if (m.rev > 0 && m.sheetEdited) return Object.assign({ ok: false, conflict: 'sheet_edited' }, m);
  }
  const state = p.state;
  if (!state || !state.days || !state.artists) return { ok: false, error: 'bad_state' };
  writeData_(state);
  writeSummary_(state, p.by);
  writeArtistSheet_(state);
  writeTimetables_(state);
  const rev = m.rev + 1;
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const pr = props_();
  pr.setProperties({ rev: String(rev), updatedAt: now, updatedBy: String(p.by || ''), artistHash: artistSheetHash_() });
  return Object.assign({ ok: true }, meta_());
}

// ---------- データ（非表示シート） ----------
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function writeData_(state) {
  const sh = sheet_(DATA_SHEET);
  sh.clear();
  const json = JSON.stringify(state);
  const rows = [];
  for (let i = 0; i < json.length; i += CHUNK) rows.push([json.slice(i, i + CHUNK)]);
  if (rows.length) sh.getRange(1, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
  if (!sh.isSheetHidden()) sh.hideSheet();
}
function readData_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  if (!sh || sh.getLastRow() < 1) return null;
  const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  try { return JSON.parse(vals.map(function (r) { return r[0]; }).join('')); } catch (e) { return null; }
}

// ---------- 概要 ----------
function writeSummary_(state, by) {
  const sh = sheet_(SUMMARY_SHEET);
  sh.clear();
  const ev = state.event || {};
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const rows = [
    ['CueFes 共有シート', ''],
    ['イベント名', ev.name || ''],
    ['会場', ev.venue || ''],
    ['開催日', (state.days || []).map(function (d) { return d.label + (d.date ? '（' + d.date + '）' : ''); }).join(' / ')],
    ['出演組数', String(uniqueNames_(state.artists).length)],
    ['最終更新', now + (by ? '（' + by + '）' : '')],
    ['', ''],
    ['このシートの使い方', '「出演者一覧」は編集OK → CueFesの「シートから読み込む」で反映。「ライブ」「特典会」のタイテは書き出すたびに作り直されます。']
  ];
  sh.getRange(1, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sh.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold').setFontColor('#5E5C57');
  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 520);
  sh.getRange(8, 2).setWrap(true);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(1);
}
function uniqueNames_(artists) {
  const seen = {};
  (artists || []).forEach(function (a) { if (a.name) seen[a.name] = 1; });
  return Object.keys(seen);
}

// ---------- 出演者一覧（編集可能） ----------
function nameOf_(list, id) { const x = (list || []).filter(function (o) { return o.id === id; })[0]; return x ? x.name : ''; }
function writeArtistSheet_(state) {
  const sh = sheet_(ARTIST_SHEET);
  sh.clear();
  const dayOrder = {};
  (state.days || []).forEach(function (d, i) { dayOrder[d.id] = i; });
  const artists = (state.artists || []).slice().sort(function (a, b) {
    const da = dayOrder[a.dayId] == null ? 99 : dayOrder[a.dayId];
    const db = dayOrder[b.dayId] == null ? 99 : dayOrder[b.dayId];
    if (da !== db) return da - db;
    return toMin_(a.liveStart || a.tokutenStart) - toMin_(b.liveStart || b.tokutenStart);
  });
  const rows = artists.map(function (a) {
    return [a.id, nameOf_(state.days.map(function (d) { return { id: d.id, name: d.label }; }), a.dayId), a.name || '',
      nameOf_(state.stages, a.stageId), a.liveStart || '', a.liveEnd || '',
      nameOf_(state.tokutenSpots, a.tokutenSpotId), a.tokutenStart || '', a.tokutenEnd || '',
      nameOf_(state.greenRooms, a.greenRoomId), a.callTime || '', a.memo || ''];
  });
  const all = [ARTIST_HEADERS].concat(rows);
  const range = sh.getRange(1, 1, all.length, ARTIST_HEADERS.length);
  range.setNumberFormat('@').setValues(all);
  sh.getRange(1, 1, 1, ARTIST_HEADERS.length).setFontWeight('bold').setBackground('#1F1F1F').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, Math.max(all.length, 2), 1).setFontColor('#9A978F');
  sh.getRange(1, 1).setNote('IDは変更しないでください。新しい出演者はIDを空欄にして行を追加すると、読み込み時に追加されます。');
  sh.setColumnWidth(1, 90); sh.setColumnWidth(3, 200); sh.setColumnWidth(12, 260);
  // 選択肢（入力ミス防止）
  const listRule = function (names) {
    return names.length ? SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(true).build() : null;
  };
  const n = Math.max(rows.length + 50, 60);
  const setRule = function (col, names) { const r = listRule(names); if (r) sh.getRange(2, col, n, 1).setDataValidation(r); };
  setRule(2, (state.days || []).map(function (d) { return d.label; }));
  setRule(4, (state.stages || []).map(function (s) { return s.name; }));
  setRule(7, (state.tokutenSpots || []).map(function (s) { return s.name; }));
  setRule(10, (state.greenRooms || []).map(function (s) { return s.name; }));
  sh.getRange(2, 5, n, 2).setNumberFormat('@');
  sh.getRange(2, 8, n, 2).setNumberFormat('@');
  sh.getRange(2, 11, n, 1).setNumberFormat('@');
}
function readArtistRows_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ARTIST_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(1, 1, sh.getLastRow(), ARTIST_HEADERS.length).getDisplayValues();
  const head = vals[0];
  return vals.slice(1).filter(function (r) { return r.join('').trim() !== ''; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = String(r[i] == null ? '' : r[i]).trim(); });
    return o;
  });
}
function artistSheetHash_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ARTIST_SHEET);
  if (!sh || sh.getLastRow() < 1) return '';
  const vals = sh.getRange(1, 1, sh.getLastRow(), ARTIST_HEADERS.length).getDisplayValues();
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(vals), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

// ---------- タイテ（時間×ステージの見やすい表） ----------
function toMin_(t) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(t || ''));
  return m ? Number(m[1]) * 60 + Number(m[2]) : 99999;
}
function fmt_(min) {
  const h = Math.floor(min / 60), m = min % 60;
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}
function writeTimetables_(state) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wanted = {};
  (state.days || []).forEach(function (day) {
    const live = buildGrid_(state, day, 'live');
    const tok = buildGrid_(state, day, 'tokuten');
    if (live) { renderGrid_(LIVE_PREFIX + day.label, live, day, 'ライブ'); wanted[LIVE_PREFIX + day.label] = 1; }
    if (tok) { renderGrid_(TOKUTEN_PREFIX + day.label, tok, day, '特典会'); wanted[TOKUTEN_PREFIX + day.label] = 1; }
  });
  ss.getSheets().forEach(function (sh) {
    const n = sh.getName();
    if ((n.indexOf(LIVE_PREFIX) === 0 || n.indexOf(TOKUTEN_PREFIX) === 0) && !wanted[n] && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}
function buildGrid_(state, day, kind) {
  const laneKey = kind === 'live' ? 'stageId' : 'tokutenSpotId';
  const sKey = kind === 'live' ? 'liveStart' : 'tokutenStart';
  const eKey = kind === 'live' ? 'liveEnd' : 'tokutenEnd';
  const items = (state.artists || []).filter(function (a) {
    return a.dayId === day.id && a[laneKey] && toMin_(a[sKey]) < 99999 && toMin_(a[eKey]) < 99999;
  });
  // その日に非表示のステージ/会場は除外（ただし出演者が割り当てられていれば表示）
  const hidden = (kind === 'live' ? day.hiddenStageIds : day.hiddenTokutenIds) || [];
  const lanes = ((kind === 'live' ? state.stages : state.tokutenSpots) || []).filter(function (l) {
    return hidden.indexOf(l.id) < 0 || items.some(function (a) { return a[laneKey] === l.id; });
  });
  if (!items.length || !lanes.length) return null;
  let start = Math.min.apply(null, items.map(function (a) { return toMin_(a[sKey]); }));
  let end = Math.max.apply(null, items.map(function (a) { return toMin_(a[eKey]); }));
  start = Math.floor(start / 30) * 30; end = Math.ceil(end / 30) * 30;
  return { lanes: lanes, items: items, start: start, end: end, laneKey: laneKey, sKey: sKey, eKey: eKey };
}
function renderGrid_(name, g, day, kindLabel) {
  const sh = sheet_(name);
  sh.clear();
  const maxR = sh.getMaxRows(), maxC = sh.getMaxColumns();
  sh.getRange(1, 1, maxR, maxC).breakApart();
  const slots = (g.end - g.start) / SLOT_MIN;
  const cols = g.lanes.length + 1;
  const rowsNeeded = slots + 2;
  if (maxR < rowsNeeded) sh.insertRowsAfter(maxR, rowsNeeded - maxR);
  if (maxC < cols) sh.insertColumnsAfter(maxC, cols - maxC);

  const values = [], bgs = [];
  values.push([day.label + (day.date ? ' ' + day.date : '') + ' ' + kindLabel + 'タイムテーブル'].concat(new Array(cols - 1).fill('')));
  bgs.push(new Array(cols).fill('#FFFFFF'));
  values.push(['時間'].concat(g.lanes.map(function (l) { return l.name; })));
  bgs.push(new Array(cols).fill('#1F1F1F'));
  for (let i = 0; i < slots; i++) {
    const t = g.start + i * SLOT_MIN;
    values.push([t % 30 === 0 ? fmt_(t) : ''].concat(new Array(cols - 1).fill('')));
    bgs.push(['#F6F5F2'].concat(new Array(cols - 1).fill(t % 60 === 0 ? '#FAFAF8' : '#FFFFFF')));
  }
  const merges = [];
  g.items.forEach(function (a) {
    const c = g.lanes.map(function (l) { return l.id; }).indexOf(a[g.laneKey]);
    if (c < 0) return;
    const r0 = Math.max(0, Math.floor((toMin_(a[g.sKey]) - g.start) / SLOT_MIN));
    const r1 = Math.min(slots, Math.ceil((toMin_(a[g.eKey]) - g.start) / SLOT_MIN));
    if (r1 <= r0) return;
    values[r0 + 2][c + 1] = a.name + '\n' + a[g.sKey] + '–' + a[g.eKey];
    for (let r = r0; r < r1; r++) bgs[r + 2][c + 1] = /^#[0-9a-fA-F]{6}$/.test(a.color || '') ? a.color : '#FFD3DC';
    merges.push([r0 + 3, c + 2, r1 - r0]);
  });
  const all = sh.getRange(1, 1, values.length, cols);
  all.setNumberFormat('@').setValues(values).setBackgrounds(bgs)
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setFontSize(10).setWrap(true);
  sh.getRange(1, 1, 1, cols).merge().setFontSize(13).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange(2, 1, 1, cols).setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(3, 1, slots, 1).setFontColor('#5E5C57').setVerticalAlignment('top');
  merges.forEach(function (m) {
    const rg = sh.getRange(m[0], m[1], m[2], 1);
    if (m[2] > 1) rg.merge();
    rg.setFontWeight('bold').setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sh.setFrozenRows(2); sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 56);
  for (let c = 2; c <= cols; c++) sh.setColumnWidth(c, 150);
  sh.setRowHeightsForced(3, slots, 16);
  sh.setRowHeight(1, 30); sh.setRowHeight(2, 26);
}
