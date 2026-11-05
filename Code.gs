/**
 * 海上釣堀 釣果ノート 同期用スクリプト（Google スプレッドシートに貼り付けて使います）
 * 1. 下の KEY を好きな合言葉に変える（アプリの設定にも同じ合言葉を入れます）
 * 2. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    実行するユーザー：自分　／　アクセスできるユーザー：全員
 * 3. 表示された「ウェブアプリのURL」（…/exec）をアプリの設定に入れる
 */
const KEY = 'ここに合言葉';

const DATA_FILE = 'tsuribori-data.json';   // 全データ（スプレッドシートと同じフォルダに作られます）
const PHOTO_FOLDER = 'tsuribori-photos';   // 写真の保存フォルダ
const SHEET_NAME = '釣果一覧';              // 見る用の一覧シート

function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(p);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  if (!p.key || p.key !== KEY) return out_({ ok: false, error: '合言葉が違います' });
  var lock = LockService.getScriptLock();
  if (p.key === KEY && p.action === 'tide') {  // 潮汐は保存しないのでロック不要
    try { return out_({ ok: true, tide: tide_(p.stn, p.date) }); } catch (err) { return out_({ ok: false, error: String(err) }); }
  }
  lock.waitLock(25000);
  try {
    switch (p.action) {
      case 'ping':     return out_({ ok: true });
      case 'load':     return out_({ ok: true, data: readData_(), photos: listPhotos_() });
      case 'save':     writeData_(p.data); writeSheet_(p.data); return out_({ ok: true });
      case 'putPhoto': putPhoto_(p.id, p.data); return out_({ ok: true });
      case 'photo':    return out_({ ok: true, data: getPhoto_(p.id) });
      case 'delPhoto': delPhoto_(p.id); return out_({ ok: true });
      case 'tide':     return out_({ ok: true, tide: tide_(p.stn, p.date) });
    }
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 保存場所 ---------- */
function folder_() {
  var ss = SpreadsheetApp.getActive();
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}
function photoFolder_() {
  var f = folder_(), it = f.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : f.createFolder(PHOTO_FOLDER);
}

/* ---------- データ ---------- */
function readData_() {
  var it = folder_().getFilesByName(DATA_FILE);
  if (!it.hasNext()) return null;
  return JSON.parse(it.next().getBlob().getDataAsString('UTF-8'));
}
function writeData_(d) {
  var json = JSON.stringify(d), f = folder_(), it = f.getFilesByName(DATA_FILE);
  if (it.hasNext()) it.next().setContent(json);
  else f.createFile(DATA_FILE, json, MimeType.PLAIN_TEXT);
}

/* ---------- 写真 ---------- */
function photoFile_(id) {
  var it = photoFolder_().getFilesByName(id + '.jpg');
  return it.hasNext() ? it.next() : null;
}
function listPhotos_() {
  var ids = [], it = photoFolder_().getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (!f.isTrashed()) ids.push(f.getName().replace(/\.jpg$/, ''));
  }
  return ids;
}
function putPhoto_(id, dataUrl) {
  if (!id || !dataUrl || photoFile_(id)) return;
  var b64 = String(dataUrl).split(',')[1];
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', id + '.jpg');
  photoFolder_().createFile(blob);
}
function getPhoto_(id) {
  var f = photoFile_(id);
  return f ? 'data:image/jpeg;base64,' + Utilities.base64Encode(f.getBlob().getBytes()) : null;
}
function delPhoto_(id) {
  var f = photoFile_(id);
  if (f) f.setTrashed(true);
}

/* ---------- 見る用の一覧シート ---------- */
function writeSheet_(d) {
  if (!d || !d.trips) return;
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  var rows = [['日付', '釣り場', '潮', '潮の流れ', '水温', '当たりエサ', '名前', '釣座', '位置メモ', '魚種', '匹数', 'メモ']];
  var me = d.me || '自分';
  var trips = d.trips.slice().sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  trips.forEach(function (t) {
    var flow = t.flowNote || [t.flowSpeed, t.flowDir].filter(String).join('・');
    var members = (t.members && t.members.length) ? t.members.slice() : [me];
    members.sort(function (a, b) { return (a === me ? -1 : 0) - (b === me ? -1 : 0); });
    members.forEach(function (p) {
      var c = (t.counts || {})[p] || {}, pos = (t.seats || {})[p] || (p === me && !(t.rafts || {})[p] ? (t.ikada || '') : ''), raft = (t.rafts || {})[p] || '';
      var seat = !raft ? pos : !pos ? raft : (/^[0-9０-９]/.test(pos) ? raft + pos : raft + '-' + pos);
      var note = (t.pnotes || {})[p] || '';
      var fish = Object.keys(c).filter(function (f) { return +c[f] > 0; });
      var base = [t.date, t.place || '', t.tide || '', flow, t.waterTemp || '', (t.baits || []).join('・'), p, seat, note];
      if (!fish.length) rows.push(base.concat(['', 0, t.memo || '']));
      fish.forEach(function (f) { rows.push(base.concat([f, +c[f], t.memo || ''])); });
    });
  });
  sh.clear();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#dde9ef');
}

/* ---------- 潮汐（気象庁 潮位表テキストデータ） ---------- */
// stn: 地点記号（和歌山=WY、大阪=OS など）　date: 'YYYY-MM-DD'
function tide_(stn, date) {
  if (!/^[A-Z0-9]{2}$/.test(String(stn)) || !/^\d{4}-\d{2}-\d{2}$/.test(String(date))) throw new Error('地点または日付が不正です');
  var y = date.slice(0, 4), txt = tideText_(stn, y);
  var yy = +y.slice(2), mm = +date.slice(5, 7), dd = +date.slice(8, 10);
  var lines = txt.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var L = lines[i];
    if (L.length < 136) continue;
    if (+L.substr(72, 2) !== yy || +L.substr(74, 2) !== mm || +L.substr(76, 2) !== dd) continue;
    var hourly = [];
    for (var h = 0; h < 24; h++) hourly.push(+L.substr(h * 3, 3));
    var pick = function (start) {
      var r = [];
      for (var k = 0; k < 4; k++) {
        var t = L.substr(start + k * 7, 4), cm = L.substr(start + k * 7 + 4, 3);
        if (t.trim() === '9999' || t.trim() === '') continue;
        var hh = +t.slice(0, 2), mi = +t.slice(2, 4);
        r.push({ t: hh + ':' + ('0' + mi).slice(-2), cm: +cm });
      }
      return r;
    };
    return { stn: stn, date: date, high: pick(80), low: pick(108), hourly: hourly };
  }
  throw new Error('その日の潮位表が見つかりません');
}
function tideText_(stn, y) {
  var cache = CacheService.getScriptCache(), ck = 'tide_' + stn + '_' + y;
  var c = cache.get(ck);
  if (c) return c;
  var urls = [
    'https://www.data.jma.go.jp/kaiyou/data/db/tide/suisan/txt/' + y + '/' + stn + '.txt',
    'https://www.data.jma.go.jp/gmd/kaiyou/data/db/tide/suisan/txt/' + y + '/' + stn + '.txt'
  ];
  for (var i = 0; i < urls.length; i++) {
    var res = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var txt = res.getContentText('Shift_JIS');
      try { cache.put(ck, txt, 21600); } catch (e) {}
      return txt;
    }
  }
  throw new Error('気象庁の潮位表を取得できませんでした');
}
