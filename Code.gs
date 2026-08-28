/**********************************************************************
 * 業務進度板 —— Google 試算表後端
 *
 * 用途：讓多台電腦共同編輯同一份紀錄，並提供家長專屬的唯讀檢視。
 * 安裝：見「說明.html」的圖文教學。簡單講就是
 *        試算表 → 擴充功能 → Apps Script → 貼上這整份 → 部署為網頁應用程式
 *
 * 資料存放：JSON 檔存在這份試算表旁邊的 Drive 資料夾裡，
 *          試算表本身只放一份「看得懂的檢視表」（可隨時手動更新）。
 *
 * 這份程式不會把資料傳給任何第三方，全部留在老師自己的 Google 帳號裡。
 **********************************************************************/

/* ---------------------------------------------------------------------
   選用：連線密碼
   留空 = 只要知道網址的人都能讀寫（網址很長很難猜，但畢竟是公開的）。
   想多一層保護的話，在這裡填一段自訂密碼，例如 'ming2026'，
   然後在程式的「共用連線」欄位貼成：
       https://script.google.com/macros/s/XXXX/exec?k=ming2026
   家長頁不受影響，家長仍然只用自己的六碼代碼。
   --------------------------------------------------------------------- */
var ACCESS_KEY = '';

var PROP_FILE = 'BOARD_FILE_ID';
var PROP_REV  = 'BOARD_REV';
var DATA_NAME = '業務進度板_資料.json';
var PREV_NAME = '業務進度板_資料_前一版.json';
var PROP_PREV = 'BOARD_PREV_ID';
var TZ        = 'Asia/Taipei';
var LOCK_MS   = 25000;

/* ==================================================================
   對外入口
   ================================================================== */

function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents);
    out = route_(body, (e && e.parameter) || {});
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return json_(out);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action) return json_(route_(p, p));
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:28px;line-height:1.8">' +
    '<h2 style="margin:0 0 10px">業務進度板 —— 後端運作中 ✓</h2>' +
    '<p>部署成功了。請把<b>目前這一頁的網址</b>複製起來，' +
    '貼進程式「設定 → 肆、共用連線」的欄位裡。</p>' +
    '<p style="color:#777;font-size:.9em">網址結尾應該是 <code>/exec</code>。' +
    '這一頁沒有任何資料，可以放心關掉。</p></div>'
  );
}

function route_(body, query) {
  switch (body.action) {
    case 'ping':
      return { ok: true, ver: 1, time: stamp_(), locked: !!ACCESS_KEY };
    case 'get':
      if (!authed_(query, body)) return { ok: false, error: 'auth' };
      return read_();
    case 'put':
      if (!authed_(query, body)) return { ok: false, error: 'auth' };
      return write_(body);
    case 'parent':
      return parentView_(body);          /* 家長端用自己的六碼代碼，不需要連線密碼 */
    default:
      return { ok: false, error: 'unknown action' };
  }
}

function authed_(query, body) {
  if (!ACCESS_KEY) return true;
  var k = (query && query.k) || (body && body.key) || '';
  return String(k) === String(ACCESS_KEY);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function stamp_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}
function todayStr_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

/* ==================================================================
   讀寫
   ================================================================== */

function fileByProp_(propKey, name, createIfMissing) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(propKey);
  if (id) {
    try { return DriveApp.getFileById(id); } catch (err) { /* 檔案被刪掉了 */ }
  }
  if (!createIfMissing) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var f = folder.createFile(name, '{}', 'application/json');
  props.setProperty(propKey, f.getId());
  return f;
}

function dataFile_(createIfMissing) {
  return fileByProp_(PROP_FILE, DATA_NAME, createIfMissing);
}

/* 覆蓋前的上一版。萬一有人在某台電腦上誤按「載入示範資料」之類的動作，
   把整個班級洗掉又同步上來，可以從試算表選單整份救回去。 */
function prevFile_(createIfMissing) {
  return fileByProp_(PROP_PREV, PREV_NAME, createIfMissing);
}

function read_() {
  var props = PropertiesService.getScriptProperties();
  var rev = Number(props.getProperty(PROP_REV) || 0);
  var f = dataFile_(false);
  if (!f) return { ok: true, rev: 0, data: null };
  try {
    var txt = f.getBlob().getDataAsString('UTF-8');
    var data = JSON.parse(txt);
    if (!data || !data.students) return { ok: true, rev: rev, data: null };
    return { ok: true, rev: rev, data: data };
  } catch (err) {
    return { ok: true, rev: rev, data: null };
  }
}

function write_(body) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_MS)) return { ok: false, error: 'busy' };
  try {
    var props = PropertiesService.getScriptProperties();
    var rev = Number(props.getProperty(PROP_REV) || 0);
    var sent = Number(body.rev || 0);

    /* 版本不一致 —— 有人搶先存過了，把雲端目前的內容送回去讓程式自己合併 */
    if (sent !== rev) {
      var cur = read_();
      return { ok: false, conflict: true, rev: cur.rev, data: cur.data };
    }
    if (!body.data || !body.data.students) return { ok: false, error: 'empty data' };

    /* 先把即將被覆蓋的內容抄一份到「前一版」，再寫入新的 */
    try {
      var old = dataFile_(false);
      if (old) {
        prevFile_(true).setContent(JSON.stringify({
          rev: rev,
          savedAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
          data: JSON.parse(old.getBlob().getDataAsString('UTF-8'))
        }));
      }
    } catch (errPrev) { /* 留存失敗不影響正常寫入 */ }

    dataFile_(true).setContent(JSON.stringify(body.data));
    rev = rev + 1;
    props.setProperty(PROP_REV, String(rev));
    noteStatus_(body.data, rev);
    return { ok: true, rev: rev };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/* 在試算表第一張表寫幾行狀態，讓老師打開試算表就知道有沒有在跑 */
function noteStatus_(data, rev) {
  try {
    var sh = sheet_('連線狀態');
    sh.clear();
    var students = (data.students && data.students.list) || [];
    var days = Object.keys(data.days || {}).length;
    sh.getRange(1, 1, 6, 2).setValues([
      ['業務進度板　連線狀態', ''],
      ['最後更新', stamp_()],
      ['版本號', rev],
      ['名冊人數', students.length],
      ['已建檔天數', days],
      ['說明', '資料本體在同資料夾的 ' + DATA_NAME + '。要看得懂的表格請用上方「業務進度板」選單 →「更新檢視表」。']
    ]);
    sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
    sh.getRange(1, 1, 6, 1).setFontWeight('bold');
    sh.setColumnWidth(1, 130);
    sh.setColumnWidth(2, 560);
  } catch (err) { /* 寫狀態失敗不影響存檔 */ }
}

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/* ==================================================================
   家長檢視 —— 只回傳該名學生的資料，其他同學的姓名與進度不會離開伺服器
   ================================================================== */

function parentView_(body) {
  var r = read_();
  var data = r.data;
  if (!data) return { ok: false, reason: 'nodata' };
  if (!data.prefs || !data.prefs.parent) return { ok: false, reason: 'closed' };

  var code = String(body.code || '').trim().toUpperCase();
  if (!code) return { ok: false, reason: 'bad' };

  var list = (data.students && data.students.list) || [];
  var me = null;
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].code || '').toUpperCase() === code) { me = list[i]; break; }
  }
  if (!me) return { ok: false, reason: 'bad' };

  var days = data.days || {};
  var date = pickDate_(days, body.date);
  var rec = days[date] || { notes: [], cases: [], done: {}, watch: [] };

  var cases = (rec.cases || []).map(function (k) {
    return { text: k.text, done: !!(rec.done || {})[me.id + '|' + k.id] };
  });
  var notes = (rec.notes || []).map(function (n) {
    return { text: n.text, flag: !!n.flag };
  });

  /* 這孩子的補件清單（若有被列入催件名單） */
  var owing = [];
  var fl = (data.flags || {})[me.id];
  if (fl && !fl.gone) {
    owing = (fl.items || []).filter(function (it) { return !it.done; })
                            .map(function (it) { return it.text; });
  }
  var watched = (rec.watch || []).indexOf(me.id) >= 0;

  /* 最近十個有派工的上班日，給家長看趨勢 */
  var hist = Object.keys(days).filter(function (d) {
    return d <= date && (days[d].cases || []).length;
  }).sort().slice(-10).map(function (d) {
    var rr = days[d], total = (rr.cases || []).length, done = 0;
    (rr.cases || []).forEach(function (k) {
      if ((rr.done || {})[me.id + '|' + k.id]) done++;
    });
    return { date: d, total: total, done: done };
  });

  return {
    ok: true,
    org: { name: (data.org && data.org.name) || '', sub: (data.org && data.org.sub) || '',
           teacher: (data.org && data.org.teacher) || '' },
    lex: (data.prefs && data.prefs.lex) || 'corp',
    date: date,
    today: todayStr_(),
    dates: Object.keys(days).filter(function (d) {
      return (days[d].notes || []).length || (days[d].cases || []).length;
    }).sort(),
    student: { seat: me.seat || '', name: me.name || '' },
    notes: notes,
    cases: cases,
    owing: owing,
    watched: watched,
    history: hist
  };
}

/* 預設顯示今天。今天還沒發布的話，挑「離今天最近」的一天：
   通常就是昨天，但若老師已先把明天的內容準備好，就顯示明天。
   距離一樣時以過去那天為準。 */
function pickDate_(days, want) {
  if (want && days[want]) return want;
  var t = todayStr_();
  var has = function (d) {
    return days[d] && ((days[d].notes || []).length || (days[d].cases || []).length);
  };
  if (has(t)) return t;
  var ms = function (d) { return new Date(d.replace(/-/g, '/')).getTime(); };
  var base = ms(t);
  var best = null, bestGap = Infinity;
  Object.keys(days).forEach(function (d) {
    if (!has(d)) return;
    var gap = Math.abs(ms(d) - base);
    if (gap < bestGap || (gap === bestGap && ms(d) < base)) { best = d; bestGap = gap; }
  });
  return best || t;
}

/* ==================================================================
   給老師看的檢視表（手動更新，不影響日常存檔速度）
   ================================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('業務進度板')
    .addItem('更新檢視表', 'rebuildView')
    .addSeparator()
    .addItem('檢查安裝狀態', 'checkSetup')
    .addItem('還原雲端前一版', 'restorePrev')
    .addToUi();
}

function checkSetup() {
  var r = read_();
  var ui = SpreadsheetApp.getUi();
  if (!r.data) {
    ui.alert('尚未收到任何資料。\n\n請確認：\n' +
             '1. 已經部署為「網頁應用程式」，且存取權設為「任何人」\n' +
             '2. 已把 /exec 結尾的網址貼進程式的「共用連線」欄位\n' +
             '3. 在程式裡改一筆資料，讓它上傳一次');
    return;
  }
  var n = (r.data.students.list || []).length;
  ui.alert('運作正常 ✓\n\n版本號：' + r.rev +
           '\n名冊人數：' + n +
           '\n已建檔天數：' + Object.keys(r.data.days || {}).length +
           '\n家長檢視：' + (r.data.prefs && r.data.prefs.parent ? '已啟用' : '關閉') +
           '\n最後更新：' + stamp_());
}

/* 從選單執行：把雲端資料整份退回上一版。
   例如某台電腦誤按了會清空資料的動作、又同步上來的時候。 */
function restorePrev() {
  var ui = SpreadsheetApp.getUi();
  var pf = prevFile_(false);
  if (!pf) { ui.alert('目前沒有可以還原的前一版。'); return; }

  var snap;
  try { snap = JSON.parse(pf.getBlob().getDataAsString('UTF-8')); }
  catch (err) { ui.alert('前一版的檔案讀不出來。'); return; }
  if (!snap || !snap.data || !snap.data.students) {
    ui.alert('目前沒有可以還原的前一版。'); return;
  }

  var cur = read_();
  var nOld = (snap.data.students.list || []).length;
  var dOld = Object.keys(snap.data.days || {}).length;
  var nNow = cur.data ? (cur.data.students.list || []).length : 0;
  var dNow = cur.data ? Object.keys(cur.data.days || {}).length : 0;

  var ans = ui.alert(
    '還原雲端前一版',
    '目前雲端：' + nNow + ' 人、' + dNow + ' 個工作日\n' +
    '前一版：　' + nOld + ' 人、' + dOld + ' 個工作日' +
    '（存於 ' + (snap.savedAt || '未知時間') + '）\n\n' +
    '要把雲端還原成前一版嗎？\n' +
    '還原後，各台電腦下次同步就會拿到這一版。',
    ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_MS)) { ui.alert('系統忙碌中，請稍後再試一次。'); return; }
  try {
    var props = PropertiesService.getScriptProperties();
    var rev = Number(props.getProperty(PROP_REV) || 0);
    /* 把「現在這一版」換到前一版的位置，這樣還原本身也可以再還原一次 */
    if (cur.data) {
      pf.setContent(JSON.stringify({
        rev: rev,
        savedAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
        data: cur.data
      }));
    }
    dataFile_(true).setContent(JSON.stringify(snap.data));
    rev = rev + 1;
    props.setProperty(PROP_REV, String(rev));
    noteStatus_(snap.data, rev);
  } finally {
    lock.releaseLock();
  }
  ui.alert('已還原　' + nOld + ' 人、' + dOld + ' 個工作日。\n\n' +
           '請到各台電腦按一下右上角的狀態標記（或重新整理）取得新版本。');
}

function rebuildView() {
  var r = read_();
  if (!r.data) {
    SpreadsheetApp.getUi().alert('目前還沒有資料可以整理。');
    return;
  }
  var data = r.data;
  buildRoster_(data);
  buildDaily_(data);
  buildOpen_(data);
  SpreadsheetApp.getActiveSpreadsheet().toast('檢視表已更新（' + stamp_() + '）', '業務進度板', 5);
}

function teamName_(data, t) {
  if (!t) return '未編組';
  var m = (data.teams && data.teams.map) || {};
  return m[t] || ('第' + t + '組');
}

function buildRoster_(data) {
  var sh = sheet_('名冊');
  sh.clear();
  var rows = [['座號', '姓名', '組別', '家長代碼']];
  (data.students.list || []).slice()
    .sort(function (a, b) { return (a.seat || 0) - (b.seat || 0); })
    .forEach(function (s) {
      rows.push([s.seat || '', s.name || '', teamName_(data, s.team), s.code || '']);
    });
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e6d9bc');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 4);
}

function buildDaily_(data) {
  var sh = sheet_('每日紀錄');
  sh.clear();
  var students = data.students.list || [];
  var rows = [['日期', '星期', '新案件數', '待簽核件數', '已核准格', '總格數', '完成率', '催件人數']];
  var W = ['日', '一', '二', '三', '四', '五', '六'];
  Object.keys(data.days || {}).sort().forEach(function (ds) {
    var d = data.days[ds];
    var cases = (d.cases || []).length;
    var total = cases * students.length, done = 0;
    students.forEach(function (s) {
      (d.cases || []).forEach(function (k) {
        if ((d.done || {})[s.id + '|' + k.id]) done++;
      });
    });
    var w = W[new Date(ds.replace(/-/g, '/')).getDay()];
    rows.push([ds, w, (d.notes || []).length, cases, done, total,
               total ? Math.round(done / total * 100) / 100 : '',
               (d.watch || []).length]);
  });
  sh.getRange(1, 1, rows.length, 8).setValues(rows);
  sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#e6d9bc');
  if (rows.length > 1) sh.getRange(2, 7, rows.length - 1, 1).setNumberFormat('0%');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 8);
}

function buildOpen_(data) {
  var sh = sheet_('未結明細');
  sh.clear();
  var students = data.students.list || [];
  var byId = {};
  students.forEach(function (s) { byId[s.id] = s; });

  var rows = [['日期', '座號', '姓名', '組別', '未完成項目']];
  var keys = Object.keys(data.days || {}).sort();
  keys = keys.slice(-40);                       /* 只列最近 40 個建檔日，免得表太長 */
  keys.forEach(function (ds) {
    var d = data.days[ds];
    (d.cases || []).forEach(function (k) {
      students.forEach(function (s) {
        if (!(d.done || {})[s.id + '|' + k.id]) {
          rows.push([ds, s.seat || '', s.name || '', teamName_(data, s.team), k.text]);
        }
      });
    });
  });

  /* 催件名單另外附在後面 */
  var flags = data.flags || {};
  var extra = [];
  Object.keys(flags).forEach(function (sid) {
    var f = flags[sid], s = byId[sid];
    if (!f || f.gone || !s) return;
    (f.items || []).forEach(function (it) {
      if (!it.done) extra.push(['（催件）', s.seat || '', s.name || '',
                                teamName_(data, s.team), it.text + '　自 ' + f.since + ' 起']);
    });
  });
  if (extra.length) { rows.push(['', '', '', '', '']); rows = rows.concat(extra); }

  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3d5d1');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 5);
}
