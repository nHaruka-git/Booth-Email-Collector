/** front.gs — Web UI backend (Apps Script) **/

// 設定は ScriptProperties に保存し、起動時に CONFIG に反映。
// main.gs にある収集ロジック: processBOOTHSalesEmails, getOrCreateSheet, clearProcessedFlags, triggerSync_, setupTriggers などを前提。

const PROP = PropertiesService.getScriptProperties();
const CFG_KEYS = {
  SHEET_ID: 'SPREADSHEET_ID',
  MAX_SCAN: 'MAX_SCAN_COUNT',
  TRIGGER:  'TRIGGER_ENABLED',
  SETUP:    'SETUP_DONE',
  FREQ:     'TRIGGER_FREQUENCY_MIN'
};
const LOG_SHEET_NAME = 'BOOTH_LOGS';

/* ---- 基本 ---- */
function getSetupDone() {
  return PROP.getProperty(CFG_KEYS.SETUP) === '1';
}
function ensureConfigLoaded() {
  const sid = PROP.getProperty(CFG_KEYS.SHEET_ID);
  const msc = PROP.getProperty(CFG_KEYS.MAX_SCAN);
  if (sid) CONFIG.SPREADSHEET_ID = sid;
  if (msc) CONFIG.MAX_SCAN_COUNT = Number(msc) || CONFIG.MAX_SCAN_COUNT;
}
function doGet() {
  ensureConfigLoaded();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('BOOTH 売上ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---- 設定API ---- */
function getConfigForUI() {
  ensureConfigLoaded();
  const triggerOn = getTriggerStatus();
  const freq = (typeof getTriggerFrequency_ === 'function') ? getTriggerFrequency_() : 5;
  return {
    SPREADSHEET_ID: CONFIG.SPREADSHEET_ID,
    MAX_SCAN_COUNT: CONFIG.MAX_SCAN_COUNT,
    TRIGGER_ENABLED: triggerOn,
    TRIGGER_FREQUENCY_MIN: freq
  };
}

function saveConfigFromUI(cfg) {
  const sidIn = (typeof cfg?.SPREADSHEET_ID === 'string') ? cfg.SPREADSHEET_ID.trim() : '';
  const maxIn = cfg?.MAX_SCAN_COUNT;

  if (sidIn) {
    PROP.setProperty(CFG_KEYS.SHEET_ID, sidIn);
    CONFIG.SPREADSHEET_ID = sidIn;
  } else if (!PROP.getProperty(CFG_KEYS.SHEET_ID)) {
    throw new Error('SPREADSHEET_ID は空にできません');
  }

  if (maxIn != null) {
    const n = Number(maxIn);
    if (!Number.isFinite(n) || n <= 0) throw new Error('MAX_SCAN_COUNT は正の数');
    PROP.setProperty(CFG_KEYS.MAX_SCAN, String(n));
    CONFIG.MAX_SCAN_COUNT = n;
  }

  return getConfigForUI();
}

/* ---- セットアップ / トリガー ---- */
function runInitialSetupFromUI() {
  // 初期セットアップ後は強制ON（要件どおり）
  return _setupCore_('ui'); // main.gs 側の実装を使用
}

function setTriggerConfigFromUI(payload) {
  ensureConfigLoaded();
  const sp = PropertiesService.getScriptProperties();
  const setupDone = sp.getProperty(CFG_KEYS.SETUP) === '1';

  const minutes = Number(payload && payload.minutes) || 5;
  const enabled = !!(payload && payload.enabled);

  sp.setProperty(CFG_KEYS.TRIGGER, enabled ? '1' : '0');
  sp.setProperty(CFG_KEYS.FREQ, String(minutes));

  // セットアップ前は実体OFF固定（要件）
  const actualEnabled = setupDone ? _applyTrigger_(enabled, minutes) : false;

  addLogFromServer && addLogFromServer('INFO','トリガー更新',
    {desiredEnabled: enabled, minutes, applied: actualEnabled, setupDone});

  return { enabled: actualEnabled, minutes };
}

function _applyTrigger_(enabled, minutes) {
  // 既存削除
  ScriptApp.getProjectTriggers().forEach(t=>{
    if (t.getHandlerFunction && t.getHandlerFunction() === 'processBOOTHSalesEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });
  if (!enabled) {
    addLogFromServer && addLogFromServer('INFO','トリガー未作成',{reason:'disabled'});
    return false;
  }
  let tb = ScriptApp.newTrigger('processBOOTHSalesEmails').timeBased();
  if (minutes === 60) {
    tb = tb.everyHours(1);
  } else if ([1,5,10,15,30].includes(minutes)) {
    tb = tb.everyMinutes(minutes);
  } else {
    minutes = 5; // フォールバック
    tb = tb.everyMinutes(5);
  }
  tb.create();
  addLogFromServer && addLogFromServer('INFO','トリガー作成/更新',{minutesApplied: minutes});
  return _isTriggerOn_();
}

/* ---- ダッシュボード統計 ---- */
function getDashboardStats() {
  try {
    ensureConfigLoaded();
    const id = CONFIG.SPREADSHEET_ID;
    const stats = { collectedRows: 0, scannedThreads: 0, setupDone: getSetupDone() };

    // 記載済み件数
    if (id && id !== 'YOUR_SPREADSHEET_ID_HERE') {
      try {
        const sh = SpreadsheetApp.openById(id).getSheetByName(CONFIG.SHEET_NAME);
        if (sh) stats.collectedRows = Math.max(0, sh.getLastRow() - 1);
      } catch(e) { /* ignore */ }
    }

    // スキャン件数（ページングで全件カウント）
    try {
      const query = 'from:noreply@booth.pm (subject:商品が購入されました OR subject:ご注文が確定しました)';
      let total = 0;
      const pageSize = 500;
      for (let start = 0; ; start += pageSize) {
        const batch = GmailApp.search(query, start, pageSize);
        if (!batch.length) break;
        total += batch.length;
        if (batch.length < pageSize) break;
      }
      stats.scannedThreads = total;
    } catch(e) { /* ignore */ }

    return stats;
  } catch(e) {
    return { collectedRows: null, scannedThreads: null, error: String(e) };
  }
}


/* ---- 収集のワンショット実行 ---- */
function runCollectorOnceFromUI() {
  try {
    ensureConfigLoaded();
    return processBOOTHSalesEmails(); // main.gs
  } catch (e) {
    return { ok:false, error:true, message:String(e) };
  }
}

/* ---- ログAPI ---- */
function openSpreadsheetSafe_() {
  try {
    ensureConfigLoaded();
    const id = CONFIG.SPREADSHEET_ID;
    if (!id || id === 'YOUR_SPREADSHEET_ID_HERE') return null;
    return SpreadsheetApp.openById(id);
  } catch (e) {
    addLogFromServer('WARN','openById失敗',{error:String(e)});
    return null;
  }
}
function ensureLogSheet_() {
  const ss = openSpreadsheetSafe_();
  if (!ss) return null;
  let sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.getRange(1,1,1,4).setValues([['日時','レベル','メッセージ','コンテキスト']]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function addLogFromServer(level, message, context) {
  try {
    const ss = openSpreadsheetSafe_();
    if (!ss) return false;
    let sh = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(LOG_SHEET_NAME);
      sh.getRange(1,1,1,4).setValues([['日時','レベル','メッセージ','コンテキスト']]);
      sh.setFrozenRows(1);
    }
    sh.appendRow([ new Date(), String(level||'INFO'), String(message||''), context ? JSON.stringify(context).slice(0,5000) : '' ]);
    return true;
  } catch(e) { return false; }
}
function getRecentLogs(limit) {
  const tz = Session.getScriptTimeZone();
  const nowStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');
  const sh = ensureLogSheet_();
  if (!sh) {
    return [{ time: nowStr, level: 'INFO', message: 'ログ未取得: SPREADSHEET_ID未設定/アクセス権なし/ID不正', context: '' }];
  }
  const last = sh.getLastRow();
  if (last <= 1) return [];
  const n = Math.min(Math.max(1, Number(limit)||200), 1000);
  const start = Math.max(2, last - n + 1);
  const vals = sh.getRange(start,1,last-start+1,4).getValues();
  return vals.map(([ts,lv,msg,ctx])=>({
    time: Utilities.formatDate(new Date(ts), tz, 'yyyy/MM/dd HH:mm:ss'),
    level: lv, message: msg, context: ctx
  })).reverse();
}
function clearLogs() {
  const sh = ensureLogSheet_();
  if (!sh) return false;
  if (sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,4).clearContent();
  return true;
}

/* ---- 分析データ API ---- */
// シート A:E = [日時, 注文番号, 商品名, 版, 金額]
function _readRows_() {
  ensureConfigLoaded();
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  const vals = sheet.getRange(2,1,last-1,5).getValues();
  return vals.map(r => ({
    date: (r[0] instanceof Date)? r[0] : new Date(r[0]),
    orderNumber: r[1],
    product: String(r[2]||''),
    variant: String(r[3]||''),
    amount: Number(r[4])||0
  })).filter(x=>x.product && x.amount>=0 && x.date instanceof Date && !isNaN(x.date));
}
function _startOfDay(d){ const x=new Date(d); x.setHours(0,0,0,0); return x; }
function _startOfWeek(d){ const x=_startOfDay(d); const dow=(x.getDay()+6)%7; x.setDate(x.getDate()-dow); return x; } // ISO的（月曜始まり）
function _startOfMonth(d){ const x=new Date(d.getFullYear(), d.getMonth(), 1); x.setHours(0,0,0,0); return x; }
function _formatDate(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function _formatWeek(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'YYYY-ww'); }
function _formatMonth(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM'); }
/** 週ラベル: 年-月-週（週番号は年内週） */
function _formatWeekYMW(d){
  const tz = Session.getScriptTimeZone();
  const y = Utilities.formatDate(d, tz, 'yyyy');
  const m = Utilities.formatDate(d, tz, 'MM');
  const w = Utilities.formatDate(d, tz, 'ww'); // 年内週番号
  return `${y}-${m}-${w}`;
}

/** 商品別 件数/金額 折れ線
 * granularity: 'daily' | 'weekly' | 'monthly'
 * lookback: daily=7|30, weekly=4|12, monthly=未使用
 */
function getSeriesByProduct(granularity, lookback) {
  const rows = _readRows_();

  // 下限日
  const now = new Date();
  let from = new Date(2000,0,1);
  if (granularity === 'daily') {
    const days = Number(lookback) || 30;
    const today0 = _startOfDay(now);
    from = new Date(today0);
    from.setDate(from.getDate() - (days - 1));
  } else if (granularity === 'weekly') {
    const w = Number(lookback) || 12;
    const thisWeek = _startOfWeek(now);
    from = new Date(thisWeek);
    from.setDate(from.getDate() - 7 * (w - 1));
  } else if (granularity === 'monthly') {
    const m = Number(lookback) || 12;
    const thisMonth = _startOfMonth(now);
    from = new Date(thisMonth);
    from.setMonth(from.getMonth() - (m - 1)); // 直近mカ月（今月含む）
  }

  const labeller =
    (granularity === 'weekly')  ? _startOfWeek :
    (granularity === 'monthly') ? _startOfMonth : _startOfDay;

  const fmt =
    (granularity === 'weekly')  ? _formatWeekYMW :
    (granularity === 'monthly') ? _formatMonth : _formatDate;

  const byKey = new Map();
  rows.forEach(r=>{
    if (r.date < from) return;
    const bucket = labeller(r.date);
    const key = bucket.getTime();
    if (!byKey.has(key)) byKey.set(key, new Map());
    const m = byKey.get(key);
    if (!m.has(r.product)) m.set(r.product, {count:0, sum:0});
    const o = m.get(r.product);
    o.count += 1;
    o.sum   += r.amount;
  });

  const keys = [...byKey.keys()].sort((a,b)=>a-b);
  const products = new Set();
  keys.forEach(k => { for (const p of byKey.get(k).keys()) products.add(p); });
  const productList = [...products];

  const countRows = [];
  const amountRows = [];
  for (const k of keys) {
    const t = new Date(Number(k));
    const label = fmt(t);
    const m = byKey.get(k);
    const cntRow = [label]; const amtRow = [label];
    for (const p of productList) {
      const v = m.get(p) || {count:0,sum:0};
      cntRow.push(v.count);
      amtRow.push(v.sum);
    }
    countRows.push(cntRow);
    amountRows.push(amtRow);
  }
  return { products: productList, counts: countRows, amounts: amountRows, granularity, lookback: lookback || null };
}

/** 商品ごと時間帯別（1時間）件数: 週/月/累計 */
function getTimeOfDayByProduct(rangeType, product) {
  const all = _readRows_();
  const rows = all.filter(r=>!product || r.product === product);
  const now = new Date();
  const from =
    (rangeType === 'weekly')  ? new Date(now.getFullYear(), now.getMonth(), now.getDate()-6) :
    (rangeType === 'monthly') ? new Date(now.getFullYear(), now.getMonth(), 1) :
                                new Date(2000,0,1);

  const buckets = Array.from({length:24}, ()=>0); // 0..23 時

  rows.forEach(r=>{
    if (r.date < from) return;
    const h = r.date.getHours();
    buckets[h] += 1;
  });

  const labels = Array.from({length:24},(_,i)=>`${String(i).padStart(2,'0')}:00`);
  const data = labels.map((lab,i)=>[lab, buckets[i]]);
  const products = [...new Set(all.map(r=>r.product))];
  return { product, rangeType, data, labels, products };
}
