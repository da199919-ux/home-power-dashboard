/**
 * 家庭用電儀表板：每日快照（與既有的 syncPowerHistory 程式並存，名稱都加上 dash 前綴）
 * - dashSetup()：建立每天 23:50（台北）的觸發條件，並立即產生一次快照
 * - dashRefreshSnapshot()：把快照 JSON 存到隱藏工作表「_儀表板快照」A1
 * - doGet()：網頁應用程式，回傳最近一次的快照 JSON 給儀表板
 */
const DASH_TZ = 'Asia/Taipei';
const DASH_SNAP_SHEET = '_儀表板快照';
const DASH_APPS = ['熱水器', '咖啡機', '飲水機', '廚房熱水器', '門口燈', '除濕機'];
const DASH_DAILY_DAYS = 120;

function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASH_SNAP_SHEET);
  const json = (sh && sh.getRange('A1').getValue()) || JSON.stringify(dashBuildSnapshot_());
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function dashSetup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dashRefreshSnapshot')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dashRefreshSnapshot').timeBased().everyDays(1).atHour(23).nearMinute(50).inTimezone(DASH_TZ).create();
  dashRefreshSnapshot();
}

function dashRefreshSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DASH_SNAP_SHEET);
  if (!sh) { sh = ss.insertSheet(DASH_SNAP_SHEET); sh.hideSheet(); }
  sh.getRange('A1').setValue(JSON.stringify(dashBuildSnapshot_()));
}

function dashNum_(v) {
  if (v === '' || v == null) return null;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}
function dashParseTime_(text) {
  const m = String(text).match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})\D+(\d{1,2}):(\d{2})/);
  return m ? { ms: Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5]) } : null;
}
function dashPad_(n) { return String(n).padStart(2, '0'); }
function dashReadTable_(name, lastN) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const last = sh.getLastRow(), cols = sh.getLastColumn();
  if (last < 2) return [];
  const head = sh.getRange(1, 1, 1, cols).getDisplayValues()[0];
  const start = lastN ? Math.max(2, last - lastN + 1) : 2;
  return sh.getRange(start, 1, last - start + 1, cols).getDisplayValues()
    .filter(r => r.some(c => c !== ''))
    .map(r => Object.fromEntries(head.map((k, i) => [k, r[i]])));
}

function dashBuildSnapshot_() {
  const daily = dashReadTable_('用電每日彙總').map(r => {
    const m = String(r['日期']).match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    const date = m ? m[1] + '-' + dashPad_(m[2]) + '-' + dashPad_(m[3]) : String(r['日期']);
    return [date, dashNum_(r['完整度%']), dashNum_(r['全屋購電kWh']), dashNum_(r['全屋售電kWh']), dashNum_(r['太陽能發電kWh']),
      dashNum_(r['一樓購電kWh']), dashNum_(r['樓上購電kWh']), dashNum_(r['平均負載W']), dashNum_(r['最大負載W']), r['彙總狀態'],
      ...DASH_APPS.map(a => dashNum_(r[a + '用電kWh']))];
  }).sort((a, b) => a[0] < b[0] ? -1 : 1).slice(-DASH_DAILY_DAYS);

  const snaps = dashReadTable_('用電快照', 900)
    .map(r => ({ r, t: dashParseTime_(r['時段']) }))
    .filter(x => x.t && dashNum_(x.r['全屋負載W']) != null)
    .sort((a, b) => a.t.ms - b.t.ms);
  if (!snaps.length) throw new Error('用電快照沒有資料');
  const L = snaps[snaps.length - 1].r, lastMs = snaps[snaps.length - 1].t.ms;

  const buckets = new Map();
  snaps.filter(x => x.t.ms > lastMs - 48 * 3600e3).forEach(({ r, t }) => {
    const k = Math.floor(t.ms / 900e3) * 900e3;
    const b = buckets.get(k) || [[], [], []];
    ['全屋負載W', '全屋太陽能W', '全屋市電W'].forEach((c, i) => { const v = dashNum_(r[c]); if (v != null) b[i].push(v); });
    buckets.set(k, b);
  });
  const avg = a => a.length ? Math.round(a.reduce((x, y) => x + y, 0) / a.length) : null;
  const curve = [...buckets.entries()].sort((a, b) => a[0] - b[0]).map(([k, b]) => {
    const d = new Date(k);
    return [dashPad_(d.getUTCMonth() + 1) + '-' + dashPad_(d.getUTCDate()) + ' ' + dashPad_(d.getUTCHours()) + ':' + dashPad_(d.getUTCMinutes()), ...b.map(avg)];
  });

  return {
    generatedAt: Utilities.formatDate(new Date(), DASH_TZ, 'yyyy-MM-dd HH:mm'),
    latest: {
      time: L['時段'], status: L['資料狀態'], tuya: L['Tuya狀態'], growatt: L['Growatt狀態'], worker: L['Worker版本'], error: L['錯誤訊息'],
      loadW: dashNum_(L['全屋負載W']), solarW: dashNum_(L['全屋太陽能W']), gridW: dashNum_(L['全屋市電W']),
      grid1W: dashNum_(L['一樓市電W']), grid2W: dashNum_(L['樓上市電W']),
      appliances: DASH_APPS.map(a => [a, dashNum_(L[a + '功率W']), dashNum_(L[a + '今日kWh'])])
    },
    daily, curve
  };
}
