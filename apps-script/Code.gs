/**
 * CTA 名片庫｜Google Apps Script 網頁應用程式（綁在「CTA 名片庫」試算表）
 *
 * doGet  ?pw=密碼            → 回傳名片 JSON（密碼比對 Script Properties 的 PASSWORD）
 * doPost {action:'write', key, header, rows}  → 用寫入金鑰 WRITE_KEY 覆寫「名片」工作表
 * doPost {action:'init', password, key}       → 只在兩個屬性都還沒設定時可用一次，之後永久失效
 *
 * 密碼與金鑰只存在 Script Properties，程式碼裡沒有，可以放心放上 GitHub。
 */
const 工作表 = '名片';
const 備份前綴 = '備份_';
const 保留備份 = 5;

function 回應_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function 屬性_() { return PropertiesService.getScriptProperties(); }

// 防暴力猜密碼：每次錯誤延遲 2 秒；1 小時內錯 30 次就鎖 15 分鐘
function 已鎖定_() { return CacheService.getScriptCache().get('lock') === '1'; }
function 記錄失敗_() {
  const c = CacheService.getScriptCache();
  const n = Number(c.get('fail') || 0) + 1;
  c.put('fail', String(n), 3600);
  if (n >= 30) { c.put('lock', '1', 900); c.remove('fail'); }
  Utilities.sleep(2000);
}

function 等長比對_(a, b) {
  a = String(a || ''); b = String(b || '');
  if (!a || !b) return false;
  let diff = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function doGet(e) {
  if (已鎖定_()) return 回應_({ ok: false, error: 'locked' });
  const pw = e && e.parameter ? e.parameter.pw : '';
  const 正確 = 屬性_().getProperty('PASSWORD');
  if (!正確) return 回應_({ ok: false, error: 'not_initialized' });
  if (!等長比對_(pw, 正確)) { 記錄失敗_(); return 回應_({ ok: false, error: 'wrong_password' }); }
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(工作表);
  if (!sh || sh.getLastRow() < 1) return 回應_({ ok: true, header: [], rows: [], updatedAt: '' });
  const 值 = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  return 回應_({
    ok: true,
    header: 值[0],
    rows: 值.slice(1),
    updatedAt: 屬性_().getProperty('UPDATED_AT') || ''
  });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return 回應_({ ok: false, error: 'bad_json' }); }
  if (body.action === 'init') return 初始化_(body);
  if (body.action !== 'write') return 回應_({ ok: false, error: 'unknown_action' });
  if (已鎖定_()) return 回應_({ ok: false, error: 'locked' });
  if (!等長比對_(body.key, 屬性_().getProperty('WRITE_KEY'))) { 記錄失敗_(); return 回應_({ ok: false, error: 'wrong_key' }); }

  const header = body.header, rows = body.rows;
  if (!Array.isArray(header) || !header.length || !Array.isArray(rows)) return 回應_({ ok: false, error: 'bad_data' });
  if (rows.some(r => !Array.isArray(r) || r.length !== header.length)) return 回應_({ ok: false, error: 'bad_row_length' });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(工作表);
    const 舊筆數 = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
    // 防誤刪：新資料少於舊資料一半時拒絕，除非明確 force
    if (!body.force && 舊筆數 >= 10 && rows.length < 舊筆數 / 2) {
      return 回應_({ ok: false, error: 'too_few_rows', old: 舊筆數, new: rows.length });
    }
    if (sh && 舊筆數 > 0) {
      const 名 = 備份前綴 + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss');
      sh.copyTo(ss).setName(名).hideSheet();
      const 備份們 = ss.getSheets().filter(s => s.getName().indexOf(備份前綴) === 0)
        .sort((a, b) => a.getName() < b.getName() ? -1 : 1);
      備份們.slice(0, Math.max(0, 備份們.length - 保留備份)).forEach(s => ss.deleteSheet(s));
    }
    if (!sh) sh = ss.insertSheet(工作表, 0);
    sh.clear();
    const 全部 = [header].concat(rows);
    const rg = sh.getRange(1, 1, 全部.length, header.length);
    rg.setNumberFormat('@');           // 全部當文字，保留 09 開頭的手機與 001 編號
    rg.setValues(全部);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#2E4057').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    const 時間 = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    屬性_().setProperty('UPDATED_AT', 時間);
    return 回應_({ ok: true, count: rows.length, updatedAt: 時間 });
  } finally {
    lock.releaseLock();
  }
}

function 初始化_(body) {
  const p = 屬性_();
  if (p.getProperty('PASSWORD') || p.getProperty('WRITE_KEY')) return 回應_({ ok: false, error: 'already_initialized' });
  const pw = String(body.password || ''), key = String(body.key || '');
  if (pw.length < 6 || key.length < 24) return 回應_({ ok: false, error: 'too_short' });
  p.setProperties({ PASSWORD: pw, WRITE_KEY: key });
  return 回應_({ ok: true });
}

/** 在編輯器手動執行一次：建立「名片」工作表並完成授權。 */
function 初始設定() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(工作表)) ss.insertSheet(工作表, 0);
  Logger.log('完成。PASSWORD 已設定：' + !!屬性_().getProperty('PASSWORD') + '；WRITE_KEY 已設定：' + !!屬性_().getProperty('WRITE_KEY'));
}
