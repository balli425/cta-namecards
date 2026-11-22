/**
 * CTA 名片庫｜Google Apps Script 網頁應用程式（綁在「CTA 名片庫」試算表）
 *
 * doGet  ?pw=密碼            → 回傳名片 JSON（密碼比對 Script Properties 的 PASSWORD）
 * doPost {action:'write', key, header, rows}  → 用寫入金鑰 WRITE_KEY 覆寫「名片」工作表
 * doPost {action:'init', password, key}       → 只在兩個屬性都還沒設定時可用一次，之後永久失效
 * doPost {action:'set_password', key, password} → 用寫入金鑰更改查詢密碼（本機 同步到Google.py --改密碼 新密碼）
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
  if (body.action === 'upload' || body.action === 'inbox_count') return 手機上傳_(body);
  if (['inbox_list', 'inbox_file', 'inbox_done'].indexOf(body.action) >= 0) return 電腦取件_(body);
  if (body.action !== 'write' && body.action !== 'set_password') return 回應_({ ok: false, error: 'unknown_action' });
  if (已鎖定_()) return 回應_({ ok: false, error: 'locked' });
  if (!等長比對_(body.key, 屬性_().getProperty('WRITE_KEY'))) { 記錄失敗_(); return 回應_({ ok: false, error: 'wrong_key' }); }
  if (body.action === 'set_password') {
    const np = String(body.password || '');
    if (np.length < 6) return 回應_({ ok: false, error: 'too_short' });
    屬性_().setProperty('PASSWORD', np);
    return 回應_({ ok: true });
  }

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

/* ===== 手機拍名片上傳（2026-11-22 新增） =====
 * 照片暫存在本試算表的隱藏工作表「待歸檔照片」（不需要雲端硬碟權限，只用原本的試算表權限）。
 * 一列＝一張照片：代碼｜狀態｜上傳時間｜面｜備註｜歸檔結果｜片段數｜資料片段…（base64 切成每格 45,000 字）
 * 同一張名片的正反面共用同一個「代碼」。
 * 電腦取回照片、歸檔完成後，狀態改「已處理」並清掉照片資料（照片已存進 Dropbox 的照片資料夾），保留紀錄列。
 * doPost {action:'upload', pw, id, images:[{side:'正面'|'背面', data:base64}], note}  → 用查詢密碼（手機已記住）
 * doPost {action:'inbox_count', pw}                → 待歸檔幾張名片
 * doPost {action:'inbox_list', key}                → 電腦用寫入金鑰：列出待歸檔
 * doPost {action:'inbox_file', key, id}            → 取一張照片（base64），id＝代碼_面
 * doPost {action:'inbox_done', key, folder, result} → 標記已處理（folder＝代碼）
 */
const 照片表名 = '待歸檔照片';
const 照片表頭 = ['代碼', '狀態', '上傳時間', '面', '備註', '歸檔結果', '片段數'];
const 片段長 = 45000;

function 照片表_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(照片表名);
  if (!sh) {
    sh = ss.insertSheet(照片表名);
    sh.getRange(1, 1, 1, 照片表頭.length).setValues([照片表頭]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}
// 只讀前 7 欄（不含照片資料），回傳 [{row, 代碼, 狀態, 上傳時間, 面, 備註, 片段數}]
function 照片列_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 照片表頭.length).getValues().map((r, i) => ({
    row: i + 2, 代碼: String(r[0]), 狀態: String(r[1]), 上傳時間: String(r[2]), 面: String(r[3]), 備註: String(r[4]), 片段數: Number(r[6]) || 0
  }));
}
function 待歸檔代碼_(列) {
  const s = [];
  列.forEach(x => { if (x.狀態 === '待歸檔' && s.indexOf(x.代碼) < 0) s.push(x.代碼); });
  return s;
}

function 手機上傳_(body) {
  if (已鎖定_()) return 回應_({ ok: false, error: 'locked' });
  if (!等長比對_(body.pw, 屬性_().getProperty('PASSWORD'))) { 記錄失敗_(); return 回應_({ ok: false, error: 'wrong_password' }); }
  const sh = 照片表_();
  if (body.action === 'inbox_count') return 回應_({ ok: true, count: 待歸檔代碼_(照片列_(sh)).length });

  const id = String(body.id || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 12);
  const imgs = Array.isArray(body.images) ? body.images : [];
  if (!id || !imgs.length || imgs.length > 2) return 回應_({ ok: false, error: 'bad_data' });
  if (imgs.some(x => !x || typeof x.data !== 'string' || !x.data || x.data.length > 4000000)) return 回應_({ ok: false, error: 'too_large' });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const 列 = 照片列_(sh);
    // 同一張重送（手機網路不穩重試）不會多一份
    if (列.some(x => x.代碼 === id)) return 回應_({ ok: true, dup: true, count: 待歸檔代碼_(列).length });
    const 時間 = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
    const 備註 = String(body.note || '').slice(0, 500);
    imgs.forEach((x, i) => {
      const side = i === 0 ? '正面' : '背面';
      const 片 = [];
      for (let p = 0; p < x.data.length; p += 片段長) 片.push(x.data.slice(p, p + 片段長));
      const row = [id, '待歸檔', 時間, side, 備註, '', 片.length].concat(片);
      const r = sh.getLastRow() + 1;
      const rg = sh.getRange(r, 1, 1, row.length);
      rg.setNumberFormat('@');
      rg.setValues([row]);
    });
    return 回應_({ ok: true, folder: id, count: 待歸檔代碼_(照片列_(sh)).length });
  } finally {
    lock.releaseLock();
  }
}

function 電腦取件_(body) {
  if (已鎖定_()) return 回應_({ ok: false, error: 'locked' });
  if (!等長比對_(body.key, 屬性_().getProperty('WRITE_KEY'))) { 記錄失敗_(); return 回應_({ ok: false, error: 'wrong_key' }); }
  const sh = 照片表_();
  const 列 = 照片列_(sh);

  if (body.action === 'inbox_list') {
    const 清單 = 待歸檔代碼_(列).map(代碼 => {
      const 張 = 列.filter(x => x.代碼 === 代碼 && x.狀態 === '待歸檔');
      return {
        folder: 代碼,
        info: { 上傳時間: 張[0].上傳時間, 張數: 張.length, 面: 張.map(x => x.面), 備註: 張[0].備註 },
        files: 張.map(x => ({ id: 代碼 + '_' + x.面, name: (x.面 === '正面' ? '1' : '2') + '_' + x.面 + '.jpg' }))
      };
    });
    return 回應_({ ok: true, items: 清單 });
  }
  if (body.action === 'inbox_file') {
    const m = String(body.id || '').match(/^([A-Za-z0-9]+)_(正面|背面)$/);
    const x = m && 列.filter(r => r.代碼 === m[1] && r.面 === m[2] && r.狀態 === '待歸檔')[0];
    if (!x) return 回應_({ ok: false, error: 'not_found' });
    const data = sh.getRange(x.row, 照片表頭.length + 1, 1, x.片段數).getValues()[0].join('');
    return 回應_({ ok: true, name: (x.面 === '正面' ? '1' : '2') + '_' + x.面 + '.jpg', data: data });
  }
  // inbox_done：狀態改已處理、寫入歸檔結果、清掉照片資料（照片已存進 Dropbox）
  const 代碼 = String(body.folder || '');
  const 張 = 列.filter(x => x.代碼 === 代碼 && x.狀態 === '待歸檔');
  if (!張.length) return 回應_({ ok: false, error: 'not_found' });
  const 結果 = JSON.stringify(body.result || {}).slice(0, 45000);
  張.forEach(x => {
    sh.getRange(x.row, 2).setValue('已處理');
    sh.getRange(x.row, 6).setValue(結果);
    if (x.片段數) sh.getRange(x.row, 照片表頭.length + 1, 1, x.片段數).clearContent();
  });
  return 回應_({ ok: true, count: 待歸檔代碼_(照片列_(sh)).length });
}

/** 在編輯器手動執行一次：建立「名片」與「待歸檔照片」工作表，並完成授權。 */
function 初始設定() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(工作表)) ss.insertSheet(工作表, 0);
  照片表_();
  Logger.log('完成。PASSWORD 已設定：' + !!屬性_().getProperty('PASSWORD') + '；WRITE_KEY 已設定：' + !!屬性_().getProperty('WRITE_KEY') + '；隱藏工作表：' + 照片表名);
}
