/**
 * Pause it Yoho 庫存管理系統 — Google Apps Script 後端
 *
 * 使用方式：
 * 1. 開啟 https://sheets.new 建立空白試算表
 * 2. 選單「擴充功能 → Apps Script」開啟編輯器
 * 3. 刪除預設代碼，貼上本檔全部內容，儲存（Ctrl/Cmd+S）
 * 4. 「部署 → 新增部署作業」→ 齒輪選「網路應用程式」
 * 5. 執行身分：自己；存取權：任何人 → 部署並授權
 * 6. 複製產生的網址（結尾 /exec）貼到庫存系統的「設定」頁
 *
 * 資料結構：
 *   Items 工作表：id | name | cat（儲存位置） | unit | stock | min | supplier
 *   Transactions 工作表：id | itemId | type | qty | before | after | note | ts
 *   Locations 工作表：name
 */

const SHEET_ITEMS = 'Items';
const SHEET_TXS = 'Transactions';
const SHEET_LOCS = 'Locations';
const SHEET_UNITS = 'Units';
const ITEM_HEADERS = ['id', 'name', 'cat', 'unit', 'stock', 'min', 'suggest', 'supplier'];
const TX_HEADERS = ['id', 'itemId', 'type', 'qty', 'before', 'after', 'note', 'ts'];
const LOC_HEADERS = ['name'];
const UNIT_HEADERS = ['name'];

/** 取得或建立工作表（含標題列） */
function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 讀取標題列以下所有資料列 */
function readRows(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}

/** 清除標題列以下所有資料（保留標題） */
function clearData(sheet) {
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
}

/** 輸出 JSON 回應 */
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GET：讀取全部資料 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getAll';
  if (action === 'getAll') return json(getAll());
  return json({ error: 'unknown action: ' + action });
}

/** POST：寫入全部資料 */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ error: 'invalid JSON' });
  }
  if (body.action === 'saveAll') {
    try { return json(saveAll(body.data)); }
    catch (err) { return json({ ok: false, error: '伺服器錯誤：' + err.toString() }); }
  }
  if (body.action === 'savePO') return json(savePO(body.data));
  return json({ error: 'unknown action: ' + body.action });
}

/** 讀取 Items + Transactions + Locations + Units，回傳 JSON（Items 按 cat 排列） */
function getAll() {
  const itemSheet = getSheet(SHEET_ITEMS, ITEM_HEADERS);
  const txSheet = getSheet(SHEET_TXS, TX_HEADERS);
  const locSheet = getSheet(SHEET_LOCS, LOC_HEADERS);
  const unitSheet = getSheet(SHEET_UNITS, UNIT_HEADERS);
  var items = readRows(itemSheet).map(function (r) {
    var sug = Number(r[6]);
    return {
      id: r[0], name: r[1], cat: r[2], unit: r[3],
      stock: Number(r[4]) || 0, min: Number(r[5]) || 0,
      suggestedStock: isNaN(sug) ? null : sug,
      supplier: r[7] || ''
    };
  });
  items.sort(function (a, b) {
    var ca = (a.cat || '').toString();
    var cb = (b.cat || '').toString();
    if (ca < cb) return -1;
    if (ca > cb) return 1;
    return 0;
  });
  const txs = readRows(txSheet).map(function (r) {
    return {
      id: r[0], itemId: r[1], type: r[2], qty: Number(r[3]) || 0,
      before: Number(r[4]) || 0, after: Number(r[5]) || 0,
      note: r[6] || '', ts: Number(r[7]) || 0
    };
  });
  const locations = readRows(locSheet).map(function (r) { return r[0] || ''; }).filter(function (n) { return n; });
  const units = readRows(unitSheet).map(function (r) { return r[0] || ''; }).filter(function (n) { return n; });
  return { items: items, txs: txs, locations: locations, units: units, syncedAt: Date.now() };
}

/** 全量覆寫 Items + Transactions + Locations + Units（Items 按 cat 排列）
 *  ⚠️ 安全機制：避免空白／陳舊資料覆蓋掉雲端最新數據 */
function saveAll(data) {
  const items = (data.items || []).slice();
  const txs = (data.txs || []);
  const locations = (data.locations || []);
  const units = (data.units || []);

  // ▸ 安全檢查 1：拒絕寫入空白資料（items 為 0 筆）
  if (items.length === 0 && txs.length === 0) {
    return { ok: false, error: 'SAFETY_BLOCK: 拒絕上傳空白資料，請重新從雲端下載後再操作。' };
  }

  // ▸ 安全檢查 2：如果雲端已有資料，但上傳的只有極少筆數，發出警告
  const itemSheet = getSheet(SHEET_ITEMS, ITEM_HEADERS);
  const cloudItemCount = Math.max(0, itemSheet.getLastRow() - 1); // 減去標題列
  if (cloudItemCount > 10 && items.length < 3) {
    return { ok: false, error: 'SAFETY_BLOCK: 雲端現有 ' + cloudItemCount + ' 筆物料，但只上傳了 ' + items.length + ' 筆。為防止資料遺失，已拒絕此操作。請先從雲端下載最新資料。' };
  }

  // ▸ 安全機制：清除前先備份到 _Backup 工作表（備份失敗不影響主流程）
  _safeBackup();

  const txSheet = getSheet(SHEET_TXS, TX_HEADERS);
  const locSheet = getSheet(SHEET_LOCS, LOC_HEADERS);
  const unitSheet = getSheet(SHEET_UNITS, UNIT_HEADERS);

  clearData(itemSheet);
  clearData(txSheet);
  clearData(locSheet);
  clearData(unitSheet);

  items.sort(function (a, b) {
    var ca = (a.cat || '').toString();
    var cb = (b.cat || '').toString();
    if (ca < cb) return -1;
    if (ca > cb) return 1;
    return 0;
  });
  items.forEach(function (it) {
    itemSheet.appendRow([it.id, it.name, it.cat, it.unit, it.stock, it.min, it.suggestedStock != null ? it.suggestedStock : '', it.supplier || '']);
  });
  txs.forEach(function (t) {
    txSheet.appendRow([t.id, t.itemId, t.type, t.qty, t.before, t.after, t.note || '', t.ts]);
  });
  locations.forEach(function (name) {
    locSheet.appendRow([name]);
  });
  units.forEach(function (name) {
    unitSheet.appendRow([name]);
  });
  return { ok: true, items: items.length, txs: txs.length, locations: locations.length, units: units.length };
}

/** 安全備份：失敗不影響主流程 */
function _safeBackup() {
  try { _backupBeforeOverwrite(); } catch (e) {
    // 備份失敗不應阻止正常儲存
    Logger.log('Backup skipped: ' + e);
  }
}

/** 在清除資料前，將目前所有工作表內容備份到 _Backup_* 工作表（固定名稱，每次覆蓋舊備份） */
function _backupBeforeOverwrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    { name: SHEET_ITEMS, headers: ITEM_HEADERS },
    { name: SHEET_TXS, headers: TX_HEADERS },
    { name: SHEET_LOCS, headers: LOC_HEADERS },
    { name: SHEET_UNITS, headers: UNIT_HEADERS }
  ];
  sheets.forEach(function (s) {
    var src = ss.getSheetByName(s.name);
    if (!src) return;
    var lastRow = src.getLastRow();
    if (lastRow <= 1) return;
    var data = src.getRange(1, 1, lastRow, src.getLastColumn()).getValues();
    var backupName = '_Backup_' + s.name;
    var old = ss.getSheetByName(backupName);
    if (old) ss.deleteSheet(old);
    var backup = ss.insertSheet(backupName);
    backup.getRange(1, 1, data.length, data[0].length).setValues(data);
    backup.hideSheet();
  });
}

/** 建立供應商採購單新分頁 */
function savePO(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var safeName = (data.supplier || '採購單').replace(/[/\\?*:[\]]/g, '');
  var sheetName = '採購單_' + safeName + '_' + data.date;
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet(sheetName);

  sheet.getRange(1, 1).setValue('Pause it Yoho — 供應商採購單').setFontWeight('bold').setFontSize(14).setFontColor('#3b2417');
  sheet.getRange(2, 1).setValue('供應商：' + data.supplier);
  sheet.getRange(2, 4).setValue('日期：' + data.date);

  var headers = ['#', '物料名稱', '儲存位置', '當前庫存', '安全庫存', '建議採購量'];
  sheet.getRange(4, 1, 1, 6).setValues([headers]).setFontWeight('bold').setBackground('#f5ede0').setFontColor('#3b2417');

  var rows = (data.items || []).map(function (it, i) {
    return [i + 1, it.name, it.location, it.stock + ' ' + it.unit, it.min + ' ' + it.unit, it.need > 0 ? (it.need + ' ' + it.unit) : '—'];
  });
  if (rows.length > 0) {
    sheet.getRange(5, 1, rows.length, 6).setValues(rows);
    data.items.forEach(function (it, i) {
      if (it.need > 0) {
        sheet.getRange(5 + i, 1, 1, 6).setBackground('#fbf0db');
        sheet.getRange(5 + i, 6).setFontColor('#b3422b').setFontWeight('bold');
      }
    });
  }

  var sumRow = 5 + rows.length + 1;
  var needCount = (data.items || []).filter(function (it) { return it.need > 0; }).length;
  sheet.getRange(sumRow, 1).setValue('共 ' + (data.items || []).length + ' 項，其中 ' + needCount + ' 項需採購').setFontWeight('bold');
  sheet.getRange(sumRow + 2, 1).setValue('採購人：____________');
  sheet.getRange(sumRow + 2, 3).setValue('日期：____________');
  sheet.getRange(sumRow + 2, 5).setValue('供應商簽收：____________');

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 120);
  sheet.setFrozenRows(4);

  return { ok: true, sheetName: sheetName, url: ss.getUrl() };
}
