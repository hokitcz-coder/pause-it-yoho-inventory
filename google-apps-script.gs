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
const ITEM_HEADERS = ['id', 'name', 'cat', 'unit', 'stock', 'min', 'suggest', 'price', 'supplier'];
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
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

/** 確保工作表有足夠的行數來寫入資料 */
function ensureRows(sheet, neededRows) {
  const current = sheet.getMaxRows();
  const headerAndNeeded = 1 + neededRows; // 標題 + 資料行
  if (current < headerAndNeeded) {
    sheet.insertRows(current + 1, headerAndNeeded - current);
  }
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
  if (action === 'monthlyReport') {
    const ym = (e && e.parameter && e.parameter.ym) || '';
    return json(getMonthlyReport(ym));
  }
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
  if (body.action === 'saveAll') return json(saveAll(body.data));
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
      price: Number(r[7]) || 0,
      supplier: r[8] || ''
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

  // ▸ 效能優化：資料沒變就跳過寫入
  const hash = _hash(JSON.stringify({ i: items, t: txs, l: locations, u: units }));
  const props = PropertiesService.getDocumentProperties();
  const lastHash = props.getProperty('LAST_SAVE_HASH');
  if (hash === lastHash) {
    return { ok: true, items: items.length, txs: txs.length, skipped: true };
  }

  // 安全檢查通過，清除舊資料後寫入
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

  // 批量寫入（setValues 比逐行 appendRow 快數十倍）
  if (items.length > 0) {
    ensureRows(itemSheet, items.length);
    var itemRows = items.map(function (it) {
      return [it.id, it.name, it.cat, it.unit, it.stock, it.min, it.suggestedStock != null ? it.suggestedStock : '', it.price != null ? it.price : '', it.supplier || ''];
    });
    itemSheet.getRange(2, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
  }
  if (txs.length > 0) {
    ensureRows(txSheet, txs.length);
    var txRows = txs.map(function (t) {
      return [t.id, t.itemId, t.type, t.qty, t.before, t.after, t.note || '', t.ts];
    });
    txSheet.getRange(2, 1, txRows.length, txRows[0].length).setValues(txRows);
  }
  if (locations.length > 0) {
    ensureRows(locSheet, locations.length);
    var locRows = locations.map(function (name) { return [name]; });
    locSheet.getRange(2, 1, locRows.length, 1).setValues(locRows);
  }
  if (units.length > 0) {
    ensureRows(unitSheet, units.length);
    var unitRows = units.map(function (name) { return [name]; });
    unitSheet.getRange(2, 1, unitRows.length, 1).setValues(unitRows);
  }
  props.setProperty('LAST_SAVE_HASH', hash);
  return { ok: true, items: items.length, txs: txs.length, locations: locations.length, units: units.length };
}

/** 月度報表：匯總當月入庫金額、出庫成本、消耗排行
 *  @param {string} ym - 格式 YYYY-MM，不傳則用當月 */
function getMonthlyReport(ym) {
  if (!ym) {
    var now = new Date();
    ym = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
  }
  var parts = ym.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  var monthStart = new Date(y, m - 1, 1).getTime();
  var monthEnd = new Date(y, m, 1).getTime();

  var itemSheet = getSheet(SHEET_ITEMS, ITEM_HEADERS);
  var allItems = readRows(itemSheet);
  var itemMap = {};
  allItems.forEach(function (r) {
    itemMap[r[0]] = { name: r[1] || '', unit: r[3] || '', price: Number(r[7]) || 0 };
  });

  var txSheet = getSheet(SHEET_TXS, TX_HEADERS);
  var allTxs = readRows(txSheet);
  var monthTxs = allTxs.filter(function (r) {
    var ts = Number(r[7]) || 0;
    return ts >= monthStart && ts < monthEnd;
  });

  var itemStats = {};
  var totalInAmt = 0, totalOutAmt = 0;

  monthTxs.forEach(function (r) {
    var itemId = r[1];
    var type = r[2];
    var qty = Number(r[3]) || 0;
    var info = itemMap[itemId] || { name: '(已刪除)', unit: '', price: 0 };
    var price = info.price;
    var amt = Math.round(qty * price * 100) / 100;

    if (!itemStats[itemId]) {
      itemStats[itemId] = { id: itemId, name: info.name, unit: info.unit, price: price, inQty: 0, outQty: 0, inAmt: 0, outAmt: 0 };
    }
    if (type === 'in') {
      itemStats[itemId].inQty = Math.round((itemStats[itemId].inQty + qty) * 1000) / 1000;
      itemStats[itemId].inAmt = Math.round((itemStats[itemId].inAmt + amt) * 100) / 100;
      totalInAmt = Math.round((totalInAmt + amt) * 100) / 100;
    } else if (type === 'out') {
      itemStats[itemId].outQty = Math.round((itemStats[itemId].outQty + qty) * 1000) / 1000;
      itemStats[itemId].outAmt = Math.round((itemStats[itemId].outAmt + amt) * 100) / 100;
      totalOutAmt = Math.round((totalOutAmt + amt) * 100) / 100;
    }
  });

  var list = Object.keys(itemStats).map(function (k) { return itemStats[k]; });
  list.sort(function (a, b) { return b.outAmt - a.outAmt; });

  var inCount = monthTxs.filter(function (r) { return r[2] === 'in'; }).length;
  var outCount = monthTxs.filter(function (r) { return r[2] === 'out'; }).length;

  return {
    ym: ym,
    period: y + '年' + m + '月',
    totalInAmt: totalInAmt,
    totalOutAmt: totalOutAmt,
    inCount: inCount,
    outCount: outCount,
    itemCount: list.length,
    items: list
  };
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

/** 計算字串的 MD5 哈希（用於比對資料是否變更） */
function _hash(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
