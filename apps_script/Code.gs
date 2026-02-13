/**
 * 投資戰情室 V6.25
 * 修正重點：
 * 1. 修正 getDashboardData 讀取「資產統計(彙整)」的列數 (從 Row 5 改為 Row 1)。
 * 2. 修正 updateMarketData 尋找欄位的邏輯，增加容錯。
 */

const CONFIG = {
  SPREADSHEET_ID: "1HM2MvZepqo1LVvgRoWwQ-1NmWKxo3ASAcXc2wECPgZU",
  SHEET_LOGS: "買賣紀錄_2026",
  SHEET_HISTORY: "淨值歷史",
  SHEET_ASSETS: "資產統計(彙整)",
  SHEET_REGIONS: "投資地區"
};

/* ================================
   0️⃣ 強制授權
================================ */
function forceAuth() {
  UrlFetchApp.fetch("https://www.google.com");
  Logger.log("授權完成");
}

/* ================================
   1️⃣ 網頁入口
================================ */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("apps_script/ui") // 請確認您的檔案是在 apps_script 資料夾下還是在根目錄，若在根目錄請改為 "ui"
    .setTitle("投資戰情室")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================================
   2️⃣ 手動更新市價（核心）
================================ */
function updateMarketData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  if (!sh) return "找不到工作表";

  // 嘗試自動判斷標題列在第幾列 (優先找 Row 1, 找不到找 Row 5)
  let headerRow = 1;
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let symbolCol = headers.indexOf("Yahoo代號(Symbol)") + 1;

  if (symbolCol <= 0) {
    // 如果第 1 列找不到代號，試試看第 5 列 (相容舊格式)
    headers = sh.getRange(5, 1, 1, sh.getLastColumn()).getValues()[0];
    symbolCol = headers.indexOf("Yahoo代號(Symbol)") + 1;
    if (symbolCol > 0) headerRow = 5;
  }

  // 如果還是找不到代號欄位，代表這張表可能不支援自動更新，直接跳過不報錯
  if (symbolCol <= 0) {
    console.warn("無法執行自動更新：找不到 'Yahoo代號(Symbol)' 欄位");
    return "跳過更新";
  }

  const priceCol = headers.indexOf("目前市價") + 1;
  if (priceCol <= 0) return "找不到市價欄";

  const startRow = headerRow + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return "無資料";

  const data = sh.getRange(startRow, symbolCol, lastRow - startRow + 1, 1).getValues();
  const prices = [];

  for (let i = 0; i < data.length; i++) {
    const symbol = String(data[i][0] || "").trim();
    if (!symbol) {
      prices.push([""]);
      continue;
    }
    const price = fetchYahooPrice(symbol);
    prices.push([price]);
    Utilities.sleep(20);
  }

  sh.getRange(startRow, priceCol, prices.length, 1).setValues(prices);
  SpreadsheetApp.flush();
  return "更新完成";
}

/* ================================
   3️⃣ Yahoo 抓價
================================ */
function fetchYahooPrice(symbol) {
  try {
    const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?interval=1d`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (json.chart && json.chart.result && json.chart.result.length > 0) {
      return json.chart.result[0].meta.regularMarketPrice;
    }
    return "";
  } catch (e) {
    Logger.log("Fetch Error: " + symbol);
    return "";
  }
}

/* ================================
   4️⃣ Dashboard 讀取 (修正讀取位置)
================================ */
function getDashboardData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  /* ===== 1. 淨值歷史 ===== */
  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  let history = [];
  if (histSh && histSh.getLastRow() >= 2) {
    const histData = histSh.getRange(Math.max(2, histSh.getLastRow() - 29), 1, 30, 2).getValues();
    history = histData
      .filter(r => r[0] && parseNum_(r[1]) > 0)
      .map(r => ({
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+8", "MM/dd") : String(r[0]),
        val: parseNum_(r[1])
      }));
  }

  /* ===== 2. 資產統計(彙整) - 修正讀取 Row 1 ===== */
  const assetSh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  let investTotal = 0;
  let assets = [];

  if (assetSh && assetSh.getLastRow() >= 2) {
    // 🔥 修正：讀取第 1 列的標題 (原本是第 5 列)
    const headers = assetSh.getRange(1, 1, 1, assetSh.getLastColumn()).getValues()[0];
    
    // 對應您截圖中的欄位名稱
    const valueCol = headers.indexOf("市值(TWD)") + 1;
    // 使用「合併鍵(GroupKey)」作為名稱，若找不到則找「標的名稱」
    let nameCol = headers.indexOf("合併鍵(GroupKey)") + 1; 
    if (nameCol <= 0) nameCol = headers.indexOf("標的名稱") + 1;

    // 只有在找到欄位時才讀取
    if (valueCol > 0 && nameCol > 0) {
      // 資料從第 2 列開始 (Row 2)
      const numRows = assetSh.getLastRow() - 1;
      const values = assetSh.getRange(2, valueCol, numRows, 1).getValues();
      const names = assetSh.getRange(2, nameCol, numRows, 1).getValues();

      for (let i = 0; i < values.length; i++) {
        const val = parseNum_(values[i][0]);
        // 過濾掉 0 或負數，確保圓餅圖不報錯
        if (val > 0) {
          investTotal += val;
          assets.push({
            name: String(names[i][0] || ""),
            value: val
          });
        }
      }
    }
  }

  /* ===== 3. 投資地區 ===== */
  const regionSh = ss.getSheetByName(CONFIG.SHEET_REGIONS);
  let regions = [];
  if (regionSh && regionSh.getLastRow() >= 2) {
    const regionData = regionSh.getRange(2, 1, regionSh.getLastRow() - 1, 2).getValues();
    regions = regionData.map(r => ({
      name: String(r[0] || "").trim(),
      value: parseNum_(r[1])
    })).filter(r => r.value > 0);
  }

  /* ===== 4. 讀取摘要數據 (報酬率/損益/匯率) ===== */
  // 嘗試從買賣紀錄表讀取 (根據您之前的截圖位置 Y1:Z30)
  const logSh = ss.getSheetByName(CONFIG.SHEET_LOGS);
  let realizedReturn = 0;
  let realizedReturnTwd = 0;
  let usdRate = 32.2; // 預設值

  if (logSh) {
    const summaryData = logSh.getRange("Y1:Z30").getValues();
    for(let row of summaryData){
      const label = String(row[0]);
      const val = row[1];
      if(label.includes("已實現總損益(TWD)")) realizedReturnTwd = parseNum_(val);
      if(label.includes("已實現總損益(%)")) realizedReturn = (Number(String(val).replace("%","")) || 0) * (String(val).includes("%") ? 1 : 100);
    }
    // 嘗試讀取匯率 (假設在 H86)
    const rateVal = logSh.getRange("H86").getValue();
    if(typeof rateVal === 'number' && rateVal > 0) usdRate = rateVal;
  }

  return {
    history,
    assets,
    regions,
    investTotal,
    usdRate,
    realizedReturn,
    realizedReturnTwd
  };
}

/* ================================
   5️⃣ 數字安全解析
================================ */
function parseNum_(val) {
  if (val === "" || val === null || val === undefined) return 0;
  if (typeof val === "number") return val;
  // 移除逗號再轉數字
  return Number(String(val).replace(/,/g, "")) || 0;
}

/* ================================
   6️⃣ 交易寫入 (維持不變)
================================ */
function saveTrades(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.SHEET_LOGS);
    if (!sh) throw new Error("找不到分頁");

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const getCol = (name) => headers.indexOf(name);
    const startRow = findFirstEmptyRow_(sh);
    const rows = payload.trades.map((t, i) =>
      buildFormulaRow_(headers, payload.defaults || {}, t, startRow + i, getCol)
    );
    sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function findFirstEmptyRow_(sh) {
  const START_ROW = 86;
  const lastRow = sh.getLastRow();
  if (lastRow < START_ROW) return START_ROW;
  const values = sh.getRange(START_ROW, 1, lastRow - START_ROW + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) return START_ROW + i;
  }
  return lastRow + 1;
}

function buildFormulaRow_(headers, defaults, t, r, getCol) {
  const row = new Array(headers.length).fill("");
  const setVal = (name, val) => {
    const idx = getCol(name);
    if (idx !== -1) row[idx] = val;
  };
  setVal("日期", t.date || new Date());
  setVal("交易類型", t.type);
  setVal("名稱", t.name);
  setVal("股票代號", t.symbol);
  setVal("買入價格", Number(t.price));
  setVal("買入股數", Number(t.qty));
  setVal("狀態", "已完成");
  // ... 其他公式 ...
  return row;
}