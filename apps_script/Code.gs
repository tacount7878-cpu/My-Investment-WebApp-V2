/**
 * 投資戰情室 V6.27
 * 功能升級：
 * 1. getDashboardData 接收前端輸入的現金/貸款數據。
 * 2. 自動抓取 Yahoo 匯率並與現金數據一併寫入「庫存彙整(細項)」指定儲存格。
 * 3. 整合 updateMarketData 確保股價同時更新。
 */

const CONFIG = {
  SPREADSHEET_ID: "1HM2MvZepqo1LVvgRoWwQ-1NmWKxo3ASAcXc2wECPgZU",
  SHEET_LOGS: "買賣紀錄_2026",
  SHEET_HISTORY: "淨值歷史",
  SHEET_ASSETS: "資產統計(彙整)",
  SHEET_REGIONS: "投資地區",
  SHEET_DETAILS: "庫存彙整(細項)" // 新增：指定寫入的分頁
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
    .createHtmlOutputFromFile("apps_script/ui") // 若您的檔案在根目錄，請改為 "ui"
    .setTitle("投資戰情室")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================================
   2️⃣ 手動更新市價 (內部呼叫)
   只負責更新「資產統計(彙整)」的股價
================================ */
function updateMarketData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  if (!sh) return; 

  // 自動判斷標題列 (Row 1 或 Row 5)
  let headerRow = 1;
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let symbolCol = headers.indexOf("Yahoo代號(Symbol)") + 1;

  if (symbolCol <= 0) {
    headers = sh.getRange(5, 1, 1, sh.getLastColumn()).getValues()[0];
    symbolCol = headers.indexOf("Yahoo代號(Symbol)") + 1;
    if (symbolCol > 0) headerRow = 5;
  }

  if (symbolCol <= 0) return; 

  const priceCol = headers.indexOf("目前市價") + 1;
  if (priceCol <= 0) return;

  const startRow = headerRow + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return;

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
}

/* ================================
   3️⃣ Yahoo 抓價 & 匯率
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
    return "";
  }
}

/* ================================
   4️⃣ Dashboard 讀取與寫入 (核心入口)
================================ */
function getDashboardData(inputs) {
  // 1. 先更新股價 (Yahoo -> 資產統計表)
  try { updateMarketData(); } catch (e) {}

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // 2. 處理「庫存彙整(細項)」的寫入 (匯率 + UI 輸入值)
  const detailSh = ss.getSheetByName(CONFIG.SHEET_DETAILS);
  let freshUsdRate = 32.2; // 預設值

  if (detailSh) {
    // 2.1 抓取即時匯率
    const fetchedRate = fetchYahooPrice("USDTWD=X");
    if (fetchedRate && !isNaN(fetchedRate)) {
      freshUsdRate = Number(fetchedRate);
    }

    // 2.2 寫入匯率到 A2
    detailSh.getRange("A2").setValue(freshUsdRate);

    // 2.3 寫入前端傳來的現金與貸款 (如果有傳的話)
    if (inputs) {
      // 確保轉為數字再寫入
      if (inputs.cashTwd !== undefined && inputs.cashTwd !== "") detailSh.getRange("C2").setValue(Number(inputs.cashTwd));
      if (inputs.settleTwd !== undefined && inputs.settleTwd !== "") detailSh.getRange("E2").setValue(Number(inputs.settleTwd));
      if (inputs.cashUsd !== undefined && inputs.cashUsd !== "") detailSh.getRange("G2").setValue(Number(inputs.cashUsd));
      if (inputs.loanTwd !== undefined && inputs.loanTwd !== "") detailSh.getRange("I2").setValue(Number(inputs.loanTwd));
    }
  }

  // 3. 強制刷新計算 (確保剛寫入的數字被公式吃到)
  SpreadsheetApp.flush();

  /* ===== 以下為資料讀取 ===== */

  // 1. 淨值歷史
  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  let history = [];
  if (histSh && histSh.getLastRow() >= 2) {
    const histData = histSh.getRange(Math.max(2, histSh.getLastRow() - 29), 1, 30, 2).getValues();
    history = histData.filter(r => r[0] && parseNum_(r[1]) > 0).map(r => ({
      date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+8", "MM/dd") : String(r[0]),
      val: parseNum_(r[1])
    }));
  }

  // 2. 資產統計 (依據您 V6.25 的修正，讀取 Row 1 標題)
  const assetSh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  let investTotal = 0;
  let assets = [];
  if (assetSh && assetSh.getLastRow() >= 2) {
    const headers = assetSh.getRange(1, 1, 1, assetSh.getLastColumn()).getValues()[0];
    const valueCol = headers.indexOf("市值(TWD)") + 1;
    let nameCol = headers.indexOf("合併鍵(GroupKey)") + 1; 
    if (nameCol <= 0) nameCol = headers.indexOf("標的名稱") + 1;

    if (valueCol > 0 && nameCol > 0) {
      const numRows = assetSh.getLastRow() - 1;
      const values = assetSh.getRange(2, valueCol, numRows, 1).getValues();
      const names = assetSh.getRange(2, nameCol, numRows, 1).getValues();
      for (let i = 0; i < values.length; i++) {
        const val = parseNum_(values[i][0]);
        if (val > 0) {
          investTotal += val;
          assets.push({ name: String(names[i][0] || ""), value: val });
        }
      }
    }
  }

  // 3. 投資地區
  const regionSh = ss.getSheetByName(CONFIG.SHEET_REGIONS);
  let regions = [];
  if (regionSh && regionSh.getLastRow() >= 2) {
    const regionData = regionSh.getRange(2, 1, regionSh.getLastRow() - 1, 2).getValues();
    regions = regionData.map(r => ({ name: String(r[0] || "").trim(), value: parseNum_(r[1]) })).filter(r => r.value > 0);
  }

  // 4. 摘要數據 (這裡可以繼續讀取買賣紀錄表的摘要，或依您的需求調整)
  const logSh = ss.getSheetByName(CONFIG.SHEET_LOGS);
  let realizedReturn = 0, realizedReturnTwd = 0;
  if (logSh) {
    const summaryData = logSh.getRange("Y1:Z30").getValues();
    for(let row of summaryData){
      const label = String(row[0]);
      const val = row[1];
      if(label.includes("已實現總損益(TWD)")) realizedReturnTwd = parseNum_(val);
      if(label.includes("已實現總損益(%)")) realizedReturn = (Number(String(val).replace("%","")) || 0) * (String(val).includes("%") ? 1 : 100);
    }
  }

  return {
    history,
    assets,
    regions,
    investTotal,
    usdRate: freshUsdRate,
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
  return Number(String(val).replace(/,/g, "")) || 0;
}

/* ================================
   6️⃣ 交易寫入
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
    const rows = payload.trades.map((t, i) => buildFormulaRow_(headers, payload.defaults || {}, t, startRow + i, getCol));
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
  const setVal = (name, val) => { const idx = getCol(name); if (idx !== -1) row[idx] = val; };
  setVal("日期", t.date || new Date());
  setVal("交易類型", t.type);
  setVal("名稱", t.name);
  setVal("股票代號", t.symbol);
  setVal("買入價格", Number(t.price));
  setVal("買入股數", Number(t.qty));
  setVal("狀態", "已完成");
  return row;
}