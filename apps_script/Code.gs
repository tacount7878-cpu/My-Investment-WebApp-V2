/**
 * 投資戰情室 V6.23
 * 功能：
 * 1. 手動更新時 → 強制覆寫整欄「目前市價」
 * 2. 覆寫後觸發所有工作表重新計算
 * 3. 保留原有交易寫入邏輯
 */

const CONFIG = {
  SPREADSHEET_ID: "1HM2MvZepqo1LVvgRoWwQ-1NmWKxo3ASAcXc2wECPgZU",
  SHEET_LOGS: "買賣紀錄_2026",
  SHEET_HISTORY: "淨值歷史",
  SHEET_ASSETS: "庫存彙整(細項)",
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
    .createHtmlOutputFromFile("apps_script/ui")
    .setTitle("投資戰情室")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================================
   2️⃣ 手動更新市價（核心）
================================ */
function updateMarketData() {

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  if (!sh) throw new Error("找不到工作表：" + CONFIG.SHEET_ASSETS);

  const lastRow = sh.getLastRow();
  if (lastRow < 6) return;

  const headers = sh.getRange(5, 1, 1, sh.getLastColumn()).getValues()[0];

  const priceCol = headers.indexOf("目前市價") + 1;
  const symbolCol = headers.indexOf("Yahoo代號(Symbol)") + 1;

  if (priceCol <= 0) throw new Error("找不到目前市價欄位");
  if (symbolCol <= 0) throw new Error("找不到Yahoo代號欄位");

  const data = sh.getRange(6, symbolCol, lastRow - 5, 1).getValues();
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

  // 🔥 強制整欄覆寫
  sh.getRange(6, priceCol, prices.length, 1).setValues(prices);

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

    if (
      json.chart &&
      json.chart.result &&
      json.chart.result.length > 0
    ) {
      return json.chart.result[0].meta.regularMarketPrice;
    }

    return "";

  } catch (e) {
    Logger.log("Fetch Error: " + symbol);
    return "";
  }
}

/* ================================
   4️⃣ Dashboard 讀取
================================ */
function getDashboardData() {

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  let history = [];

  if (histSh && histSh.getLastRow() >= 2) {
    const histData = histSh.getRange(
      Math.max(2, histSh.getLastRow() - 29),
      1,
      30,
      2
    ).getValues();

    history = histData
      .filter(r => r[0] && parseNum_(r[1]) > 0)
      .map(r => ({
        date: r[0] instanceof Date
          ? Utilities.formatDate(r[0], "GMT+8", "MM/dd")
          : String(r[0]),
        val: parseNum_(r[1])
      }));
  }

  const regionSh = ss.getSheetByName(CONFIG.SHEET_REGIONS);
  let regions = [];

  if (regionSh && regionSh.getLastRow() >= 2) {
    const regionData = regionSh.getRange(
      2, 1,
      regionSh.getLastRow() - 1,
      2
    ).getValues();

    regions = regionData.map(r => ({
      name: String(r[0] || "").trim(),
      value: parseNum_(r[1])
    })).filter(r => r.value > 0);
  }

  return { history, regions };
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

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(h => String(h || "").trim());

    const getCol = (name) => headers.indexOf(name);

    const startRow = findFirstEmptyRow_(sh);

    const rows = payload.trades.map((t, i) =>
      buildFormulaRow_(headers, payload.defaults || {}, t, startRow + i, getCol)
    );

    sh.getRange(startRow, 1, rows.length, headers.length)
      .setValues(rows);

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

  const values = sh.getRange(
    START_ROW, 1,
    lastRow - START_ROW + 1,
    1
  ).getValues();

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

  return row;
}
