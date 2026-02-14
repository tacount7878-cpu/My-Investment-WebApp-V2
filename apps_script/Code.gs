/**
 * 投資戰情室 V6.41 - 後端核心代碼
 * 主要功能：
 * 1. 數據同步：接收 UI 數據並寫入「庫存彙整(細項)」(A2, C2, E2, G2, I2)。
 * 2. 股價更新：手動更新時從 Yahoo Finance 抓取最新股價。
 * 3. 歷史記錄：自動記錄每日資產總淨值於「淨值歷史」。
 * 4. 圖表支援：提供折線圖與圓餅圖所需之數據結構。
 */

const CONFIG = {
  SPREADSHEET_ID: "1HM2MvZepqo1LVvgRoWwQ-1NmWKxo3ASAcXc2wECPgZU",
  SHEET_LOGS: "買賣紀錄_2026",
  SHEET_HISTORY: "淨值歷史",
  SHEET_ASSETS: "資產統計(彙整)", 
  SHEET_DETAILS: "庫存彙整(細項)" 
};

/* ================================
   1️⃣ 網頁入口
================================ */
function doGet() {
  const possibleNames = ["ui", "ui.html", "Index", "apps_script/ui"];
  for (let name of possibleNames) {
    try {
      return HtmlService.createHtmlOutputFromFile(name)
        .setTitle("投資戰情室")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (e) {}
  }
  return HtmlService.createHtmlOutput("找不到網頁檔案，請確保檔案名稱為 ui");
}

/* ================================
   2️⃣ 手動更新股價與匯率 (Yahoo)
================================ */
function updateMarketData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_DETAILS);
  if (!sh) return;

  // 根據截圖，標題在第 5 列
  const headerRow = 5; 
  const lastRow = sh.getLastRow();
  if (lastRow <= headerRow) return;

  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  const symbolCol = headers.indexOf("Yahoo代號(Symbol)") + 1;
  const priceCol = headers.indexOf("目前市價") + 1;

  if (symbolCol <= 0 || priceCol <= 0) return;

  const data = sh.getRange(headerRow + 1, symbolCol, lastRow - headerRow, 1).getValues();
  const prices = [];

  for (let i = 0; i < data.length; i++) {
    const symbol = String(data[i][0] || "").trim();
    if (!symbol) {
      prices.push([""]);
      continue;
    }
    const price = fetchYahooPrice(symbol);
    prices.push([price]);
    Utilities.sleep(10); // 避免過度請求
  }

  sh.getRange(headerRow + 1, priceCol, prices.length, 1).setValues(prices);
}

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
   3️⃣ Dashboard 核心邏輯
================================ */
function getDashboardData(inputs, isManualUpdate) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const detailSh = ss.getSheetByName(CONFIG.SHEET_DETAILS);
  let freshUsdRate = 32.2; 

  // --- 處理寫入邏輯 (僅限手動更新時) ---
  if (isManualUpdate === true) {
    try { updateMarketData(); } catch (e) {}
    
    // 獲取即時匯率
    const fetchedRate = fetchYahooPrice("USDTWD=X");
    if (fetchedRate && !isNaN(fetchedRate)) freshUsdRate = Number(fetchedRate);

    if (detailSh) {
      // A2: 匯率
      detailSh.getRange("A2").setValue(freshUsdRate);
      if (inputs) {
        // C2: 台幣現金, E2: 交割中, G2: 美元現金, I2: 貸款
        if (inputs.cashTwd !== "") detailSh.getRange("C2").setValue(Number(inputs.cashTwd));
        if (inputs.settleTwd !== "") detailSh.getRange("E2").setValue(Number(inputs.settleTwd));
        if (inputs.cashUsd !== "") detailSh.getRange("G2").setValue(Number(inputs.cashUsd));
        if (inputs.loanTwd !== "") detailSh.getRange("I2").setValue(Number(inputs.loanTwd));
      }
    }
    SpreadsheetApp.flush(); // 強制刷新試算表公式
  } else if (detailSh) {
    // 初始載入則讀取既有匯率
    freshUsdRate = Number(detailSh.getRange("A2").getValue()) || 32.2;
  }

  // --- 讀取資產細項 (用於圓餅圖) ---
  const assetSh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  let investTotal = 0;
  let assets = [];
  
  if (assetSh && assetSh.getLastRow() >= 2) {
    const headers = assetSh.getRange(1, 1, 1, assetSh.getLastColumn()).getValues()[0];
    const valueCol = headers.indexOf("市值(TWD)") + 1;
    const nameCol = headers.indexOf("合併鍵(GroupKey)") + 1 || headers.indexOf("標的名稱") + 1;

    if (valueCol > 0 && nameCol > 0) {
      const rows = assetSh.getLastRow() - 1;
      const vals = assetSh.getRange(2, valueCol, rows, 1).getValues();
      const names = assetSh.getRange(2, nameCol, rows, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        const val = parseNum_(vals[i][0]);
        const name = String(names[i][0] || "").trim();
        if (val > 0 && name !== "" && name !== "#N/A") {
          investTotal += val;
          assets.push({ name: name, value: val });
        }
      }
    }
  }

  // --- 計算當前總淨值 ---
  let currentTotalNetWorth = investTotal;
  if (detailSh) {
    const c = Number(detailSh.getRange("C2").getValue() || 0);
    const e = Number(detailSh.getRange("E2").getValue() || 0);
    const g = Number(detailSh.getRange("G2").getValue() || 0);
    const i = Number(detailSh.getRange("I2").getValue() || 0);
    currentTotalNetWorth += c + e + (g * freshUsdRate) - i;
  }

  // --- 處理歷史記錄 ---
  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  if (isManualUpdate === true && histSh) {
    const now = new Date();
    const lastRow = histSh.getLastRow();
    let isSameDay = false;
    if (lastRow >= 2) {
      const lastDate = histSh.getRange(lastRow, 1).getValue();
      if (lastDate instanceof Date && Utilities.formatDate(now, "GMT+8", "yyyyMMdd") === Utilities.formatDate(lastDate, "GMT+8", "yyyyMMdd")) {
        isSameDay = true;
      }
    }
    if (isSameDay) {
      histSh.getRange(lastRow, 2).setValue(currentTotalNetWorth);
    } else {
      histSh.appendRow([now, currentTotalNetWorth]);
    }
  }

  // --- 獲取地區分布數據 ---
  const regionSh = ss.getSheetByName("投資地區");
  let regions = [];
  if (regionSh && regionSh.getLastRow() >= 2) {
    regions = regionSh.getRange(2, 1, regionSh.getLastRow() - 1, 2).getValues()
      .map(r => ({ name: String(r[0] || "").trim(), value: parseNum_(r[1]) }))
      .filter(r => r.value > 0);
  }

  // --- 獲取折線圖數據 (最後30天) ---
  let history = [];
  if (histSh && histSh.getLastRow() >= 2) {
    history = histSh.getRange(2, 1, histSh.getLastRow() - 1, 2).getValues()
      .filter(r => r[0] && parseNum_(r[1]) > 0).slice(-30)
      .map(r => ({
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+8", "MM/dd") : String(r[0]),
        val: parseNum_(r[1])
      }));
  }

  // --- 獲取已實現損益摘要 ---
  const logSh = ss.getSheetByName(CONFIG.SHEET_LOGS);
  let realizedReturn = 0, realizedReturnTwd = 0;
  if (logSh) {
    const sumData = logSh.getRange("Y1:Z30").getValues();
    for(let row of sumData){
      const label = String(row[0]);
      if(label.includes("已實現總損益(TWD)")) realizedReturnTwd = parseNum_(row[1]);
      if(label.includes("已實現總損益(%)")) realizedReturn = (Number(String(row[1]).replace("%","")) || 0) * (String(row[1]).includes("%") ? 1 : 100);
    }
  }

  return { history, assets, regions, investTotal, usdRate: freshUsdRate, realizedReturn, realizedReturnTwd };
}

/* ================================
   4️⃣ 交易紀錄儲存邏輯
================================ */
function saveTrades(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.SHEET_LOGS);
    if (!sh) throw new Error("找不到買賣紀錄工作表");

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const startRow = findFirstEmptyRow_(sh);
    const rows = payload.trades.map((t, i) => buildFormulaRow_(headers, t, startRow + i, (name) => headers.indexOf(name)));
    
    sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    SpreadsheetApp.flush();
    return { ok: true, row: startRow };
  } finally { lock.releaseLock(); }
}

function parseNum_(val) {
  if (val === "" || val === null || val === undefined) return 0;
  if (typeof val === "number") return val;
  return Number(String(val).replace(/,/g, "")) || 0;
}

function findFirstEmptyRow_(sh) {
  const START = 86; // 根據您的需求從第 86 列開始找空位
  const lastRow = sh.getLastRow();
  if (lastRow < START) return START;
  const vals = sh.getRange(START, 1, lastRow - START + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) { if (!vals[i][0]) return START + i; }
  return lastRow + 1;
}

function buildFormulaRow_(headers, t, r, getCol) {
  const row = new Array(headers.length).fill("");
  const set = (n, v) => { const i = getCol(n); if (i !== -1) row[i] = v; };
  set("日期", t.date || new Date());
  set("交易類型", t.type);
  set("名稱", t.name);
  set("股票代號", t.symbol);
  set("買入價格", Number(t.price));
  set("買入股數", Number(t.qty));
  set("狀態", "已完成");
  set("價金(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), I${r}*J${r}, K${r}*L${r})`);
  set("應收付(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), P${r}-M${r}-N${r}, P${r}+M${r}+N${r})`);
  set("損益(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), Q${r}-O${r}, "")`);
  set("報酬率", `=IF(AND(ISNUMBER(R${r}), O${r}<>0), R${r}/O${r}, "")`);
  set("成本(TWD)", `=IF(O${r}<>"", O${r}*IF(H${r}="",1,H${r}), "")`);
  set("應收付(TWD)", `=Q${r}*IF(H${r}="",1,H${r})`);
  set("損益(TWD)", `=IF(R${r}<>"", R${r}*IF(H${r}="",1,H${r}), "")`);
  return row;
}

function forceAuth() { UrlFetchApp.fetch("https://www.google.com"); Logger.log("授權完成"); }