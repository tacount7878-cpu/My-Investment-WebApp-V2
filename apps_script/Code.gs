/**
 * 投資戰情室 V6.30 - 全功能終極整合版
 * 整合了：
 * 1. Yahoo 即時股價更新 (updateMarketData)
 * 2. UI 現金/貸款數據回寫 (getDashboardData -> SHEET_DETAILS)
 * 3. 總淨值歷史自動記錄 (getDashboardData -> SHEET_HISTORY，同日覆蓋/異日新增)
 */

const CONFIG = {
  SPREADSHEET_ID: "1HM2MvZepqo1LVvgRoWwQ-1NmWKxo3ASAcXc2wECPgZU",
  SHEET_LOGS: "買賣紀錄_2026",
  SHEET_HISTORY: "淨值歷史",
  SHEET_ASSETS: "資產統計(彙整)",
  SHEET_REGIONS: "投資地區",
  SHEET_DETAILS: "庫存彙整(細項)" // ★ 寫入目標 1
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
  const possibleNames = ["ui", "ui.html", "Index", "apps_script/ui"];
  for (let name of possibleNames) {
    try {
      return HtmlService.createHtmlOutputFromFile(name)
        .setTitle("投資戰情室")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (e) {}
  }
  return HtmlService.createHtmlOutput("找不到網頁檔案");
}

/* ================================
   2️⃣ 手動更新市價 (Yahoo) -> 更新資產統計表
================================ */
function updateMarketData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  if (!sh) return;

  // 根據截圖，標題在 Row 1
  let headerRow = 1; 
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let symbolCol = headers.indexOf("Yahoo代號(Symbol)") + 1;

  // 容錯：如果 Row 1 找不到，試試 Row 5
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
    Utilities.sleep(10);
  }

  sh.getRange(startRow, priceCol, prices.length, 1).setValues(prices);
}

/* ================================
   3️⃣ Yahoo 抓價 & 匯率工具
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
   這裡包含了您要的所有功能！
================================ */
function getDashboardData(inputs) {
  // 1. 先去更新資產表的股價
  try { updateMarketData(); } catch (e) {}

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // 2. ★ 寫入匯率與現金貸款到「庫存彙整(細項)」
  const detailSh = ss.getSheetByName(CONFIG.SHEET_DETAILS);
  let freshUsdRate = 32.2; 

  if (detailSh) {
    // 2.1 抓即時匯率
    const fetchedRate = fetchYahooPrice("USDTWD=X");
    if (fetchedRate && !isNaN(fetchedRate)) freshUsdRate = Number(fetchedRate);

    // 2.2 寫入 A2 (匯率)
    detailSh.getRange("A2").setValue(freshUsdRate);

    // 2.3 寫入 UI 傳來的數據 (C2, E2, G2, I2)
    // 這裡就是您指定的覆蓋功能
    if (inputs) {
      if (inputs.cashTwd !== "") detailSh.getRange("C2").setValue(Number(inputs.cashTwd));
      if (inputs.settleTwd !== "") detailSh.getRange("E2").setValue(Number(inputs.settleTwd));
      if (inputs.cashUsd !== "") detailSh.getRange("G2").setValue(Number(inputs.cashUsd));
      if (inputs.loanTwd !== "") detailSh.getRange("I2").setValue(Number(inputs.loanTwd));
    }
  }

  // 3. 強制刷新計算 (讓公式吃到剛寫入的股價和現金)
  SpreadsheetApp.flush();

  /* ===== 以下為讀取與計算 ===== */

  // 4. 讀取並計算「投資部位總市值」
  const assetSh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  let investTotal = 0;
  let assets = [];
  
  if (assetSh && assetSh.getLastRow() >= 2) {
    // 修正：讀取 Row 1 標題
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

  // 5. 計算當下「資產總淨值」
  // 公式：投資總值 + 台幣現金 + 交割現金 + (美金 * 匯率) - 貸款
  let currentTotalNetWorth = investTotal;
  if (inputs) {
    currentTotalNetWorth += Number(inputs.cashTwd || 0);
    currentTotalNetWorth += Number(inputs.settleTwd || 0);
    currentTotalNetWorth += (Number(inputs.cashUsd || 0) * freshUsdRate);
    currentTotalNetWorth -= Number(inputs.loanTwd || 0);
  }

  // 6. ★ 寫入淨值歷史 (同日覆蓋，異日新增)
  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  if (histSh) {
    const now = new Date();
    const lastRow = histSh.getLastRow();
    
    let isSameDay = false;
    // 檢查最後一筆資料的日期
    if (lastRow >= 2) {
      const lastDateVal = histSh.getRange(lastRow, 1).getValue();
      if (lastDateVal instanceof Date) {
        const todayStr = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
        const lastDateStr = Utilities.formatDate(lastDateVal, "GMT+8", "yyyyMMdd");
        if (todayStr === lastDateStr) {
          isSameDay = true;
        }
      }
    }

    if (isSameDay) {
      // 同一天 -> 覆蓋最後一行 (更新時間與淨值)
      histSh.getRange(lastRow, 1).setValue(now);
      histSh.getRange(lastRow, 2).setValue(currentTotalNetWorth);
    } else {
      // 不同天 -> 新增一行
      // 如果表格是空的，補標題
      if (lastRow < 2 && histSh.getRange(1,1).getValue() === "") {
         histSh.appendRow(["時間", "資產總淨值(TWD)"]); 
      }
      histSh.appendRow([now, currentTotalNetWorth]);
    }
  }

  // 7. 讀取歷史數據 (用於前端畫圖)
  let history = [];
  if (histSh && histSh.getLastRow() >= 2) {
    const histData = histSh.getRange(Math.max(2, histSh.getLastRow() - 29), 1, 30, 2).getValues();
    history = histData.filter(r => r[0] && parseNum_(r[1]) > 0).map(r => ({
      date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+8", "MM/dd") : String(r[0]),
      val: parseNum_(r[1])
    }));
  }

  // 8. 投資地區
  const regionSh = ss.getSheetByName(CONFIG.SHEET_REGIONS);
  let regions = [];
  if (regionSh && regionSh.getLastRow() >= 2) {
    const regionData = regionSh.getRange(2, 1, regionSh.getLastRow() - 1, 2).getValues();
    regions = regionData.map(r => ({ name: String(r[0] || "").trim(), value: parseNum_(r[1]) })).filter(r => r.value > 0);
  }

  // 9. 摘要數據 (報酬率等)
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
  setVal("價金(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), I${r}*J${r}, K${r}*L${r})`);
  setVal("應收付(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), P${r}-M${r}-N${r}, P${r}+M${r}+N${r})`);
  setVal("損益(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), Q${r}-O${r}, "")`);
  setVal("報酬率", `=IF(AND(ISNUMBER(R${r}), O${r}<>0), R${r}/O${r}, "")`);
  setVal("成本(TWD)", `=IF(O${r}<>"", O${r}*IF(H${r}="",1,H${r}), "")`);
  setVal("應收付(TWD)", `=Q${r}*IF(H${r}="",1,H${r})`);
  setVal("損益(TWD)", `=IF(R${r}<>"", R${r}*IF(H${r}="",1,H${r}), "")`);
  return row;
}