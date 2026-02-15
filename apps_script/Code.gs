/**
 * 投資戰情室 V6.56 - 穩定基底 + Gemini 2.0 Flash Lite (成本優化版)
 * 修正項目：
 * 1. 修正模型名稱：改為 gemini-2.0-flash-lite (符合您的 ListModels 查詢結果)。
 * 2. 成本優化：選用 Lite 系列模型，提供極高性價比且穩定的對話體驗。
 * 3. 維持所有核心記帳回寫邏輯 (A2, C2, E2, G2, I2)。
 */

// 🔥 唯一正確的 Gemini API Key
const GEMINI_API_KEY = "AIzaSyC5hvpL40X9uQ6pnhc1L9QPLbSFxR2AG58";

const CONFIG = {
  SPREADSHEET_ID: "1HM2MvZepqo1LVvgRoWwQ-1NmWKxo3ASAcXc2wECPgZU",
  SHEET_LOGS: "買賣紀錄_2026",
  SHEET_HISTORY: "淨值歷史",
  SHEET_ASSETS: "資產統計(彙整)",
  SHEET_REGIONS: "投資地區",
  SHEET_DETAILS: "庫存彙整(細項)" 
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
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
    } catch (e) {}
  }
  return HtmlService.createHtmlOutput("找不到網頁檔案，請確保檔案名稱為 ui");
}

/* ================================
   2️⃣ 手動更新市價 (Yahoo)
================================ */
function updateMarketData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_DETAILS);
  if (!sh) return;

  const headerRow = 5; 
  const lastRow = sh.getLastRow();
  if (lastRow <= headerRow) return;

  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  const symbolCol = headers.indexOf("Yahoo代號(Symbol)") + 1;
  const priceCol = headers.indexOf("目前市價") + 1;

  if (symbolCol <= 0 || priceCol <= 0) return;

  const data = sh.getRange(headerRow + 1, symbolCol, lastRow - headerRow, 1).getValues();
  const prices = data.map(row => {
    const symbol = String(row[0] || "").trim();
    return symbol ? [fetchYahooPrice(symbol)] : [""];
  });

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

  if (isManualUpdate === true) {
    try { updateMarketData(); } catch (e) {}
    const fetchedRate = fetchYahooPrice("USDTWD=X");
    if (fetchedRate && !isNaN(fetchedRate)) freshUsdRate = Number(fetchedRate);

    if (detailSh) {
      detailSh.getRange("A2").setValue(freshUsdRate);
      if (inputs) {
        if (inputs.cashTwd !== "") detailSh.getRange("C2").setValue(Number(inputs.cashTwd));
        if (inputs.settleTwd !== "") detailSh.getRange("E2").setValue(Number(inputs.settleTwd));
        if (inputs.cashUsd !== "") detailSh.getRange("G2").setValue(Number(inputs.cashUsd));
        if (inputs.loanTwd !== "") detailSh.getRange("I2").setValue(Number(inputs.loanTwd));
      }
    }
    SpreadsheetApp.flush();
  } else if (detailSh) {
    freshUsdRate = Number(detailSh.getRange("A2").getValue()) || 32.2;
  }

  const assetSh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  let investTotal = 0, assets = [];
  if (assetSh && assetSh.getLastRow() >= 2) {
    const headers = assetSh.getRange(1, 1, 1, assetSh.getLastColumn()).getValues()[0];
    const valueCol = headers.indexOf("市值(TWD)") + 1;
    let nameCol = headers.indexOf("合併鍵(GroupKey)") + 1 || headers.indexOf("標的名稱") + 1;

    if (valueCol > 0 && nameCol > 0) {
      const vals = assetSh.getRange(2, valueCol, assetSh.getLastRow() - 1, 1).getValues();
      const names = assetSh.getRange(2, nameCol, assetSh.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        const val = parseNum_(vals[i][0]);
        const name = String(names[i][0] || "").trim();
        if (val > 0 && name && name !== "#N/A") {
          investTotal += val;
          assets.push({ name: name, value: val });
        }
      }
    }
  }

  let currentTotalNetWorth = investTotal;
  if (detailSh) {
    currentTotalNetWorth += Number(detailSh.getRange("C2").getValue() || 0) +
                            Number(detailSh.getRange("E2").getValue() || 0) +
                            (Number(detailSh.getRange("G2").getValue() || 0) * freshUsdRate) -
                            Number(detailSh.getRange("I2").getValue() || 0);
  }

  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  if (isManualUpdate === true && histSh) {
    const now = new Date(), lastRow = histSh.getLastRow();
    let isSameDay = false;
    if (lastRow >= 2) {
      const lastDate = histSh.getRange(lastRow, 1).getValue();
      if (lastDate instanceof Date && Utilities.formatDate(now, "GMT+8", "yyyyMMdd") === Utilities.formatDate(lastDate, "GMT+8", "yyyyMMdd")) isSameDay = true;
    }
    if (isSameDay) histSh.getRange(lastRow, 2).setValue(currentTotalNetWorth);
    else histSh.appendRow([now, currentTotalNetWorth]);
  }

  let history = [];
  if (histSh && histSh.getLastRow() >= 2) {
    history = histSh.getRange(2, 1, histSh.getLastRow() - 1, 2).getValues()
      .filter(r => r[0] && parseNum_(r[1]) > 0).slice(-30)
      .map(r => ({ date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+8", "MM/dd") : String(r[0]), val: parseNum_(r[1]) }));
  }

  const regionSh = ss.getSheetByName(CONFIG.SHEET_REGIONS);
  let regions = [];
  if (regionSh && regionSh.getLastRow() >= 2) {
    regions = regionSh.getRange(2, 1, regionSh.getLastRow() - 1, 2).getValues()
      .map(r => ({ name: String(r[0] || "").trim(), value: parseNum_(r[1]) })).filter(r => r.value > 0);
  }

  const logSh = ss.getSheetByName(CONFIG.SHEET_LOGS);
  let realizedReturn = 0, realizedReturnTwd = 0;
  if (logSh) {
    const summary = logSh.getRange("Y1:Z30").getValues();
    summary.forEach(row => {
      const label = String(row[0]);
      if (label.includes("已實現總損益(TWD)")) realizedReturnTwd = parseNum_(row[1]);
      if (label.includes("已實現總損益(%)")) realizedReturn = (Number(String(row[1]).replace("%","")) || 0) * (String(row[1]).includes("%") ? 1 : 100);
    });
  }

  return { history, assets, regions, investTotal, usdRate: freshUsdRate, realizedReturn, realizedReturnTwd };
}

/* ================================
   4️⃣ 對話式咪咪：AI 分析邏輯 (Gemini 2.0 Flash Lite)
================================ */
function callGeminiAnalysis(userQuery) {
  if (!GEMINI_API_KEY) return "⚠️ 請先在 Code.gs 中設定 API Key";

  // 取得最新資產數據
  const data = getDashboardData(null, false);
  const assetStr = data.assets.map(a => `${a.name}(${Math.round(a.value/10000)}萬)`).join("、");
  
  const prompt = `
    你是一位專業、毒舌但熱心的私人財富顧問「咪咪」。
    總市值：${Math.round(data.investTotal).toLocaleString()} TWD
    已實現損益：${Math.round(data.realizedReturnTwd).toLocaleString()} TWD
    主要持倉：${assetStr}
    即時匯率：${data.usdRate}
    主人問題：${userQuery}
    回答150字內，幽默直接。直接回文字，不要使用 Markdown。
  `;

  /**
   * 🔥 修改點：更換為 gemini-2.0-flash-lite
   * 使用 v1beta 正確路徑，提供穩定且大量的低成本服務。
   */
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" + GEMINI_API_KEY;
  
  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }]
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const json = JSON.parse(response.getContentText());
    
    // 錯誤診斷
    if (json.error) return "AI 錯誤: " + json.error.message;
    
    return json.candidates?.[0]?.content?.parts?.[0]?.text || "咪咪今天罷工中 😼";
  } catch (e) {
    return "連線失敗：" + e.message;
  }
}

function parseNum_(val) {
  if (val === "" || val === null || val === undefined) return 0;
  if (typeof val === "number") return val;
  return Number(String(val).replace(/,/g, "")) || 0;
}

function saveTrades(p) { return { ok: true }; } // 預留擴充