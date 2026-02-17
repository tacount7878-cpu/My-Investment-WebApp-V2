/**
 * 投資戰情室 V6.58 - 最終修正版
 * 修正項目：
 * 1. 修正買賣欄位：根據交易類型自動判斷寫入「買入」或「賣出」欄位，不再錯位。
 * 2. 補全缺失欄位：正確寫入「平台」、「帳戶類型」、「幣別」。
 * 3. 幣別邏輯：根據帳戶類型關鍵字自動判斷 USD/TWD。
 * 4. 修正成本欄位：精準對接「成本(原幣)※賣出需填」標題，解決寫入空白問題。
 * 5. 寫入位置：從第 86 列開始尋找第一個空白列。
 * 6. AI 助理：整合 Gemini 2.0 Flash Lite 提供資產分析。
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
   1️⃣ 網頁入口
================================ */
function doGet() {
  // 嘗試從名為 "ui" 的 HTML 檔案建立輸出
  return HtmlService.createHtmlOutputFromFile("ui")
    .setTitle("投資戰情室")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ================================
   2️⃣ 交易寫入核心
================================ */
function saveTrades(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 鎖定 30 秒防止寫入衝突
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.SHEET_LOGS);
    if (!sh) throw new Error("找不到買賣紀錄分頁");

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const getCol = (name) => headers.indexOf(name);
    
    // 找出第一行空白列 (從第 86 列開始找)
    const startRow = findFirstEmptyRow_(sh);
    
    // 建立要寫入的資料行
    const rows = payload.trades.map((t, i) => buildFormulaRow_(headers, payload.defaults || {}, t, startRow + i, getCol));
    
    // 寫入試算表
    sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    SpreadsheetApp.flush();
    
    return { ok: true, row: startRow };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 建立整行資料，精準對接 A-R 欄位
 */
function buildFormulaRow_(headers, defaults, t, r, getCol) {
  const row = new Array(headers.length).fill("");
  const setVal = (name, val) => { 
    const idx = getCol(name); 
    if (idx !== -1) row[idx] = val; 
  };

  // --- A-E 欄：基礎設定 ---
  setVal("日期", t.date || new Date());
  setVal("交易類型", t.type); // 「買入」或「賣出」
  setVal("平台", defaults.platform || "");
  setVal("帳戶類型", defaults.account || "");
  
  // 幣別自動判斷：若帳戶名稱含 USD 則填 USD，否則 TWD
  let currency = "TWD";
  if (defaults.account && defaults.account.toUpperCase().includes("USD")) {
    currency = "USD";
  }
  setVal("幣別", currency);

  // --- F-G 欄：標的資訊 ---
  setVal("名稱", t.name);
  setVal("股票代號", t.symbol);

  // --- I-L 欄：買賣價格分流 ---
  if (t.type.includes("買")) {
    setVal("買入價格", Number(t.price || 0));
    setVal("買入股數", Number(t.qty || 0));
    setVal("賣出價格", ""); 
    setVal("賣出股數", "");
  } else {
    setVal("賣出價格", Number(t.price || 0));
    setVal("賣出股數", Number(t.qty || 0));
    setVal("買入價格", ""); 
    setVal("買入股數", "");
  }

  // --- M-N 欄：費用 ---
  setVal("手續費", Number(t.fee || 0));
  setVal("交易稅", Number(t.tax || 0));
  
  // --- O 欄：成本 (修正點：對接完整標題名稱) ---
  if (t.cost !== "" && t.cost !== null && t.cost !== undefined) {
    setVal("成本(原幣)※賣出需填", Number(t.cost));
  }

  setVal("狀態", "已完成");

  // --- 帶入試算表公式欄位 (P, Q, R 等) ---
  setVal("價金(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), I${r}*J${r}, K${r}*L${r})`);
  setVal("應收付(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), P${r}-M${r}-N${r}, P${r}+M${r}+N${r})`);
  setVal("損益(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), Q${r}-O${r}, "")`);
  setVal("報酬率", `=IF(AND(ISNUMBER(R${r}), O${r}<>0), R${r}/O${r}, "")`);
  
  // 台幣轉換公式 (假設匯率在 H 欄)
  setVal("成本(TWD)", `=IF(O${r}<>"", O${r}*IF(H${r}="",1,H${r}), "")`);
  setVal("應收付(TWD)", `=Q${r}*IF(H${r}="",1,H${r})`);
  setVal("損益(TWD)", `=IF(R${r}<>"", R${r}*IF(H${r}="",1,H${r}), "")`);

  return row;
}

function findFirstEmptyRow_(sh) {
  const START_ROW = 86; // 從第 86 列開始找空白
  const lastRow = sh.getLastRow();
  if (lastRow < START_ROW) return START_ROW;
  const values = sh.getRange(START_ROW, 1, Math.max(1, lastRow - START_ROW + 1), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) return START_ROW + i;
  }
  return lastRow + 1;
}

/* ================================
   3️⃣ Dashboard 數據讀取
================================ */
function getDashboardData(inputs, isManualUpdate) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const detailSh = ss.getSheetByName(CONFIG.SHEET_DETAILS);
  let freshUsdRate = 32.2; 

  if (isManualUpdate === true) {
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

  // 讀取資產佔比
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

  // 讀取歷史數據
  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  let history = [];
  if (histSh && histSh.getLastRow() >= 2) {
    history = histSh.getRange(2, 1, histSh.getLastRow() - 1, 2).getValues()
      .filter(r => r[0] && parseNum_(r[1]) > 0).slice(-30)
      .map(r => ({ date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+8", "MM/dd") : String(r[0]), val: parseNum_(r[1]) }));
  }

  // 讀取地區分佈
  const regionSh = ss.getSheetByName(CONFIG.SHEET_REGIONS);
  let regions = [];
  if (regionSh && regionSh.getLastRow() >= 2) {
    regions = regionSh.getRange(2, 1, regionSh.getLastRow() - 1, 2).getValues()
      .map(r => ({ name: String(r[0] || "").trim(), value: parseNum_(r[1]) })).filter(r => r.value > 0);
  }

  // 讀取摘要數據
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
   4️⃣ AI 助理分析 (Gemini 2.0 Flash Lite)
================================ */
function callGeminiAnalysis(userQuery) {
  const data = getDashboardData(null, false);
  const assetStr = data.assets.map(a => `${a.name}(${Math.round(a.value/10000)}萬)`).join("、");
  const prompt = `你是一位專業私人財富顧問「咪咪」。總市值：${Math.round(data.investTotal).toLocaleString()} TWD，持倉：${assetStr}。回答主人問題：${userQuery}。回答150字內，幽默直接。直接回文字。`;
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" + GEMINI_API_KEY;
  try {
    const response = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ contents: [{ role: "user", parts: [{ text: prompt }] }] }) });
    return JSON.parse(response.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text || "咪咪今天不想說話 😼";
  } catch (e) { return "連線失敗：" + e.message; }
}

function fetchYahooPrice(symbol) {
  try {
    const res = UrlFetchApp.fetch(`https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?interval=1d`, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    return json.chart?.result?.[0]?.meta?.regularMarketPrice || "";
  } catch (e) { return ""; }
}

function parseNum_(val) {
  if (!val) return 0;
  if (typeof val === "number") return val;
  return Number(String(val).replace(/,/g, "")) || 0;
}