/**
 * 投資戰情室 V6.60 - 安全保護版
 * 修正項目：
 * 1. 隱藏 API Key：使用 PropertiesService 安全讀取，防止再次被判定為洩漏。
 * 2. 指令碼屬性教學：請至「專案設定 (⚙️)」->「指令碼屬性」新增名為 GEMINI_API_KEY 的屬性。
 * 3. 維持所有 6.59 版修正：包含成本欄位對接、買賣分流與 AI 穩定邏輯。
 */

// 🔒 安全讀取方式：不再將明文 KEY 寫在這裡
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

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
  return HtmlService.createHtmlOutputFromFile("ui")
    .setTitle("投資戰情室")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ================================
   2️⃣ 交易寫入核心
================================ */
function saveTrades(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.SHEET_LOGS);
    if (!sh) throw new Error("找不到買賣紀錄分頁");

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const getCol = (name) => headers.indexOf(name);
    
    const startRow = findFirstEmptyRow_(sh);
    const rows = payload.trades.map((t, i) => buildFormulaRow_(headers, payload.defaults || {}, t, startRow + i, getCol));
    
    sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    SpreadsheetApp.flush();
    
    return { ok: true, row: startRow };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function buildFormulaRow_(headers, defaults, t, r, getCol) {
  const row = new Array(headers.length).fill("");
  const setVal = (name, val) => { 
    const idx = getCol(name); 
    if (idx !== -1) row[idx] = val; 
  };

  setVal("日期", t.date || new Date());
  setVal("交易類型", t.type); 
  setVal("平台", defaults.platform || "");
  setVal("帳戶類型", defaults.account || "");
  
  let currency = "TWD";
  if (defaults.account && defaults.account.toUpperCase().includes("USD")) {
    currency = "USD";
  }
  setVal("幣別", currency);

  setVal("名稱", t.name);
  setVal("股票代號", t.symbol);

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

  setVal("手續費", Number(t.fee || 0));
  setVal("交易稅", Number(t.tax || 0));
  
  if (t.cost !== "" && t.cost !== null && t.cost !== undefined) {
    setVal("成本(原幣)※賣出需填", Number(t.cost));
  }

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

function findFirstEmptyRow_(sh) {
  const START_ROW = 86; 
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
function getDashboardData(inputs) {

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const detailSh = ss.getSheetByName(CONFIG.SHEET_DETAILS);

  let freshUsdRate = 32.2;

  if (detailSh) {

    // 每次都抓匯率（避免卡住）
    const fetchedRate = fetchYahooPrice("USDTWD=X");
    if (fetchedRate && !isNaN(fetchedRate)) {
      freshUsdRate = Number(fetchedRate);
    }

    detailSh.getRange("A2").setValue(freshUsdRate);

    if (inputs) {
      if (inputs.cashTwd !== "")
        detailSh.getRange("C2").setValue(Number(inputs.cashTwd));
      if (inputs.settleTwd !== "")
        detailSh.getRange("E2").setValue(Number(inputs.settleTwd));
      if (inputs.cashUsd !== "")
        detailSh.getRange("G2").setValue(Number(inputs.cashUsd));
      if (inputs.loanTwd !== "")
        detailSh.getRange("I2").setValue(Number(inputs.loanTwd));
    }
  }

  SpreadsheetApp.flush();

  /* ===== 以下維持不變 ===== */

  const assetSh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  let investTotal = 0, assets = [];

  if (assetSh && assetSh.getLastRow() >= 2) {

    const headers = assetSh.getRange(1, 1, 1, assetSh.getLastColumn())
      .getValues()[0];

    const valueCol = headers.indexOf("市值(TWD)") + 1;
    let nameCol = headers.indexOf("合併鍵(GroupKey)") + 1;

    if (nameCol <= 0)
      nameCol = headers.indexOf("標的名稱") + 1;

    if (valueCol > 0 && nameCol > 0) {

      const vals = assetSh.getRange(
        2, valueCol,
        assetSh.getLastRow() - 1, 1
      ).getValues();

      const names = assetSh.getRange(
        2, nameCol,
        assetSh.getLastRow() - 1, 1
      ).getValues();

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

  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  let history = [];

  if (histSh && histSh.getLastRow() >= 2) {
    history = histSh.getRange(
      2, 1,
      histSh.getLastRow() - 1, 2
    ).getValues()
      .filter(r => r[0] && parseNum_(r[1]) > 0)
      .slice(-30)
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
    regions = regionSh.getRange(
      2, 1,
      regionSh.getLastRow() - 1, 2
    ).getValues()
      .map(r => ({
        name: String(r[0] || "").trim(),
        value: parseNum_(r[1])
      }))
      .filter(r => r.value > 0);
  }

  const logSh = ss.getSheetByName(CONFIG.SHEET_LOGS);
  let realizedReturn = 0, realizedReturnTwd = 0;

  if (logSh) {
    const summary = logSh.getRange("Y1:Z30").getValues();
    summary.forEach(row => {
      const label = String(row[0]);
      if (label.includes("已實現總損益(TWD)"))
        realizedReturnTwd = parseNum_(row[1]);
      if (label.includes("已實現總損益(%)"))
        realizedReturn =
          (Number(String(row[1]).replace("%", "")) || 0);
    });
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
   4️⃣ AI 助理分析 (Gemini 2.0 Flash Lite)
================================ */
function callGeminiAnalysis(userQuery) {
  if (!GEMINI_API_KEY) return "⚠️ 請先在專案設定中設定 GEMINI_API_KEY 屬性";

  const data = getDashboardData(null);
  const assetStr = data.assets.map(a => `${a.name}(${Math.round(a.value/10000)}萬)`).join("、");
  
  const prompt = `你是一位專業私人財富顧問「咪咪」。總市值：${Math.round(data.investTotal).toLocaleString()} TWD，已實現損益：${Math.round(data.realizedReturnTwd).toLocaleString()} TWD，持倉：${assetStr}。回答主人問題：${userQuery}。回答150字內，幽默直接。直接回文字。`;
  
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" + GEMINI_API_KEY;
  
  try {
    const response = UrlFetchApp.fetch(url, { 
      method: 'post', 
      contentType: 'application/json', 
      payload: JSON.stringify({ contents: [{ role: "user", parts: [{ text: prompt }] }] }),
      muteHttpExceptions: true 
    });
    
    const json = JSON.parse(response.getContentText());
    if (json.error) return "AI 錯誤: " + json.error.message;
    return json.candidates?.[0]?.content?.parts?.[0]?.text || "咪咪今天不想說話 😼";
  } catch (e) { 
    return "連線失敗：" + e.message; 
  }
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