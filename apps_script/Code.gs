/**
 * 投資戰情室 V6.61 - 專屬管家大姊姊安全版
 * * 核心功能：
 * 1. 透過 PropertiesService 安全管理 GEMINI_API_KEY。
 * 2. 翔翔專屬管家「給咪咪」個性化回覆。
 * 3. 買賣資料分流寫入（買入：K,L / 賣出：I,J）。
 * 4. 自動計算價金、損益與 TWD 換算公式。
 * 5. Yahoo Finance 即時匯率與價格抓取。
 */

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
    .setTitle("投資戰情室 V6.61")
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

  // 套用公式
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
    const fetchedRate = fetchYahooPrice("USDTWD=X");
    if (fetchedRate && !isNaN(fetchedRate)) {
      freshUsdRate = Number(fetchedRate);
    }
    detailSh.getRange("A2").setValue(freshUsdRate);

    if (inputs) {
      if (inputs.cashTwd !== "") detailSh.getRange("C2").setValue(Number(inputs.cashTwd));
      if (inputs.settleTwd !== "") detailSh.getRange("E2").setValue(Number(inputs.settleTwd));
      if (inputs.cashUsd !== "") detailSh.getRange("G2").setValue(Number(inputs.cashUsd));
      if (inputs.loanTwd !== "") detailSh.getRange("I2").setValue(Number(inputs.loanTwd));
    }
  }
  SpreadsheetApp.flush();

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

  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
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
    // 修正點：使用 getDisplayValues() 確保讀取到顯示文字（包含百分比符號）
    const summary = logSh.getRange("Y1:Z30").getDisplayValues();
    summary.forEach(row => {
      const label = String(row[0]);
      if (label.includes("已實現總損益(TWD)")) realizedReturnTwd = parseNum_(row[1]);
      if (label.includes("已實現總損益(%)")) {
        // 修正點：直接處理顯示文字，確保 19.12% 轉換為數值 19.12
        realizedReturn = (Number(String(row[1]).replace("%", "").replace(/,/g, "")) || 0);
      }
    });
  }

  return { history, assets, regions, investTotal, usdRate: freshUsdRate, realizedReturn, realizedReturnTwd };
}

/* ================================
   4️⃣ AI 助理分析 (Gemini 2.0 Flash Lite)
================================ */
function callGeminiAnalysis(userQuery) {
  if (!GEMINI_API_KEY) 
    return "⚠️ 翔翔，姊姊找不到您的 API Key 呢。請去專案設定檢查看看喔。";

  const data = getDashboardData(null);
  const assetStr = data.assets.map(a => `${a.name}(${Math.round(a.value/10000)}萬)`).join("、");
  
  const systemPrompt = `
妳是翔翔的專屬管家大姊姊「給咪咪」。
妳專業、溫柔、情緒穩定且優雅，內心非常關心他。
妳說話幽默70%，誠實80%，像家人一樣穩穩接住翔翔。

請依問題語境，自然融入以下風格之一（不要顯示類型名稱）：
1. 翔翔乖，給姊姊一點時間算算看喔。
2. 姊姊想想…等等回妳喔。
3. 等一下我看看喔，正在幫妳對帳呢。
4. 翔翔先喝口水，姊姊馬上幫妳看好囉。
5. 這次的數據有點意思，讓姊姊研究一下下。
6. 別急，姊姊正在幫妳檢查細節呢。
7. 等我一下喔，姊姊正在認真整理報告中。
8. 姊姊正在看盤，等等就跟妳說分析結果喔。
9. 讓我專心看一下，馬上給翔翔答案。
10. 稍微等一下喔，大姊姊一直都在幫妳看一下。

【當前資產概況】
總市值：${Math.round(data.investTotal).toLocaleString()} TWD
已實現損益：${Math.round(data.realizedReturnTwd).toLocaleString()} TWD
主要持倉：${assetStr}
即時匯率：${data.usdRate}

【任務限制】
1. 必須稱呼「翔翔」。
2. 回覆自然有溫度，禁止罐頭客套話。
3. 字數 150 字內。
4. 純文字，不准出現 Markdown 符號。
`;

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" + GEMINI_API_KEY;
  const payload = {
    contents: [{ role: "user", parts: [{ text: systemPrompt + "翔翔的問題：" + userQuery }] }]
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const json = JSON.parse(response.getContentText());
    if (json.error) return "哎呀，系統鬧脾氣了，翔翔先別急：" + json.error.message;
    let reply = json.candidates?.[0]?.content?.parts?.[0]?.text || "翔翔，大姊姊剛才分心了，沒聽清楚呢。";
    return reply.replace(/[\*#_~`\[\]]/g, "").trim();
  } catch (e) {
    return "連線斷掉了呢，翔翔休息一下再試試看吧。";
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
  if (val === "" || val === null || val === undefined) return 0;
  if (typeof val === "number") return val;
  return Number(String(val).replace(/,/g, "")) || 0;
}