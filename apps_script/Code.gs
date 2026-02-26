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
   3️⃣ Dashboard 數據讀取 (V4.2 穩定強化版)
================================ */
function getDashboardData(inputs) {

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const detailSh = ss.getSheetByName(CONFIG.SHEET_DETAILS);
  let freshUsdRate = 32.2;

  /* ================================
     🟢 手動更新才同步市場資料
  =================================*/
  if (detailSh && inputs) {

    const startRow = 6;
    const lastRow = detailSh.getLastRow();

    if (lastRow >= startRow) {
      const symbols = detailSh.getRange(startRow, 2, lastRow - startRow + 1, 1).getValues();

      const priceResults = symbols.map(s => {
        const symbol = String(s[0] || "").trim();
        return [symbol ? Number(fetchYahooPrice(symbol)) : ""];
      });

      detailSh.getRange(startRow, 12, priceResults.length, 1).setValues(priceResults);
    }

    const fetchedRate = fetchYahooPrice("USDTWD=X");
    if (fetchedRate && !isNaN(fetchedRate)) {
      freshUsdRate = Number(fetchedRate);
    }

    detailSh.getRange("A2").setValue(freshUsdRate);

    if (inputs.cashTwd !== "") detailSh.getRange("C2").setValue(Number(inputs.cashTwd));
    if (inputs.settleTwd !== "") detailSh.getRange("E2").setValue(Number(inputs.settleTwd));
    if (inputs.cashUsd !== "") detailSh.getRange("G2").setValue(Number(inputs.cashUsd));
    if (inputs.loanTwd !== "") detailSh.getRange("I2").setValue(Number(inputs.loanTwd));

    SpreadsheetApp.flush();
  }

  /* ================================
     🔵 純讀取資料（AI 也會走這裡）
  =================================*/

  /* ===== 1️⃣ 資產統計(彙整) ===== */
  const assetSh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  let investTotal = 0;
  let assets = [];

  if (assetSh && assetSh.getLastRow() >= 2) {

    const headers = assetSh.getRange(1, 1, 1, assetSh.getLastColumn()).getValues()[0];

    // ⭐ 名稱欄雙重保護
    let nameCol = headers.indexOf("合併鍵(GroupKey)") + 1;
    if (nameCol <= 0) {
      nameCol = headers.indexOf("標的名稱") + 1;
    }
    if (nameCol <= 0) {
      throw new Error("資產統計(彙整)缺少名稱欄位");
    }

    const valueCol = headers.indexOf("市值(TWD)") + 1;
    const pnlCol   = headers.indexOf("損益(TWD)") + 1;
    const rateCol  = headers.indexOf("報酬率") + 1;

    if (valueCol > 0) {

      const rowCount = assetSh.getLastRow() - 1;
      const rows = assetSh.getRange(2, 1, rowCount, headers.length).getValues();

      for (let i = 0; i < rowCount; i++) {

        const name  = String(rows[i][nameCol - 1] || "").trim();
        const value = parseNum_(rows[i][valueCol - 1]);
        const pnl   = pnlCol > 0 ? parseNum_(rows[i][pnlCol - 1]) : 0;

        if (!name || value <= 0 || name === "#N/A") continue;

        investTotal += value;

        let returnRate = 0;

        // ✅ 優先使用 Sheet 現成報酬率
        if (rateCol > 0 && rows[i][rateCol - 1] !== "") {

          const rawRate = rows[i][rateCol - 1];

          if (typeof rawRate === "string" && rawRate.includes("%")) {
            returnRate = Number(rawRate.replace("%", "").replace(/,/g, ""));
          } else {
            returnRate = parseNum_(rawRate) * 100;
          }

        } else {
          // 🔁 fallback：用損益反推 (成本 = 市值 - 損益)
          const cost = value - pnl;
          if (cost !== 0) {
            returnRate = (pnl / cost) * 100;
          }
        }

        assets.push({
          name: name,
          value: value,
          returnRate: returnRate
        });
      }
    }
  }

  /* ===== 2️⃣ 淨值歷史 ===== */
  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  let history = [];

  if (histSh && histSh.getLastRow() >= 2) {
    history = histSh.getRange(2, 1, histSh.getLastRow() - 1, 2).getValues()
      .filter(r => r[0] && parseNum_(r[1]) > 0)
      .slice(-30)
      .map(r => ({
        date: r[0] instanceof Date
          ? Utilities.formatDate(r[0], "GMT+8", "MM/dd")
          : String(r[0]),
        val: parseNum_(r[1])
      }));
  }

  /* ===== 3️⃣ 投資地區 ===== */
  const regionSh = ss.getSheetByName(CONFIG.SHEET_REGIONS);
  let regions = [];

  if (regionSh && regionSh.getLastRow() >= 2) {
    regions = regionSh.getRange(2, 1, regionSh.getLastRow() - 1, 2).getValues()
      .map(r => ({
        name: String(r[0] || "").trim(),
        value: parseNum_(r[1])
      }))
      .filter(r => r.value > 0);
  }

  /* ===== 4️⃣ 已實現統計 ===== */
  const logSh = ss.getSheetByName(CONFIG.SHEET_LOGS);
  let realizedReturn = 0;
  let realizedReturnTwd = 0;

  if (logSh) {
    const summary = logSh.getRange("Y1:Z30").getDisplayValues();
    summary.forEach(row => {
      const label = String(row[0] || "");
      if (label.includes("已實現總損益(TWD)")) {
        realizedReturnTwd = parseNum_(row[1]);
      }
      if (label.includes("已實現總損益(%)")) {
        realizedReturn = Number(String(row[1]).replace("%", "").replace(/,/g, "")) || 0;
      }
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
   4️⃣ AI 助理分析 - V4.2 Balanced War Room
================================ */
function callGeminiAnalysis(userQuery) {

  if (!GEMINI_API_KEY) 
    return "⚠️ 翔翔，API Key 未設定。";

  const data = getDashboardData(null);

  /* ================================
     📊 風險排序引擎（純程式）
  ================================= */

  let processedAssets = [];
  let hasCriticalRisk = false;

  if (data.assets && data.assets.length > 0) {

    processedAssets = data.assets.map(a => {
      const rate = Number(a.returnRate) || 0;

      if (rate <= -10) hasCriticalRisk = true;

      return {
        name: a.name,
        value: a.value,
        rateNum: rate
      };
    });

    // 由低到高排序（風險優先）
    processedAssets.sort((a, b) => a.rateNum - b.rateNum);
  }

  const assetStr = processedAssets.map(a =>
    `${a.name}(市值:${Math.round(a.value/10000)}萬, 報酬:${a.rateNum.toFixed(2)}%)`
  ).join("、");

  /* ================================
     🧠 Prompt 設計
  ================================= */

  let systemPrompt = `
妳是翔翔的專屬金融管家，同時保留30%溫柔姊姊人格。
回答以理性分析為主（約70%），情緒陪伴為輔（約30%）。

回答規則：
1. 先給結論，再給觀察。
2. 若存在負報酬標的，優先說明。
3. 若存在超過20%報酬標的，提醒可能過熱。
4. 禁止冗長鋪陳。
5. 文字精準、冷靜。

【目前資產狀況】
總市值：${Math.round(data.investTotal).toLocaleString()} TWD
已實現損益：${Math.round(data.realizedReturnTwd).toLocaleString()} TWD
資產列表（已依風險排序）：
${assetStr}
`;

  if (hasCriticalRisk) {
    systemPrompt += `
⚠️ 系統提示：偵測到跌幅超過10%的標的，必須優先說明其風險。
`;
  }

  systemPrompt += `
【輸出限制】
- 必須稱呼「翔翔」
- 純文字
- 不得使用 Markdown
- 150字內
`;

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" + GEMINI_API_KEY;

  const payload = {
    contents: [{
      role: "user",
      parts: [{
        text: systemPrompt + "\n翔翔的問題：" + userQuery
      }]
    }]
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());

    if (json.error)
      return "翔翔，系統干擾：" + json.error.message;

    let reply = json.candidates?.[0]?.content?.parts?.[0]?.text
                || "翔翔，訊號短暫中斷。";

    return reply.replace(/[\*#_~`\[\]]/g, "").trim();

  } catch (e) {
    return "翔翔，伺服器波動，稍後再試。";
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