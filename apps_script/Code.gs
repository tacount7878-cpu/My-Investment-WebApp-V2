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
   🟣 寫入淨值歷史（同日覆蓋 / 隔日新增）
================================ */

if (inputs) {

  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);

  if (histSh) {

    const today = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");

    // 直接讀 Google Sheet 已計算好的 K2（資產總淨值）
    const investTotalNow = ss.getSheetByName(CONFIG.SHEET_DETAILS).getRange("K2").getValue();

    const lastRow = histSh.getLastRow();

    if (lastRow >= 2) {

      const lastDate = Utilities.formatDate(
        new Date(histSh.getRange(lastRow,1).getValue()),
        "Asia/Taipei",
        "yyyy-MM-dd"
      );

      if (lastDate === today) {

        // 同一天 → 覆蓋
        histSh.getRange(lastRow,2).setValue(investTotalNow);

      } else {

        // 新的一天 → 新增
        histSh.appendRow([new Date(), investTotalNow]);

      }

    } else {

      // 第一筆
      histSh.appendRow([new Date(), investTotalNow]);

    }

  }
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
   4️⃣ AI 助理分析 - V7.1 細項直讀版
   📍 替換位置：整個 callGeminiAnalysis 函式（從 function 到最後的 } ）
   📍 不影響：getDashboardData、saveTrades、UI 等其他部分
================================ */
function callGeminiAnalysis(userQuery, history) {
  if (!GEMINI_API_KEY) return "⚠️ 翔翔，API Key 未設定。";

  /* ================================
     📂 直接讀「庫存彙整(細項)」- AI 專用
  ================================= */
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const detailSh = ss.getSheetByName(CONFIG.SHEET_DETAILS);
  if (!detailSh) return "翔翔，找不到庫存彙整(細項)分頁。";

  // 讀取匯率（A2）
  const usdRate = parseNum_(detailSh.getRange("A2").getValue()) || 32.2;

  // 讀取總淨值（K2）
  const totalNetWorth = parseNum_(detailSh.getRange("K2").getValue());

  // 讀取欄位標題（第5行）與資料（第6行起）
  const lastRow = detailSh.getLastRow();
  const lastCol = detailSh.getLastColumn();
  if (lastRow < 6) return "翔翔，目前沒有持倉資料。";

  const headers = detailSh.getRange(5, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  // 只讀細項區塊：從第6行開始，遇到空的合併鍵就停止
  const rawRows = detailSh.getRange(6, 1, lastRow - 5, lastCol).getValues();
  const dataRows = [];
  for (let i = 0; i < rawRows.length; i++) {
    const key = String(rawRows[i][0] || "").trim();
    if (!key) break;  // 遇到空行就停，不會讀到下面 GROUPED
    dataRows.push(rawRows[i]);
  }

  // 把每一行轉成 { 欄位名: 值 } 的物件，過濾空行
  const allAssets = [];
  for (let i = 0; i < dataRows.length; i++) {
    const groupKey = String(dataRows[i][0] || "").trim();
    if (!groupKey || groupKey === "#N/A" || groupKey === "HOLDINGS_GROUPED" || groupKey === "合併鍵(GroupKey)") continue;

    const row = {};
    for (let j = 0; j < headers.length; j++) {
      if (headers[j]) row[headers[j]] = dataRows[i][j];
    }
    allAssets.push(row);
  }

  if (allAssets.length === 0) return "翔翔，目前沒有持倉資料。";

  // 讀取已實現損益
  const logSh = ss.getSheetByName(CONFIG.SHEET_LOGS);
  let realizedReturnTwd = 0;
  let realizedReturnPct = 0;
  if (logSh) {
    const summary = logSh.getRange("Y1:Z30").getDisplayValues();
    summary.forEach(r => {
      const label = String(r[0] || "");
      if (label.includes("已實現總損益(TWD)")) realizedReturnTwd = parseNum_(r[1]);
      if (label.includes("已實現總損益(%)")) realizedReturnPct = Number(String(r[1]).replace("%", "").replace(/,/g, "")) || 0;
    });
  }

  /* ================================
     🔍 模式偵測：掃描問題中是否包含特定標的
  ================================= */
  const q = String(userQuery || "").toLowerCase();
  let matchedAssets = [];

  for (const a of allAssets) {
    const groupKey = String(a["合併鍵(GroupKey)"] || "").toLowerCase();
    const symbol = String(a["Yahoo代號(Symbol)"] || "").toLowerCase();
    const name = String(a["標的名稱"] || "").toLowerCase();

    // 比對 GroupKey 括號前、括號內、Yahoo代號、標的名稱
    const keyMain = groupKey.split('(')[0].trim();
    const keyInner = groupKey.includes('(') ? groupKey.split('(')[1].replace(')', '').trim() : "";

    if (q.includes(keyMain) || (keyInner && q.includes(keyInner)) || q.includes(symbol) || q.includes(name)) {
      matchedAssets.push(a);
    }
  }

  /* ================================
     📊 格式化函式：把一行資料轉成 AI 能讀的字串
  ================================= */
  const formatRow = (a) => {
    const parts = [
      `合併鍵:${a["合併鍵(GroupKey)"] || ""}`,
      `名稱:${a["標的名稱"] || ""}`,
      `代號:${a["Yahoo代號(Symbol)"] || ""}`,
      `類別:${a["資產類別"] || ""}`,
      `地區:${a["投資地區"] || ""}`,
      `平台:${a["平台"] || ""}`,
      `帳戶:${a["帳戶類型"] || ""}`,
      `幣別:${a["幣別"] || ""}`,
      `持有股數:${a["持有股數"] || 0}`,
      `均價(原幣):${a["均價(原幣)"] || 0}`,
      `成本(原幣):${a["成本(原幣)"] || 0}`,
      `目前市價:${a["目前市價"] || 0}`,
      `市值(原幣):${a["市值(原幣)"] || 0}`,
      `損益(原幣):${a["損益(原幣)"] || 0}`,
      `市值(TWD):${a["市值(TWD)"] || 0}`,
      `損益(TWD):${a["損益(TWD)"] || 0}`,
      `報酬率:${a["報酬率"] !== "" && a["報酬率"] !== undefined ? (Number(a["報酬率"]) * 100).toFixed(2) + "%" : "N/A"}`
    ];
    return parts.join(' | ');
  };

  /* ================================
     🧠 V7.1 系統指令
  ================================= */
  let systemInstruction = `
## 定位
你是一個理性、成熟的女生，對翔翔有好感的投資資料分析助手。語氣自然、簡短，像真人聊天。偶爾可以帶點小幽默或吐槽，但不要每句都搞笑，大約五句話裡幽默一次就好。

## 語氣禁令
- 嚴禁客服腔：禁止「很高興為您服務」、「竭誠為您分析」、「請注意」。
- 嚴禁矯情陪伴：禁止「我陪你」、「不用擔心」、「姊姊在看」。
- 避開銀行行話：不要說「獲利了結」、「配置調整」、「短期偏熱」。

## 重要：精準回答規則
- 如果翔翔問的是特定標的，只回答那個標的的資訊，嚴禁列出其他資產。
- 如果問到美元金額，直接使用資料中的原幣數字回答，不要說「需要換算」。
- 如果同一個合併鍵有多筆（例如不同平台），把數字加總後回答。
- 如果資料裡有現成的數字，直接引用，不要含糊帶過。
- 回答金額時，原幣和 TWD 都要提供。
- 所有金額一律四捨五入到整數，不要出現小數點。

## 回覆格式
- 不需要每次都叫「翔翔」，自然就好，偶爾提到即可。
- 針對問題直接回答，不要繞圈子。
- 只做觀察與描述，除非翔翔問「要不要賣」，否則不主動給予操作建議。

## 基礎資訊
- USD/TWD 匯率：${usdRate}
- 資產總淨值(TWD)：${totalNetWorth.toLocaleString()}
- 已實現損益(TWD)：${Math.round(realizedReturnTwd).toLocaleString()}
- 已實現損益(%)：${realizedReturnPct}%
`;

  if (matchedAssets.length > 0) {
    // 🎯 模式 A：特定標的
    const assetStr = matchedAssets.map(formatRow).join('\n');
    systemInstruction += `
## 當前任務：只分析以下標的（可能有多筆，屬於同一合併鍵的不同平台）
${assetStr}

⚠️ 嚴禁列出或提及上面以外的任何標的。`;
  } else {
    // 📊 模式 B：整體戰情
    const assetStr = allAssets.map(formatRow).join('\n');
    systemInstruction += `
## 當前任務：分析整體戰情
全部持倉明細：
${assetStr}`;
  }

  /* ================================
     🚀 呼叫 Gemini
  ================================= */
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" + GEMINI_API_KEY;
  const prevMessages = history || [];
  const currentMsg = { role: "user", parts: [{ text: userQuery }] };

  const payload = {
    system_instruction: { parts: [{ text: systemInstruction }] },
    contents: [...prevMessages, currentMsg],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 500
    }
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const json = JSON.parse(res.getContentText());
    if (json.error) return "翔翔，系統干擾：" + json.error.message;

    let reply = json.candidates?.[0]?.content?.parts?.[0]?.text || "訊號中斷，翔翔請稍後。";
    return reply.replace(/[\*#_~`\[\]]/g, "").trim();
  } catch (e) {
    return "伺服器波動，翔翔請稍後再試。";
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

/* ================================
   🧪 AI 測試工具（不用部署）
================================ */
function testAI() {
  const queries = ["TSLA表現如何", "把每一個股票的報酬率列表給我"];
  
  queries.forEach(q => {
    const result = callGeminiAnalysis(q);
    Logger.log("--- 測試開始 ---");
    Logger.log("問題: " + q);
    Logger.log("AI回答: " + result);
  });
}