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
   4️⃣ AI 助理分析 - V7 雙模式務實版
   (整合中英偵測、自動排序、去客服語氣)
================================ */
function callGeminiAnalysis(userQuery) {
  if (!GEMINI_API_KEY) return "⚠️ 翔翔，API Key 未設定。";

  // 1. 抓取最新資料 (AI 專用，直接讀取 Sheet 最新狀態)
  const data = getDashboardData(null);
  if (!data || !data.assets) return "翔翔，目前讀不到資料，請檢查 Sheet 狀態。";

  /* ================================
     🔍 模式偵測：掃描問題中是否包含特定標的
  ================================= */
  let targetAsset = null;
  const q = String(userQuery || "").toLowerCase();

  // 排序：報酬率由低到高 (讓 AI 優先關注虧損標的)
  const sortedAssets = [...data.assets].sort((a, b) => a.returnRate - b.returnRate);

  // 偵測邏輯：直接掃描名稱，包含中英文識別
  for (const a of sortedAssets) {
    const fullName = String(a.name || "").toLowerCase();
    const symbolPart = fullName.split('(')[0].trim(); // 抓括號前的代號
    const namePart = fullName.includes('(') ? fullName.split('(')[1].replace(')', '').trim() : ""; // 抓括號內的中文

    if (q.includes(symbolPart) || (namePart && q.includes(namePart))) {
      targetAsset = a;
      break; 
    }
  }

  /* ================================
     📊 格式化資產字串 (提升 AI 解析效率)
  ================================= */
  const assetStr = sortedAssets.map(a => 
    `${a.name} | 市值 ${Math.round(a.value).toLocaleString()} | 報酬 ${a.returnRate.toFixed(2)}%`
  ).join('\n');

  /* ================================
     🧠 V7 務實派系統指令 (Prompt)
  ================================= */
  let systemInstruction = `
## 定位
你是一個理性、成熟、對翔翔有好感的投資資料分析助手。語氣自然、簡短，像真人聊天。

## 語氣禁令
- 嚴禁客服腔：禁止「很高興為您服務」、「竭誠為您分析」、「請注意」。
- 嚴禁矯情陪伴：禁止「我陪你」、「不用擔心」、「姊姊在看」。
- 避開銀行行話：不要說「獲利了結」、「配置調整」、「短期偏熱」。

## 任務與結構
1. 第一段：一句話總結（必須包含「翔翔」二字，位置不限）。
2. 第二段：列表顯示（若資產數量 >= 2）。格式：• 名稱：市值 X，報酬 X%。
3. 第三段：一句短評論（描述事實或波動來源，可省略）。
4. 只做觀察與描述，除非翔翔問「要不要賣」，否則不主動給予操作建議。
`;

  if (targetAsset) {
    // 🎯 模式 A：單一標的焦點模式
    const info = `${targetAsset.name} | 市值 ${Math.round(targetAsset.value).toLocaleString()} | 報酬 ${targetAsset.returnRate.toFixed(2)}%`;
    systemInstruction += `\n現在只需分析此單一標的：\n${info}`;
  } else {
    // 📊 模式 B：整體組合戰情模式
    systemInstruction += `\n現在分析整體戰情：\n總市值：${Math.round(data.investTotal).toLocaleString()} TWD\n已實現損益：${Math.round(data.realizedReturnTwd).toLocaleString()} TWD\n資產細節(已由差到好排序)：\n${assetStr}`;
  }

  /* ================================
     🚀 呼叫 Gemini
  ================================= */
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" + GEMINI_API_KEY;
  const payload = {
    contents: [{ role: "user", parts: [{ text: systemInstruction + "\n問題：" + userQuery }] }],
    generationConfig: { 
      temperature: 0.2, // 保持穩定冷靜
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
    
    // 移除所有 Markdown 符號（保持介面乾淨）
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
  const queries = ["TSLA表現如何", "目前整體戰情"];
  
  queries.forEach(q => {
    const result = callGeminiAnalysis(q);
    Logger.log("--- 測試開始 ---");
    Logger.log("問題: " + q);
    Logger.log("AI回答: " + result);
  });
}