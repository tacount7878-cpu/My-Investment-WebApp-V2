/**
 * 投資戰情室後端 V6.3 - 穩定純寫入版
 * 核心功能：
 * 1. 只負責將 UI 資料「追加」到買賣紀錄表。
 * 2. 絕對不執行清空或重整，保護使用者的初始資料。
 * 3. 確保「名稱」與「幣別」正確寫入。
 */

const CONFIG = {
  SPREADSHEET_ID: "1HM2MvZepqo1LVvgRoWwQ-1NmWKxo3ASAcXc2wECPgZU",
  SHEET_LOGS: "買賣紀錄_2026",
  SHEET_SUMMARY: "庫存彙整(統整)",
  SHEET_HISTORY: "淨值歷史"
};

/** 網頁入口 */
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile("ui")
      .setTitle("Investment War Room V6.3")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  } catch (e) {
    return HtmlService.createHtmlOutputFromFile("apps_script/ui")
      .setTitle("Investment War Room V6.3")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }
}

/** 儲存交易（唯一寫入入口） */
function saveTrades(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.SHEET_LOGS);
    if (!sh) throw new Error("找不到分頁：" + CONFIG.SHEET_LOGS);

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    
    // 欄位定位 helper (保留備用)
    const getCol = (name) => {
      const idx = headers.indexOf(name);
      return idx; // 若找不到回傳 -1，由 buildFormulaRow_ 處理
    };

    // 1. 找到寫入位置 (從 86 開始)
    const startRow = findFirstEmptyRow_(sh);

    // 2. 建立資料列 (含公式與名稱)
    const rows = payload.trades.map((t, i) => 
      buildFormulaRow_(headers, payload.defaults || {}, t, startRow + i, getCol)
    );
    
    // 3. 寫入
    sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    
    // 強制刷新試算表
    SpreadsheetApp.flush();

    return { ok: true, row: startRow };

  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

/** 從第 86 列開始尋找第一個空白列 */
function findFirstEmptyRow_(sh) {
  const START_ROW = 86;
  const lastRow = sh.getLastRow();
  if (lastRow < START_ROW) return START_ROW;

  // 只掃描 86 ~ lastRow 的 A 欄
  const range = sh.getRange(START_ROW, 1, lastRow - START_ROW + 1, 1);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "" || values[i][0] === null) {
      return START_ROW + i;
    }
  }
  return lastRow + 1;
}

/** 建立資料列（核心修正：名稱與幣別） */
function buildFormulaRow_(headers, defaults, t, r, getCol) {
  const type = String(t.type || "").trim();
  const isSell = type.includes("賣");
  const p = Number(t.price || 0), q = Number(t.qty || 0), f = Number(t.fee || 0), x = Number(t.tax || 0);
  const c = isSell ? Number(t.cost || 0) : "";

  let rowData = new Array(headers.length).fill("");

  // 1. 固定欄位填寫 (不依賴 getCol，直接鎖定順序或名稱)
  // 日期 (A欄)
  if (getCol("日期") !== -1) rowData[getCol("日期")] = t.date || new Date().toLocaleDateString('zh-TW');
  // 交易類型 (B欄)
  if (getCol("交易類型") !== -1) rowData[getCol("交易類型")] = type;
  // 平台 (C欄)
  if (getCol("平台") !== -1) rowData[getCol("平台")] = t.platform || defaults.platform;
  // 帳戶類型 (D欄)
  if (getCol("帳戶類型") !== -1) rowData[getCol("帳戶類型")] = t.account || defaults.account;

  // ✅ 幣別自動判斷 (E欄)
  const accountType = t.account || defaults.account || "";
  const platformName = t.platform || defaults.platform || "";
  const isUSD = accountType.includes("USD") || 
                platformName.includes("Firstrade") || 
                platformName.includes("IBKR") || 
                platformName.includes("美股");
  
  if (getCol("幣別") !== -1) rowData[getCol("幣別")] = isUSD ? "USD" : "TWD";

  // ✅ 關鍵修正：名稱確實寫入 (F欄)
  // 使用 || "" 確保不會寫入 undefined
  if (getCol("名稱") !== -1) rowData[getCol("名稱")] = t.name || "";

  // 股票代號 (G欄)
  if (getCol("股票代號") !== -1) rowData[getCol("股票代號")] = t.symbol || "";
  // 匯率 (H欄) - 留空
  if (getCol("匯率(可空)") !== -1) rowData[getCol("匯率(可空)")] = "";

  // 2. 數值填入
  if (isSell) {
    if (getCol("賣出價格") !== -1) rowData[getCol("賣出價格")] = p;
    if (getCol("賣出股數") !== -1) rowData[getCol("賣出股數")] = q;
    if (getCol("成本(原幣)※賣出需填") !== -1) rowData[getCol("成本(原幣)※賣出需填")] = c;
  } else {
    if (getCol("買入價格") !== -1) rowData[getCol("買入價格")] = p;
    if (getCol("買入股數") !== -1) rowData[getCol("買入股數")] = q;
  }

  if (getCol("手續費") !== -1) rowData[getCol("手續費")] = f;
  if (getCol("交易稅") !== -1) rowData[getCol("交易稅")] = x;

  // 3. 注入公式 (使用列號 r)
  if (getCol("價金(原幣)") !== -1) rowData[getCol("價金(原幣)")] = `=IF(ISNUMBER(SEARCH("賣",B${r})), I${r}*J${r}, K${r}*L${r})`;
  if (getCol("應收付(原幣)") !== -1) rowData[getCol("應收付(原幣)")] = `=IF(ISNUMBER(SEARCH("賣",B${r})), P${r}-M${r}-N${r}, P${r}+M${r}+N${r})`;
  if (getCol("損益(原幣)") !== -1) rowData[getCol("損益(原幣)")] = `=IF(ISNUMBER(SEARCH("賣",B${r})), Q${r}-O${r}, "")`;
  if (getCol("報酬率") !== -1) rowData[getCol("報酬率")] = `=IF(AND(ISNUMBER(R${r}), O${r}<>0), R${r}/O${r}, "")`;
  if (getCol("成本(TWD)") !== -1) rowData[getCol("成本(TWD)")] = `=IF(O${r}<>"", O${r}*IF(H${r}="",1,H${r}), "")`;
  if (getCol("應收付(TWD)") !== -1) rowData[getCol("應收付(TWD)")] = `=Q${r}*IF(H${r}="",1,H${r})`;
  if (getCol("損益(TWD)") !== -1) rowData[getCol("損益(TWD)")] = `=IF(R${r}<>"", R${r}*IF(H${r}="",1,H${r}), "")`;
  
  if (getCol("狀態") !== -1) rowData[getCol("狀態")] = "已完成";

  return rowData;
}

/** 讀取 Dashboard 資料（UI 用） */
function getDashboardData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sumSh = ss.getSheetByName(CONFIG.SHEET_SUMMARY);
  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);

  let holdings = [];
  if (sumSh && sumSh.getLastRow() >= 2) {
    const data = sumSh.getRange(2, 1, sumSh.getLastRow() - 1, 6).getValues();
    holdings = data.map(r => ({
      name: r[0],
      value: Number(r[3] || 0),
      roi: Number(r[5] || 0)
    })).filter(h => h.value > 0);
  }

  let history = [];
  if (histSh && histSh.getLastRow() >= 2) {
    const histData = histSh.getRange(Math.max(2, histSh.getLastRow() - 29), 1, 30, 2).getValues();
    history = histData.map(r => ({
      date: Utilities.formatDate(r[0], "GMT+8", "MM/dd"),
      val: Number(r[1] || 0)
    }));
  }

  return { holdings, history };
}

/** 🧪 模擬測試：Firstrade 買入 TSLA (驗證名稱寫入) **/
function testBuySimulation() {
  const mockPayload = {
    defaults: { 
      platform: "Firstrade(FT)",   
      account: "USD外幣帳戶" 
    },
    trades: [{
      type: "買入",
      name: "特斯拉測試V6.3", // 👈 測試名稱
      symbol: "TSLA",
      price: 350,
      qty: 1,
      fee: 0,
      tax: 0,
      date: new Date().toLocaleDateString('zh-TW')
    }]
  };
  const result = saveTrades(mockPayload);
  Logger.log("✅ 測試完成 → 寫入列號：" + result.row);
}