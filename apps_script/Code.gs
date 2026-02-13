/**
 * 投資戰情室後端 V6.16 - 超強容錯與數值解析版
 * 解決 Google Sheet 千分位逗號 (,) 導致數值被誤判為字串而消失的問題
 */
const CONFIG = {
  SPREADSHEET_ID: "1HM2MvZepqo1LVvgRoWwQ-1NmWKxo3ASAcXc2wECPgZU",
  SHEET_LOGS: "買賣紀錄_2026",
  SHEET_SUMMARY: "庫存彙整(統整)",
  SHEET_HISTORY: "淨值歷史",
  SHEET_ASSETS: "資產統計(彙整)", 
  SHEET_REGIONS: "投資地區"      
};

function doGet() {
  const pageTitle = "投資戰情室 V6.16";
  const possibleNames = ["ui", "ui.html", "Index", "apps_script/ui"];
  for (let name of possibleNames) {
    try {
      return HtmlService.createHtmlOutputFromFile(name)
        .setTitle(pageTitle)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
    } catch (e) {}
  }
  return HtmlService.createHtmlOutput("找不到網頁檔案，請確認檔名為 ui");
}

/** 🛠 核心修復：安全解析所有帶逗號的字串為數字 */
function parseNum_(val) {
  if (val === "" || val === null || val === undefined) return 0;
  if (typeof val === 'number') return val;
  // 移除所有逗號再轉為數字
  return Number(String(val).replace(/,/g, '')) || 0;
}

function getDashboardData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // 1. 淨值歷史
  const histSh = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  let history = [];
  if (histSh && histSh.getLastRow() >= 2) {
    const histData = histSh.getRange(Math.max(2, histSh.getLastRow() - 29), 1, 30, 2).getValues();
    history = histData
      .filter(r => r[0] && parseNum_(r[1]) > 0) 
      .map(r => ({
        date: (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+8", "MM/dd") : String(r[0]),
        val: parseNum_(r[1])
      }));
  }

  // 2. 資產統計 (圓餅圖 1) 
  const assetSh = ss.getSheetByName(CONFIG.SHEET_ASSETS);
  let assets = [];
  if (assetSh && assetSh.getLastRow() >= 2) {
    const assetData = assetSh.getRange(2, 1, assetSh.getLastRow() - 1, 4).getValues();
    assets = assetData.map(r => ({
      name: String(r[0] || "").trim(),
      value: parseNum_(r[3]) // 👈 強制把 D 欄含逗號的文字轉回數字
    })).filter(h => h.value > 0 && h.name !== ""); // 過濾掉空名或數值為0的行
  }

  // 3. 投資地區 (圓餅圖 2)
  const regionSh = ss.getSheetByName(CONFIG.SHEET_REGIONS);
  let regions = [];
  if (regionSh && regionSh.getLastRow() >= 2) {
    const regionData = regionSh.getRange(2, 1, regionSh.getLastRow() - 1, 2).getValues();
    regions = regionData.map(r => ({
      name: String(r[0] || "").trim(),
      value: parseNum_(r[1])
    })).filter(h => h.value > 0 && h.name !== "");
  }

  // 4. 關鍵彙總數據 (使用智慧搜尋)
  const logSh = ss.getSheetByName(CONFIG.SHEET_LOGS);
  let investTotal = 0, usdRate = 32.0, realizedReturn = 0, realizedReturnTwd = 0;

  if (logSh) {
    // 智慧掃描 Y 欄的標籤，並安全解析 Z 欄數值
    const summaryData = logSh.getRange("Y1:Z30").getValues();
    for (let i = 0; i < summaryData.length; i++) {
      const label = String(summaryData[i][0] || "").trim();
      const val = summaryData[i][1];
      
      if (label.includes("已實現總損益(TWD)")) {
        realizedReturnTwd = parseNum_(val);
      }
      if (label.includes("已實現總損益(%)")) {
        const rawStr = String(val);
        if (rawStr.includes('%')) {
          // 若試算表回傳 "22.29%"，去逗號、去%直接轉數字
          realizedReturn = Number(rawStr.replace(/,/g, '').replace(/%/g, '')) || 0;
        } else {
          // 若試算表回傳小數 0.2229，則乘 100
          realizedReturn = parseNum_(val) * 100;
        }
      }
    }

    const rawRate = logSh.getRange("H86").getValue(); 
    usdRate = parseNum_(rawRate) > 0 ? parseNum_(rawRate) : 32.2; 
    investTotal = regions.reduce((sum, item) => sum + item.value, 0);
  }

  return { history, assets, regions, investTotal, usdRate, realizedReturn, realizedReturnTwd };
}

function saveTrades(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.SHEET_LOGS);
    if (!sh) throw new Error("找不到分頁：" + CONFIG.SHEET_LOGS);

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const getCol = (name) => headers.indexOf(name);
    const startRow = findFirstEmptyRow_(sh);
    const rows = payload.trades.map((t, i) => buildFormulaRow_(headers, payload.defaults || {}, t, startRow + i, getCol));
    
    sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    SpreadsheetApp.flush();
    return { ok: true, row: startRow };
  } catch (e) {
    throw new Error(e.message);
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
  const type = String(t.type || "").trim();
  const isSell = type.includes("賣");
  let rowData = new Array(headers.length).fill("");
  const setVal = (name, val) => { const idx = getCol(name); if (idx !== -1) rowData[idx] = val; };

  setVal("日期", t.date || new Date().toLocaleDateString('zh-TW'));
  setVal("交易類型", type);
  setVal("平台", t.platform || defaults.platform);
  setVal("帳戶類型", t.account || defaults.account);
  setVal("名稱", t.name || "");
  setVal("股票代號", t.symbol || "");

  const accountType = t.account || defaults.account || "";
  const platformName = t.platform || defaults.platform || "";
  const isUSD = accountType.includes("USD") || platformName.includes("Firstrade") || platformName.includes("IBKR");
  setVal("幣別", isUSD ? "USD" : "TWD");

  if (isSell) {
    setVal("賣出價格", Number(t.price));
    setVal("賣出股數", Number(t.qty));
    setVal("成本(原幣)※賣出需填", Number(t.cost));
  } else {
    setVal("買入價格", Number(t.price));
    setVal("買入股數", Number(t.qty));
  }
  setVal("手續費", Number(t.fee || 0));
  setVal("交易稅", Number(t.tax || 0));
  setVal("狀態", "已完成");

  setVal("價金(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), I${r}*J${r}, K${r}*L${r})`);
  setVal("應收付(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), P${r}-M${r}-N${r}, P${r}+M${r}+N${r})`);
  setVal("損益(原幣)", `=IF(ISNUMBER(SEARCH("賣",B${r})), Q${r}-O${r}, "")`);
  setVal("報酬率", `=IF(AND(ISNUMBER(R${r}), O${r}<>0), R${r}/O${r}, "")`);
  setVal("成本(TWD)", `=IF(O${r}<>"", O${r}*IF(H${r}="",1,H${r}), "")`);
  setVal("應收付(TWD)", `=Q${r}*IF(H${r}="",1,H${r})`);
  setVal("損益(TWD)", `=IF(R${r}<>"", R${r}*IF(H${r}="",1,H${r}), "")`);

  return rowData;
}