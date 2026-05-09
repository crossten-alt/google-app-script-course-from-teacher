// ============================================================
// 進階練習：客戶 CRM 資料批次清理與驗證
// 對應：Session 4（Ranges、批次處理、陣列）
// ============================================================

/**
 * 批次驗證客戶資料
 * 檢查：Email 格式、電話格式、必填欄位、重複資料
 */
function TY_批次驗證客戶資料() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("客戶資料");
    if (!sheet) { SpreadsheetApp.getUi().alert("❌ 請先執行初始化"); return; }

    // 強制同步試算表狀態，確保讀取到的是最新的（特別是清理後的）資料
    SpreadsheetApp.flush();
    var 資料 = sheet.getDataRange().getValues();
    var 驗證結果 = [];
    var 問題數 = 0;
    var Email清單 = {};

    for (var i = 1; i < 資料.length; i++) {
      var 公司 = String(資料[i][0]).trim();
      var 統編 = String(資料[i][1]).trim(); // 新增統編
      var 聯絡人 = String(資料[i][2]).trim(); // 索引加1
      var Email = String(資料[i][3]).trim(); // 索引加1
      var 電話 = String(資料[i][4]).trim(); // 索引加1
      var 錯誤 = [];

      // 驗證統編 (8碼數字)
      if (統編 && !/^\d{8}$/.test(統編)) {
        錯誤.push("統編必須為 8 碼數字");
      }

      // 檢查必填
      if (!公司) 錯誤.push("公司名稱空白");
      if (!聯絡人) 錯誤.push("聯絡人空白");

      // 驗證 Email 格式
      if (Email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(Email)) {
        錯誤.push("Email 格式錯誤");
      }

      // 檢查 Email 重複
      if (Email) {
        if (Email清單[Email]) {
          錯誤.push("Email 重複（與第 " + Email清單[Email] + " 列）");
        } else {
          Email清單[Email] = i + 1;
        }
      }

      // 驗證電話（台灣格式）
      if (電話 && !/^0[0-9\-]{8,12}$/.test(電話.replace(/\s/g, ""))) {
        錯誤.push("電話格式可疑");
      }

      var 狀態 = 錯誤.length === 0 ? "✅ 正常" : "⚠️ " + 錯誤.join("；");
      驗證結果.push([狀態]);
      if (錯誤.length > 0) 問題數++;
    }

    // 批次寫入驗證結果到 I 欄
    sheet.getRange("I1").setValue("驗證結果").setFontWeight("bold");
    sheet.getRange(2, 9, 驗證結果.length, 1).setValues(驗證結果);

    // 標示有問題的列（改用批次寫入背景色，效能大幅提升並確保 UI 更新）
    var 背景色 = [];
    for (var j = 0; j < 驗證結果.length; j++) {
      if (驗證結果[j][0].indexOf("⚠️") >= 0) {
        背景色.push(["#fff3e0", "#fff3e0", "#fff3e0", "#fff3e0", "#fff3e0", "#fff3e0", "#fff3e0", "#fff3e0", "#fff3e0"]); // 多加一個
      } else {
        背景色.push(["#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff"]); // 多加一個
      }
    }
    if (背景色.length > 0) {
      sheet.getRange(2, 1, 背景色.length, 9).setBackgrounds(背景色);
    }

    sheet.autoResizeColumn(9);
    SpreadsheetApp.getUi().alert("✅ 驗證完成！\n\n總筆數：" + 驗證結果.length +
      "\n問題筆數：" + 問題數 + "\n正常筆數：" + (驗證結果.length - 問題數));

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); }
}

/**
 * 批次清理：統一格式、去除空白、修正大小寫
 */
function TY_批次清理資料() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("客戶資料");
    if (!sheet) return;

    var 資料 = sheet.getDataRange().getValues();
    var 清理計數 = 0;

    for (var i = 1; i < 資料.length; i++) {
      // 去除前後空白
      for (var j = 0; j < 資料[i].length; j++) {
        if (typeof 資料[i][j] === "string") {
          var 原值 = 資料[i][j];
          資料[i][j] = 資料[i][j].trim().replace(/\s+/g, " ");
          if (原值 !== 資料[i][j]) 清理計數++;
        }
      }

      // Email 統一小寫 (索引從 2 變成 3)
      if (資料[i][3]) {
        var 原Email = 資料[i][3];
        資料[i][3] = String(資料[i][3]).toLowerCase();
        if (原Email !== 資料[i][3]) 清理計數++;
      }

      // 電話統一格式清理 (索引從 3 變成 4)
      if (資料[i][4]) {
        var 原電話 = String(資料[i][4]).trim();
        var 純數字 = 原電話.replace(/[^0-9]/g, "");
        
        // 補足被試算表數值格式吃掉的字首 0 (例如 912345678 -> 0912345678)
        if (純數字.length === 9 && 純數字[0] !== "0") {
          純數字 = "0" + 純數字;
        }

        var 新電話 = 原電話;
        if (純數字.length === 10) {
          if (純數字.indexOf("09") === 0) {
            // 手機格式: 09xx-xxx-xxx
            新電話 = 純數字.substring(0, 4) + "-" + 純數字.substring(4, 7) + "-" + 純數字.substring(7);
          } else {
            // 2 位數區碼市話 (如 02, 04): 0x-xxxx-xxxx
            新電話 = 純數字.substring(0, 2) + "-" + 純數字.substring(2, 6) + "-" + 純數字.substring(6);
          }
        } else if (純數字.length === 9) {
          // 2 位數區碼市話 (如 03, 05, 06): 0x-xxx-xxxx
          新電話 = 純數字.substring(0, 2) + "-" + 純數字.substring(2, 5) + "-" + 純數字.substring(5);
        }

        if (原電話 !== 新電話) {
          資料[i][4] = 新電話;
          清理計數++;
        }
      }
    }

    // 批次寫回
    sheet.getRange(1, 1, 資料.length, 資料[0].length).setValues(資料);
    
    // 強制同步狀態，確保下一次「驗證」能抓到清理後的正確資料
    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert("✅ 清理完成！共修正 " + 清理計數 + " 處\n請點選「批次驗證」更新最新狀態。");

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); }
}

/**
 * 批次匯入：從「待匯入」工作表匯入新客戶（自動去重）
 */
function TY_批次匯入新客戶() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 來源 = ss.getSheetByName("待匯入");
    var 目標 = ss.getSheetByName("客戶資料");
    if (!來源 || !目標) { SpreadsheetApp.getUi().alert("❌ 缺少工作表"); return; }

    // 取得現有 Email 清單（去重用）
    var 現有資料 = 目標.getDataRange().getValues();
    var 現有Email = {};
    for (var i = 1; i < 現有資料.length; i++) {
      現有Email[String(現有資料[i][3]).toLowerCase()] = true; // Email 索引為 3
    }

    // 取得待匯入資料
    var 新資料 = 來源.getDataRange().getValues();
    var 匯入 = [];
    var 重複 = 0;

    for (var j = 1; j < 新資料.length; j++) {
      var email = String(新資料[j][3]).toLowerCase().trim(); // Email 索引為 3
      if (現有Email[email]) {
        重複++;
      } else {
        匯入.push(新資料[j]);
        現有Email[email] = true;
      }
    }

    // 批次寫入
    if (匯入.length > 0) {
      var 新起始列 = 目標.getLastRow() + 1;
      目標.getRange(新起始列, 1, 匯入.length, 匯入[0].length).setValues(匯入);
    }

    SpreadsheetApp.getUi().alert("✅ 匯入完成！\n新增：" + 匯入.length + " 筆\n跳過重複：" + 重複 + " 筆");

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); }
}

function TY_初始化客戶資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("客戶資料");
  if (!sheet) sheet = ss.insertSheet("客戶資料"); else sheet.clear();

  var 標題 = [["公司名稱", "統編", "聯絡人", "Email", "電話", "地區", "產業", "客戶等級"]];
  var 資料 = [
    ["台北科技股份有限公司", "12345678", "王大明", "wang@taipei-tech.com", "0912345678", "台北", "科技業", "A"], // 正確
    ["新竹電子有限公司", "2345678", "李小華", "LEE@hsinchu.COM", "0923-456-789", "新竹", "電子業", "B"], // 錯誤：只有7碼
    ["台中製造公司", "34567890", " 張美玲 ", "chang@taichung-mfg.com", "04-2234-5678", "台中", "製造業", "A"], // 正確
    ["高雄物流", "ABCDEFGH", "陳大文", "chen@kaohsiung", "invalid", "高雄", "物流業", "C"], // 錯誤：英文
    ["花蓮文創工作室", "56789012", "", "info@hualien-creative.com", "0389-001-234", "花蓮", "文創", "B"], // 正確
    ["桃園光電科技", "67890123", "黃志偉", "huang@taoyuan-opto.com", "0956789012", "桃園", "光電", "A"], // 正確
    ["基隆貿易商行", "78901234", "林采潔", "wang@taipei-tech.com", "02-2456-7890", "基隆", "貿易", "C"], // 正確
    ["宜蘭農產公司", "89012345", "周建國", "chou@yilan-agri.com", "0398765432", "宜蘭", "農業", "B"] // 正確
  ];

  sheet.getRange(1, 1, 1, 8).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 8).setValues(資料);
  sheet.getRange("A1:H1").setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 8; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 客戶資料已建立（含刻意的錯誤資料供驗證練習）");
}

function TY_onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 CRM (新版測試)")
    .addItem("📦 初始化客戶資料", "TY_初始化客戶資料")
    .addItem("📦 初始化待匯入資料", "TY_初始化待匯入資料")
    .addItem("🔍 批次驗證", "TY_批次驗證客戶資料")
    .addItem("🧹 批次清理", "TY_批次清理資料")
    .addItem("📥 批次匯入新客戶", "TY_批次匯入新客戶")
    .addToUi();
}

/**
 * 幫你自動建立「待匯入」測試資料
 */
function TY_初始化待匯入資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("待匯入");
  if (!sheet) sheet = ss.insertSheet("待匯入"); else sheet.clear();

  var 標題 = [["公司名稱", "統編", "聯絡人", "Email", "電話", "地區", "產業", "客戶等級"]];
  var 資料 = [
    // 第 1 筆：Email 跟原本的一模一樣，測試系統會不會把它「跳過（去重複）」
    ["台北科技股份有限公司", "12345678", "王大明", "wang@taipei-tech.com", "0912345678", "台北", "科技業", "A"], 
    
    // 第 2 筆：全新的 Email，測試系統會不會成功「新增」
    ["台南軟體開發", "98765432", "林大強", "lin@tainan-soft.com", "0988-777-666", "台南", "軟體業", "A"], 
    
    // 第 3 筆：Email 跟原本的一樣，測試「去重複」
    ["台中製造公司", "34567890", " 張美玲 ", "chang@taichung-mfg.com", "04-2234-5678", "台中", "製造業", "A"] 
  ];

  sheet.getRange(1, 1, 1, 8).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 8).setValues(資料);
  sheet.getRange("A1:H1").setBackground("#1565c0").setFontColor("#fff").setFontWeight("bold");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 8; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 「待匯入」工作表已建立！請點選選單的「批次匯入新客戶」測試看看。");
}
