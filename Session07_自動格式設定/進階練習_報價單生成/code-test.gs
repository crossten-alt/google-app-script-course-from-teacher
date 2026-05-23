// ============================================================
// 進階練習：一鍵生成客戶報價單（全新直式LOGO完美排版版）
// 對應：Session 7（格式設定、文字/數字/日期格式）
// ============================================================

/**
 * 建立自訂功能選單
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 報價單系統")
    .addItem("📦 初始化報價資料", "初始化報價資料")
    .addItem("📋 生成報價單", "生成報價單")
    .addToUi();
}

/**
 * 一鍵生成專業報價單
 */
function 生成報價單() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 客戶表 = ss.getSheetByName("報價資料");
    if (!客戶表) { 
      SpreadsheetApp.getUi().alert("❌ 請先點選「📦 初始化報價資料」建立基礎資料。"); 
      return; 
    }

    var ui = SpreadsheetApp.getUi();
    var 客戶回應 = ui.prompt("📝 報價單", "請輸入客戶名稱：", ui.ButtonSet.OK_CANCEL);
    if (客戶回應.getSelectedButton() !== ui.Button.OK) return;
    var 客戶 = 客戶回應.getResponseText().trim() || "範例客戶";

    // 產生專屬報價單號
    var 今天 = new Date();
    var 編號 = "QT-" + Utilities.formatDate(今天, "Asia/Taipei", "yyyyMMdd") + "-" +
               String(Math.floor(Math.random() * 100)).padStart(3, "0");

    // 建立新工作表並推至最前端
    var sheet = ss.insertSheet(編號, 0);

    // ============================================================
    // 欄寬與基礎設定
    // ============================================================
    sheet.setColumnWidth(1, 60);   // 項次
    sheet.setColumnWidth(2, 260);  // 品名/規格
    sheet.setColumnWidth(3, 60);   // 單位
    sheet.setColumnWidth(4, 80);   // 數量
    sheet.setColumnWidth(5, 100);  // 單價
    sheet.setColumnWidth(6, 120);  // 金額

    var allRange = sheet.getRange("A1:F50");
    allRange.setFontFamily("Arial");
    allRange.setVerticalAlignment("middle");

    // ============================================================
    // 1. 公司 Logo 與抬頭資訊區 (Row 1 ~ 3) - 針對直式圖片完美調校
    // ============================================================
    sheet.setRowHeight(1, 40); // 第一列高度
    sheet.setRowHeight(2, 40); // 第二列高度（兩列合併給予直式圖片足夠高度）

    // 左側大標題
    sheet.getRange("A1:E2").merge();
    sheet.getRange("A1").setValue("會議與系統整合報價單")
      .setFontSize(20)
      .setFontWeight("bold")
      .setFontColor("#1a237e")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");

    // 右側置入新 Logo 圖片 (合併 F1:F2，確保太空人火箭圖等比例完美呈現)
    sheet.getRange("F1:F2").merge();
    var logoUrl = "https://upload.wikimedia.org/wikipedia/zh/thumb/6/61/ONE_PIECE_Logo.svg/330px-ONE_PIECE_Logo.svg.png"; 
    sheet.getRange("F1").setValue('=IMAGE("' + logoUrl + '", 1)')
                        .setHorizontalAlignment("center")
                        .setVerticalAlignment("middle");

    // 限制 Logo 尺寸與微調位置（已換算為完美的 90% 視覺比例）
    logo.setHeight(72);          // 垂直縮放至 72 像素 (總高 80)
    logo.setWidth(108);          // 水平縮放至 108 像素 (總寬 120)
    logo.setAnchorCellYOffset(4); // 上方留空 4 像素，達成完美的垂直置中
    logo.setAnchorCellXOffset(6); // 左方留空 6 像素，達成完美的水平置中

    // 第三列公司地址資訊（順延至此，排版乾淨不重疊）
    sheet.setRowHeight(3, 25);
    sheet.getRange("A3:F3").merge();
    sheet.getRange("A3").setValue("拉夫德魯島7 號 ｜ Tel: 02-2345-6789 ｜ www.fish-tech.com")
      .setFontSize(9)
      .setFontColor("#666666")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");

    // ============================================================
    // 2. 報價單大標題與客戶資訊 (Row 5 ~ 10)
    // ============================================================
    sheet.getRange("A5:F5").merge();
    sheet.getRange("A5").setValue("📋 報 價 單")
      .setFontSize(22)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#1a237e")
      .setFontColor("#ffffff");
    sheet.setRowHeight(5, 45);

    var 有效日 = new Date(今天); 
    有效日.setDate(今天.getDate() + 30);

    // 資訊列下移一列 (從第 7 列開始填入)
    sheet.getRange(7, 1, 4, 6).setValues([
      ["報價編號", 編號, "", "客戶名稱", 客戶, ""],
      ["報價日期", Utilities.formatDate(today = 今天, "Asia/Taipei", "yyyy/MM/dd"), "", "聯絡人", "", ""],
      ["有效期限", Utilities.formatDate(有效日, "Asia/Taipei", "yyyy/MM/dd"), "", "聯絡電話", "", ""],
      ["業務人員", "草帽小子", "", "傳真", "", ""]
    ]);

    for (var r = 7; r <= 10; r++) {
      sheet.setRowHeight(r, 26);
      sheet.getRange(r, 2, 1, 2).merge().setHorizontalAlignment("left");
      sheet.getRange(r, 5, 1, 2).merge().setHorizontalAlignment("left");

      sheet.getRange(r, 1).setFontWeight("bold").setBackground("#e8eaf6").setHorizontalAlignment("center");
      sheet.getRange(r, 4).setFontWeight("bold").setBackground("#e8eaf6").setHorizontalAlignment("center");
    }
    sheet.getRange("A7:F10").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);

    // ============================================================
    // 3. 品項表格與動態生成 (Row 12 開始)
    // ============================================================
    sheet.getRange("A12:F12").setValues([["項次", "品名/規格", "單位", "數量", "單價", "金額"]]);
    sheet.getRange("A12:F12").setBackground("#283593").setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setRowHeight(12, 32);

    var 品項資料 = 客戶表.getDataRange().getValues();
    var 小計 = 0;

    for (var i = 1; i < 品項資料.length; i++) {
      var 列 = 12 + i;
      var 金額 = 品項資料[i][3] * 品項資料[i][4];
      小計 += 金額;

      sheet.setRowHeight(列, 30);
      sheet.getRange(列, 1, 1, 6).setValues([[
        i, 品項資料[i][0], 品項資料[i][1], 品項資料[i][3], 品項資料[i][4], 金額
      ]]);
      
      sheet.getRange(列, 1).setHorizontalAlignment("center");
      sheet.getRange(列, 2).setHorizontalAlignment("left");
      sheet.getRange(列, 3).setHorizontalAlignment("center");
      sheet.getRange(列, 4, 1, 3).setNumberFormat("#,##0").setHorizontalAlignment("right");

      if (i % 2 === 0) {
        sheet.getRange(列, 1, 1, 6).setBackground("#f9f9f9");
      }
    }

    // ============================================================
    // 4. 金額總計與稅金計算區
    // ============================================================
    var 稅金 = Math.round(小計 * 0.05);
    var 總計 = 小計 + 稅金;
    var 合計起始 = 13 + 品項資料.length - 1;

    var 合計資料 = [
      ["", "", "", "", "小 計", 小計],
      ["", "", "", "", "稅金 (5%)", 稅金],
      ["", "", "", "", "總 計", 總計]
    ];

    sheet.getRange(合計起始, 1, 3, 6).setValues(合計資料);
    
    for (var k = 0; k < 3; k++) {
      sheet.setRowHeight(合計起始 + k, 28);
      sheet.getRange(合計起始 + k, 5).setFontWeight("bold").setHorizontalAlignment("right");
      sheet.getRange(合計起始 + k, 6).setNumberFormat("NT$#,##0").setFontWeight("bold").setHorizontalAlignment("right");
    }
    
    sheet.getRange(合計起始 + 2, 5, 1, 2).setBackground("#e8eaf6")
      .setFontSize(13).setBorder(true, true, true, true, false, false,
        "#1a237e", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    sheet.getRange(12, 1, 品項資料.length + 3, 6)
      .setBorder(true, true, true, true, true, true, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);

    // ============================================================
    // 5. 備註條款
    // ============================================================
    var 備註列 = 合計起始 + 5;
    sheet.getRange(備註列, 1, 1, 6).merge();
    sheet.getRange(備註列, 1).setValue("📌 備註條款：").setFontWeight("bold").setFontColor("#1a237e");
    
    sheet.getRange(備註列 + 1, 1, 1, 6).merge();
    
    var 備註內文 = `1. 報價有效期限自開立日起算 30 天內有效。
2. 付款條件：驗收合格後月結 30 天匯款。
3. 交貨期：正式簽約訂購後 7~14 個工作天交付上線。
4. 以上價格均含基本系統安裝、設定與 16 小時之標準教育訓練。`;

    sheet.getRange(備註列 + 1, 1).setValue(備註內文).setWrap(true).setFontColor("#555555").setFontSize(10);
    sheet.setRowHeight(備註列 + 1, 75);

    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);

    SpreadsheetApp.getUi().alert("✅ 報價單 " + 編號 + " 已成功生成！\n總計金額：NT$ " + 總計.toLocaleString());

  } catch (錯誤) { 
    Logger.log("❌ 發生錯誤：" + 錯誤.message); 
    SpreadsheetApp.getUi().alert("❌ 生成失敗，錯誤訊息：" + 錯誤.message); 
  }
}

/**
 * 初始化「報價資料」來源工作表
 */
function 初始化報價資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("報價資料");
  if (!sheet) sheet = ss.insertSheet("報價資料"); else sheet.clear();

  var 標題 = [["品名/規格", "單位", "說明", "數量", "單價"]];
  var 資料 = [
    ["AI 智慧會議系統（基本版）", "套", "含語音辨識、自動會議紀要", 1, 180000],
    ["智慧文件管理模組", "授權", "OCR + AI 分類，10 人授權", 10, 12000],
    ["自動化報表工具", "授權", "每日/週/月報表自動產生", 5, 8500],
    ["教育訓練", "小時", "現場教育訓練（含教材）", 16, 3000],
    ["系統維護（年約）", "年", "含系統更新與技術支援", 1, 50000],
    ["客製化開發", "人天", "依需求客製功能開發", 10, 8000]
  ];

  sheet.getRange(1, 1, 1, 5).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 5).setValues(資料);
  
  sheet.getRange("A1:E1").setBackground("#283593").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D2:E7").setHorizontalAlignment("right");
  sheet.getRange("E2:E7").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  
  for (var c = 1; c <= 5; c++) {
    sheet.autoResizeColumn(c);
  }

  SpreadsheetApp.getUi().alert("✅ 「報價資料」來源工作表已成功初始化！");
}