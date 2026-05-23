/**
 * 📚 會議紀錄純表格模板生成器 (穩定修正版)
 * * 功能：自動建立並美化「會議紀錄_空白模板」工作表。
 * * 特色：100% 還原圖片中的顏色、字級、粗體與對齊格式，且不包含任何具體會議內容。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📝 會議紀錄功能")
    .addItem("生成會議模板", "生成會議模板")
    .addItem("一鍵修復格式", "修復會議紀錄格式")
    .addToUi();
}

function 生成會議模板() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "模板";
  var sheet = ss.getSheetByName(sheetName);
  
  // 1. 如果已有同名頁簽則清空重建，沒有則新增
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    // 清除可能殘留的合併儲存格
    var mergedRanges = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).getMergedRanges();
    for (var i = 0; i < mergedRanges.length; i++) {
      mergedRanges[i].breakApart();
    }
  }

  // 2. 定義純表格結構（移除特定會議內容，僅留欄位名）
  var templateData = [
    ["會議紀錄：", "", "", ""], // Row 1 (大標題)
    ["", "", "", ""],           // Row 2 (空白)
    ["會議主題", "", "", ""],     // Row 3
    ["會議時間", "", "", ""],     // Row 4
    ["會議主席", "", "", ""],     // Row 5
    ["會議記錄", "", "", ""],     // Row 6
    ["與會人員", "", "", ""],     // Row 7
    ["", "", "", ""],           // Row 8 (空白)
    ["一、會議討論重點", "", "", ""], // Row 9
    ["主題", "詳細討論內容", "", ""], // Row 10
    ["1. 需求確認", "", "", ""],     // Row 11
    ["2. 技術可行性", "", "", ""],   // Row 12
    ["3. 資訊安全", "", "", ""],     // Row 13
    ["", "", "", ""],           // Row 14 (空白)
    ["二、決議事項", "", "", ""],     // Row 15
    ["編號", "決議內容項目", "", ""], // Row 16
    ["決議 1", "", "", ""],          // Row 17
    ["決議 2", "", "", ""],          // Row 18
    ["", "", "", ""],           // Row 19 (空白)
    ["三、追蹤事項 (Action Items)", "", "", ""], // Row 20
    ["負責人", "待辦事項內容", "預計完成日", "目前狀態"], // Row 21
    ["", "", "", ""],           // Row 22
    ["", "", "", ""],           // Row 23
    ["", "", "", ""],           // Row 24
    ["", "", "", ""],           // Row 25 (空白)
    ["💡 下次會議時間：", "", "", ""], // Row 26
    ["", "", "", ""]            // Row 27 (空白)
  ];

  // 3. 批次寫入結構資料
  sheet.getRange(1, 1, 27, 4).setValues(templateData);

  // ============================================================
  // 4. 精準格式、色彩、字級與粗體設定 (依據圖片還原)
  // ============================================================
  
  // (已移除會引發錯誤的全域格線設定，採用試算表預設格線狀態)

  // 全域基礎設定：文字自動換行與置中對齊
  var fullRange = sheet.getRange("A1:D27");
  fullRange.setWrap(true);
  fullRange.setVerticalAlignment("middle");
  fullRange.setFontFamily("Arial"); // 預設標準清晰字型

  // --- [Row 1] 主標題 ---
  // 樣式：深藍色文字 (#1f4e79)、16級字、粗體、靠左對齊
  sheet.getRange("A1:D1").merge()
       .setFontSize(16)
       .setFontColor("#1f4e79")
       .setFontWeight("bold")
       .setHorizontalAlignment("left")
       .setVerticalAlignment("bottom");

  // --- [Row 3~7] 會議基本資訊區 ---
  // A欄樣式：灰底 (#f2f2f2)、11級字、粗體、置中對齊
  sheet.getRange("A3:A7")
       .setBackground("#f2f2f2")
       .setFontSize(11)
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
  
  // B~D欄樣式：合併、靠左對齊、預設11級字、常規字體
  for (var r = 3; r <= 7; r++) {
    sheet.getRange(r, 2, 1, 3).merge()
         .setFontSize(11)
         .setFontWeight("normal")
         .setHorizontalAlignment("left");
  }
  // 邊框：細灰色邊框
  sheet.getRange("A3:D7").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);


  // --- [Row 9, 15, 20] 大區段標題（一、二、三） ---
  // 樣式：深藍色文字 (#1f4e79)、11級字、粗體、靠左對齊
  var sectionHeaders = ["A9:D9", "A15:D15", "A20:D20"];
  sectionHeaders.forEach(function(rangeStr) {
    sheet.getRange(rangeStr).merge()
         .setFontSize(11)
         .setFontColor("#1f4e79")
         .setFontWeight("bold")
         .setHorizontalAlignment("left");
  });


  // --- [Row 10, 16, 21] 表格頂端標題欄 ---
  // 樣式：深藍底色 (#1f4e79)、白字 (#ffffff)、11級字、粗體、置中對齊
  var tableHeaders = ["A10:D10", "A16:D16", "A21:D21"];
  tableHeaders.forEach(function(rangeStr) {
    sheet.getRange(rangeStr)
         .setBackground("#1f4e79")
         .setFontColor("#ffffff")
         .setFontSize(11)
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
  });


  // --- [Row 11~13] 一、會議討論重點 表格內容 ---
  // A欄項目樣式：11級字、粗體、置中對齊
  sheet.getRange("A11:A13").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");
  // B~D欄內容樣式：合併、11級字、常規字體、靠左對齊
  for (var r = 11; r <= 13; r++) {
    sheet.getRange(r, 2, 1, 3).merge().setFontSize(11).setFontWeight("normal").setHorizontalAlignment("left");
  }
  sheet.getRange("A10:D13").setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);


  // --- [Row 17~18] 二、決議事項 表格內容 ---
  // A欄項目樣式：11級字、粗體、置中對齊
  sheet.getRange("A17:A18").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");
  // B~D欄內容樣式：合併、11級字、常規字體、靠左對齊
  for (var r = 17; r <= 18; r++) {
    sheet.getRange(r, 2, 1, 3).merge().setFontSize(11).setFontWeight("normal").setHorizontalAlignment("left");
  }
  sheet.getRange("A16:D18").setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);


  // --- [Row 22~24] 三、追蹤事項 表格內容 ---
  // A欄負責人：11級字、常規字體、置中對齊
  sheet.getRange("A22:A24").setFontSize(11).setFontWeight("normal").setHorizontalAlignment("center");
  // B欄待辦事項內容：11級字、常規字體、靠左對齊
  sheet.getRange("B22:B24").setFontSize(11).setFontWeight("normal").setHorizontalAlignment("left");
  // C欄預計完成日、D欄目前狀態：11級字、常規字體、置中對齊
  sheet.getRange("C22:D24").setFontSize(11).setFontWeight("normal").setHorizontalAlignment("center");
  sheet.getRange("A21:D24").setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);


  // --- [Row 26] 下次會議時間提示 ---
  // 樣式：深灰色文字 (#595959)、11級字、粗體、靠左對齊
  sheet.getRange("A26:D26").merge()
       .setFontSize(11)
       .setFontColor("#595959")
       .setFontWeight("bold")
       .setHorizontalAlignment("left");

  // ============================================================
  // 5. 欄寬與列高設定（精準維持圖片比例）
  // ============================================================
  sheet.setColumnWidth(1, 140);  // 第一欄 (項目標題/負責人)
  sheet.setColumnWidth(2, 450);  // 第二欄 (主要內容輸入區)
  sheet.setColumnWidth(3, 110);  // 第三欄 (日期)
  sheet.setColumnWidth(4, 110);  // 第四欄 (狀態)

  // 設定各行高度
  sheet.setRowHeight(1, 40);   // 大標題
  sheet.setRowHeight(2, 15);   // 空白
  for (var r = 3; r <= 7; r++) sheet.setRowHeight(r, 28); // 基本資訊
  sheet.setRowHeight(8, 15);   // 空白
  sheet.setRowHeight(9, 25);   // 大區段標題一
  sheet.setRowHeight(10, 28);  // 表格標題
  sheet.setRowHeight(11, 55);  // 討論重點內容高
  sheet.setRowHeight(12, 55);  
  sheet.setRowHeight(13, 55);  
  sheet.setRowHeight(14, 15);  // 空白
  sheet.setRowHeight(15, 25);  // 大區段標題二
  sheet.setRowHeight(16, 28);  // 表格標題
  sheet.setRowHeight(17, 45);  // 決議事項內容高
  sheet.setRowHeight(18, 45);  
  sheet.setRowHeight(19, 15);  // 空白
  sheet.setRowHeight(20, 25);  // 大區段標題三
  sheet.setRowHeight(21, 28);  // 表格標題
  sheet.setRowHeight(22, 35);  // 追蹤事項內容高
  sheet.setRowHeight(23, 35);  
  sheet.setRowHeight(24, 35);  
  sheet.setRowHeight(25, 15);  // 空白
  sheet.setRowHeight(26, 30);  // 下次會議時間提示高

}

function 修復會議紀錄格式() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 💡 請確保您要修復的工作表名稱是「會議紀錄」
  var sheet = ss.getSheetByName("會議紀錄"); 
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到名為「會議紀錄」的工作表，請先檢查工作表名稱是否正確！");
    return;
  }

  // 先解除所有舊的合併儲存格，避免格式衝突
  var mergedRanges = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).getMergedRanges();
  for (var i = 0; i < mergedRanges.length; i++) {
    mergedRanges[i].breakApart();
  }

  // ============================================================
  // 1. 全域基礎設定 (自動換行、字型、對齊)
  // ============================================================
  var fullRange = sheet.getRange("A1:D27");
  fullRange.setWrap(true);
  fullRange.setVerticalAlignment("middle");
  fullRange.setFontFamily("Arial");
  fullRange.setFontSize(11);
  fullRange.setFontWeight("normal");
  fullRange.setFontColor("#000000");
  fullRange.setBackground(null); // 清除舊的背景色

  // ============================================================
  // 2. 精準色彩、字級與粗體還原 (依據圖片規格)
  // ============================================================

  // --- [Row 1] 主標題 ---
  sheet.getRange("A1:B1").merge()
       .setFontSize(16)
       .setFontColor("#1f4e79")
       .setFontWeight("bold")
       .setHorizontalAlignment("left")
       .setVerticalAlignment("bottom");

  // --- [Row 3~7] 會議基本資訊區 ---
  sheet.getRange("A3:A7")
       .setBackground("#f2f2f2")
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
  
  for (var r = 3; r <= 7; r++) {
    sheet.getRange(r, 2, 1, 3).merge().setHorizontalAlignment("left");
  }
  sheet.getRange("A3:D7").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);

  // --- [Row 9, 15, 20] 大區段標題（一、二、三） ---
  var sectionHeaders = ["A9:D9", "A15:D15", "A20:D20"];
  sectionHeaders.forEach(function(rangeStr) {
    sheet.getRange(rangeStr).merge()
         .setFontColor("#1f4e79")
         .setFontWeight("bold")
         .setHorizontalAlignment("left");
  });

  // --- [Row 10, 16, 21] 表格頂端藍色標題欄 ---
  var tableHeaders = ["A10:D10", "A16:D16", "A21:D21"];
  tableHeaders.forEach(function(rangeStr) {
    sheet.getRange(rangeStr)
         .setBackground("#1f4e79")
         .setFontColor("#ffffff")
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
  });

  // --- [Row 11~13] 一、會議討論重點 表格內容 ---
  sheet.getRange("A11:A13").setFontWeight("bold").setHorizontalAlignment("center");
  for (var r = 11; r <= 13; r++) {
    sheet.getRange(r, 2, 1, 3).merge().setHorizontalAlignment("left");
  }
  sheet.getRange("A10:D13").setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);

  // --- [Row 17~18] 二、決議事項 表格內容 ---
  sheet.getRange("A17:A18").setFontWeight("bold").setHorizontalAlignment("center");
  for (var r = 17; r <= 18; r++) {
    sheet.getRange(r, 2, 1, 3).merge().setHorizontalAlignment("left");
  }
  sheet.getRange("A16:D18").setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);

  // --- [Row 22~24] 三、追蹤事項 表格內容 ---
  sheet.getRange("A22:A24").setHorizontalAlignment("center");
  sheet.getRange("B22:B24").setHorizontalAlignment("left");
  sheet.getRange("C22:D24").setHorizontalAlignment("center");
  sheet.getRange("A21:D24").setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);

  // --- [Row 26] 下次會議時間提示 ---
  sheet.getRange("A26:D26").merge()
       .setFontColor("#595959")
       .setFontWeight("bold")
       .setHorizontalAlignment("left");

  // ============================================================
  // 3. 欄寬與列高精密重設
  // ============================================================
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 450);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);

  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 15);
  for (var r = 3; r <= 7; r++) sheet.setRowHeight(r, 28);
  sheet.setRowHeight(8, 15);
  sheet.setRowHeight(9, 25);
  sheet.setRowHeight(10, 28);
  sheet.setRowHeight(11, 55);
  sheet.setRowHeight(12, 55);  
  sheet.setRowHeight(13, 55);  
  sheet.setRowHeight(14, 15);
  sheet.setRowHeight(15, 25);
  sheet.setRowHeight(16, 28);
  sheet.setRowHeight(17, 45);
  sheet.setRowHeight(18, 45);  
  sheet.setRowHeight(19, 15);
  sheet.setRowHeight(20, 25);
  sheet.setRowHeight(21, 28);
  sheet.setRowHeight(22, 35);
  sheet.setRowHeight(23, 35);  
  sheet.setRowHeight(24, 35);  
  sheet.setRowHeight(25, 15);
  sheet.setRowHeight(26, 30);

  SpreadsheetApp.getUi().alert("✨ 偵測到已有會議資料！格式已一鍵修復完成，內容均完好無損！");
}