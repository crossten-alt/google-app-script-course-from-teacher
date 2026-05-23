// ============================================================
// 進階練習：智慧 KPI 儀表板（完整修復 + 進度條優化版）
// 對應：Session 8（條件式格式化、避免重複觸發、多層迴圈）
// ============================================================

/**
 * 建立 KPI 儀表板（條件式格式化 + 即時更新 + 進度條）
 */
function 建立KPI儀表板() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var kpi表 = ss.getSheetByName("KPI 資料");
    if (!kpi表) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    var sheet = ss.getSheetByName("KPI 儀表板");
    if (sheet) sheet.clear(); else sheet = ss.insertSheet("KPI 儀表板", 0);

    // 標題 (配合進度條欄位，將範圍擴展至 A1:I1)
    sheet.getRange("A1:I1").merge();
    sheet.getRange("A1").setValue("📊 2026 Q2 KPI 儀表板")
      .setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center")
      .setBackground("#0d47a1").setFontColor("#fff");
    sheet.setRowHeight(1, 50);
    
    // 修正原本的語法錯誤，正確匯入時間
    sheet.getRange("A2").setValue("最後更新：" +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"))
      .setFontColor("#666");

    // 讀取 KPI 資料
    var 資料 = kpi表.getDataRange().getValues();
    
    // === KPI 卡片區 ===
    var 卡片資料 = _計算KPI摘要(資料);
    var 卡片 = [
      { 名稱: "營收達成率", 值: 卡片資料.營收達成率.toFixed(1) + "%", 顏色: "#1a73e8" },
      { 名稱: "客戶滿意度", 值: 卡片資料.平均滿意度.toFixed(1) + "%", 顏色: "#34a853" },
      { 名稱: "逾期任務", 值: 卡片資料.逾期數 + " 項", 顏色: 卡片資料.逾期數 > 3 ? "#ea4335" : "#fbbc04" },
      { 名稱: "員工績效均", 值: 卡片資料.平均績效.toFixed(0) + " 分", 顏色: "#9c27b0" }
    ];
    
    for (var k = 0; k < 卡片.length; k++) {
      var col = k * 2 + 1;
      sheet.getRange(4, col, 1, 2).merge();
      sheet.getRange(4, col).setValue(卡片[k].名稱)
        .setHorizontalAlignment("center").setFontWeight("bold").setFontColor("#fff")
        .setBackground(卡片[k].顏色);
      sheet.getRange(5, col, 1, 2).merge();
      sheet.getRange(5, col).setValue(卡片[k].值)
        .setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center")
        .setBackground(卡片[k].顏色 + "22");
      sheet.setRowHeight(5, 50);
    }

    // === 詳細 KPI 表格 ===
    sheet.getRange("A8:I8").setValues([["部門", "員工", "營收目標", "實際營收", "達成率", "進度條", "客戶滿意度", "績效分數", "狀態"]]);
    sheet.getRange("A8:I8").setBackground("#263238").setFontColor("#fff").setFontWeight("bold");

    var 輸出資料 = [];
    var 背景顏色陣列 = [];

    for (var i = 1; i < 資料.length; i++) {
      var 達成率 = 資料[i][3] > 0 ? (資料[i][4] / 資料[i][3]) : 0;
      var 狀態 = 達成率 >= 1 ? "🟢 達標" : 達成率 >= 0.8 ? "🟡 接近" : "🔴 未達";

      var 每格比例 = 0.1;
      var 實心格子數 = Math.min(10, Math.max(0, Math.floor(達成率 / 每格比例)));
      var 空心格子數 = 10 - 實心格子數;
      var 進度條 = "█".repeat(實心格子數) + "░".repeat(空心格子數);
      
      if (資料[i][3] === 0) {
        進度條 = "➖";
        狀態 = "⚪ 內勤";
      }

      輸出資料.push([
        資料[i][0], 資料[i][1], 資料[i][3], 資料[i][4],
        達成率, 進度條, 資料[i][5], 資料[i][6], 狀態
      ]);

      var rowColor = (i % 2 === 0) ? "#eceff1" : "#ffffff";
      背景顏色陣列.push([rowColor, rowColor, rowColor, rowColor, rowColor, rowColor, rowColor, rowColor, rowColor]);
    }

    var 資料筆數 = 資料.length - 1;
    var 資料範圍 = sheet.getRange(9, 1, 資料筆數, 9);
    資料範圍.setValues(輸出資料);
    資料範圍.setBackgrounds(背景顏色陣列);

    sheet.getRange(9, 3, 資料筆數, 2).setNumberFormat("#,##0");
    sheet.getRange(9, 5, 資料筆數, 1).setNumberFormat("0.0%");
    sheet.getRange(9, 6, 資料筆數, 1).setHorizontalAlignment("center").setFontColor("#00796b");
    sheet.getRange(9, 7, 資料筆數, 1).setNumberFormat("0.0%");

    // === 條件式格式化 ===
    sheet.clearConditionalFormatRules();
    var rules = [];
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(1)
      .setBackground("#c8e6c9").setFontColor("#1b5e20")
      .setRanges([sheet.getRange(9, 5, 資料筆數, 1)]).build());
      
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.8)
      .setBackground("#ffcdd2").setFontColor("#b71c1c")
      .setRanges([sheet.getRange(9, 5, 資料筆數, 1)]).build());
      
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(70)
      .setBackground("#ffcdd2").setBold(true)
      .setRanges([sheet.getRange(9, 8, 資料筆數, 1)]).build());
      
    sheet.setConditionalFormatRules(rules);

    for (var c = 1; c <= 9; c++) sheet.autoResizeColumn(c);
    sheet.setFrozenRows(8);

    SpreadsheetApp.getUi().alert("✅ KPI 儀表板已建立（含動態進度條）！");
  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); }
}

/**
 * 計算 KPI 摘要資料
 */
function _計算KPI摘要(資料) {
  var 總目標 = 0, 總實際 = 0, 滿意度總 = 0, 績效總 = 0, 逾期 = 0;
  for (var i = 1; i < 資料.length; i++) {
    總目標 += 資料[i][3];
    總實際 += 資料[i][4];
    滿意度總 += 資料[i][5];
    績效總 += 資料[i][6];
    if (資料[i][3] > 0 && 資料[i][4] / 資料[i][3] < 0.8) 逾期++;
  }
  var n = 資料.length - 1;
  return {
    營收達成率: 總目標 > 0 ? (總實際 / 總目標) * 100 : 0,
    平均滿意度: (滿意度總 / n) * 100,
    平均績效: 績效總 / n,
    逾期數: 逾期
  };
}

/**
 * 試算表更動即時觸發
 */
function onEdit(e) {
  if (e.source.getActiveSheet().getName() === "KPI 資料") {
    建立KPI儀表板();
  }
}

/**
 * 初始化測試資料
 */
function 初始化KPI資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("KPI 資料");
  if (!sheet) sheet = ss.insertSheet("KPI 資料");
  else sheet.clear();

  var 標題 = [["部門", "員工", "職位", "營收目標", "實際營收", "客戶滿意度", "績效分數"]];
  var 資料 = [
    ["業務部", "王小明", "業務專員", 500000, 520000, 0.92, 88],
    ["業務部", "黃志偉", "業務主管", 800000, 680000, 0.85, 75],
    ["行銷部", "李小華", "行銷主管", 300000, 310000, 0.88, 82],
    ["行銷部", "吳雅琪", "行銷專員", 200000, 150000, 0.78, 65],
    ["研發部", "陳大文", "工程師", 0, 0, 0.95, 92],
    ["研發部", "劉家豪", "資深工程師", 0, 0, 0.90, 90],
    ["人資部", "張美玲", "人資專員", 0, 0, 0.82, 78],
    ["財務部", "林小芬", "會計", 0, 0, 0.88, 85],
    ["業務部", "周建國", "業務專員", 400000, 280000, 0.72, 58],
    ["客服部", "許文馨", "客服主管", 0, 0, 0.95, 91]
  ];

  sheet.getRange(1, 1, 1, 7).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 7).setValues(資料);
  sheet.getRange("A1:G1").setBackground("#263238").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("D2:E11").setNumberFormat("#,##0");
  sheet.getRange("F2:F11").setNumberFormat("0%");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 7; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ KPI 資料已建立！請執行「建立 KPI 儀表板」。");
}

/**
 * 建立自訂工具選單
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 KPI 儀表板")
    .addItem("📦 初始化 KPI 資料", "初始化KPI資料")
    .addItem("📊 建立 KPI 儀表板", "建立KPI儀表板")
    .addToUi();
}