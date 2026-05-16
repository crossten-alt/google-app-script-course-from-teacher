// ============================================================
// 進階練習：智慧任務派工與自動提醒系統
// 對應：Session 5（資料結構化、function & return、觸發器）
// ============================================================

/**
 * 讀取任務資料並結構化為物件陣列
 */
function 取得結構化任務() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("任務看板");
  if (!sheet) return [];
  var 資料 = sheet.getDataRange().getValues();
  var 標題 = 資料[0];
  var 任務 = [];
  for (var i = 1; i < 資料.length; i++) {
    var obj = {};
    for (var j = 0; j < 標題.length; j++) obj[標題[j]] = 資料[i][j];
    任務.push(obj);
  }
  return 任務;
}

/**
 * 智慧派工：根據員工目前負擔自動分配任務
 * 教學重點：資料結構化、陣列操作 (filter/sort)、批次寫回效能優化
 */
function 智慧派工() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ==========================================
    // 步驟 1：讀取並準備資料
    // ==========================================
    var 任務表 = ss.getSheetByName("任務看板");
    var 員工表 = ss.getSheetByName("員工負擔");
    if (!任務表 || !員工表) { 
      SpreadsheetApp.getUi().alert("❌ 請先初始化資料"); 
      return; 
    }

    var 任務清單 = 取得結構化任務();
    var 員工資料 = 員工表.getDataRange().getValues();
    var 員工清單 = [];
    
    // 將員工資料轉為結構化物件，並記錄原本所在的「列號」方便更新
    for (var i = 1; i < 員工資料.length; i++) {
      員工清單.push({
        列號: i + 1, 
        姓名: 員工資料[i][0], 
        部門: 員工資料[i][1],
        技能: 員工資料[i][2], // 👈 新增：抓取具備技能
        目前任務數: 員工資料[i][3], // 索引往後推
        最大承載: 員工資料[i][4]  // 索引往後推
      });
    }

    // ==========================================
    // 步驟 2：篩選出需要指派的任務
    // ==========================================
    var 待指派任務 = 任務清單.filter(function(t) { 
      return !t["負責人"] || t["負責人"] === "待指派"; 
    });

    if (待指派任務.length === 0) {
      SpreadsheetApp.getUi().alert("📋 目前沒有待指派的任務！");
      return;
    }

    var 指派結果 = [];
    // 取得原始二維陣列，以便在記憶體中更新後「批次寫回」，提升效能
    var 任務陣列 = 任務表.getDataRange().getValues(); 

    // ==========================================
    // 步驟 3：依據邏輯進行派工 (核心演算法)
    // ==========================================
    待指派任務.forEach(function(任務) {
      var 任務名稱 = 任務["任務名稱"];
      var 需求部門 = 任務["部門"];

      var 需求技能 = 任務["所需技能"]; // 👈 抓取任務需要的技能

      // 🔍 階段性過濾，抓出真正的失敗原因！
      // 1. 先找同部門的人
      var 同部門員工 = 員工清單.filter(function(員工) { return 員工.部門 === 需求部門; });
      if (同部門員工.length === 0) {
        指派結果.push("⚠️ " + 任務名稱 + " → ❌ 失敗：找不到「" + 需求部門 + "」的人員");
        return; // 跳過這個任務
      }

      // 2. 再找技能符合的人
      var 技能符合員工 = 同部門員工.filter(function(員工) {
        return !需求技能 || (員工.技能 && 員工.技能.indexOf(需求技能) !== -1);
      });
      if (技能符合員工.length === 0) {
        指派結果.push("⚠️ " + 任務名稱 + " → ❌ 失敗：「" + 需求部門 + "」無人具備『" + 需求技能 + "』技能");
        return; // 跳過這個任務
      }

      // 3. 最後看誰還有空檔 (沒滿載)
      var 可用員工 = 技能符合員工.filter(function(員工) {
        return 員工.目前任務數 < 員工.最大承載;
      });
      if (可用員工.length === 0) {
        指派結果.push("⚠️ " + 任務名稱 + " → ❌ 失敗：「" + 需求部門 + "」具備技能的人皆已滿載 (太血汗啦！)");
        return; // 跳過這個任務
      }

      // 條件 C：優先分配給「目前任務數」最少的員工（最少負擔優先）
      可用員工.sort(function(a, b) { 
        return a.目前任務數 - b.目前任務數; 
      });
      
      var 最佳人選 = 可用員工[0]; // 排序後的第一筆就是任務最少的

      // 執行指派
      最佳人選.目前任務數++; // 增加員工負擔
      指派結果.push("✅ " + 任務名稱 + " → 👤 " + 最佳人選.姓名);

      // 同步更新任務表的二維陣列（注意：第 3 欄「負責人」的 index 是 2）
      for (var r = 1; r < 任務陣列.length; r++) {
        if (任務陣列[r][0] === 任務名稱) {
          任務陣列[r][2] = 最佳人選.姓名; 
          break;
        }
      }
    });

    // ==========================================
    // 步驟 4：批次寫回試算表（最佳效能做法）
    // ==========================================
    // 1. 寫回任務表 (一次更新整個範圍)
    任務表.getRange(1, 1, 任務陣列.length, 任務陣列[0].length).setValues(任務陣列);

    // 🎨 新增：幫「待指派」的儲存格上色 (黃底黑字)
    for (var r = 1; r < 任務陣列.length; r++) {
      if (任務陣列[r][2] === "待指派" || !任務陣列[r][2]) {
        // 第 3 欄 (C欄負責人)，改成黃底 (#fff000) 黑字
        任務表.getRange(r + 1, 3).setBackground("#fff000").setFontColor("#000000");
      } else {
        // 已經指派出去的，恢復預設的白底黑字
        任務表.getRange(r + 1, 3).setBackground("#ffffff").setFontColor("#000000");
      }
    }

    // 2. 寫回員工負擔表 (更新 D 欄的目前任務數，因為插入了技能欄位)
    for (var k = 0; k < 員工清單.length; k++) {
      員工表.getRange(員工清單[k].列號, 4).setValue(員工清單[k].目前任務數);
    }

    SpreadsheetApp.getUi().alert("🎯 智慧派工完成！\n\n" + 指派結果.join("\n"));

  } catch (錯誤) { 
    Logger.log("❌ 發生錯誤：" + 錯誤.message); 
  }
}

/**
 * 逾期任務自動提醒
 */
function 逾期提醒() {
  var 任務 = 取得結構化任務();
  var 今天 = new Date();
  今天.setHours(0, 0, 0, 0);

  var 逾期 = 任務.filter(function(t) {
    if (t["狀態"] === "已完成") return false;
    var 截止 = new Date(t["截止日期"]);
    return 截止 < 今天;
  });

  var 即將到期 = 任務.filter(function(t) {
    if (t["狀態"] === "已完成") return false;
    var 截止 = new Date(t["截止日期"]);
    var 差天 = Math.ceil((截止 - 今天) / 86400000);
    return 差天 >= 0 && 差天 <= 3;
  });

  // 更新狀態欄
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("任務看板");
  var 資料 = sheet.getDataRange().getValues();
  for (var i = 1; i < 資料.length; i++) {
    if (資料[i][5] === "已完成") continue;
    var d = new Date(資料[i][4]);
    if (d < 今天) {
      sheet.getRange(i + 1, 6).setValue("🔴 已逾期");
      sheet.getRange(i + 1, 1, 1, 7).setBackground("#ffebee");
    } else if (Math.ceil((d - 今天) / 86400000) <= 3) {
      sheet.getRange(i + 1, 6).setValue("🟡 即將到期");
      sheet.getRange(i + 1, 1, 1, 7).setBackground("#fff8e1");
    }
  }

  var msg = "📋 任務提醒\n🔴 逾期：" + 逾期.length + " 筆\n🟡 即將到期：" + 即將到期.length + " 筆";
  Logger.log(msg);
  if (逾期.length > 0 || 即將到期.length > 0) {
    SpreadsheetApp.getUi().alert(msg);
  }
}

function 初始化任務資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 任務看板
  var t = ss.getSheetByName("任務看板");
  if (!t) t = ss.insertSheet("任務看板"); else t.clear();
  t.getRange(1, 1, 1, 8).setValues([["任務名稱", "優先級", "負責人", "部門", "所需技能", "截止日期", "狀態", "進度(%)"]]);
  var 任務 = [
    ["系統效能優化", "高", "待指派", "研發部", "後端", new Date(2026, 4, 10), "進行中", 30],
    ["Q2 行銷企劃", "高", "待指派", "行銷部", "企劃", new Date(2026, 4, 15), "未開始", 0],
    ["員工教育訓練", "中", "張美玲", "人資部", "培訓", new Date(2026, 4, 20), "進行中", 50],
    ["客戶滿意度調查", "中", "待指派", "業務部", "客服", new Date(2026, 3, 28), "未開始", 0],
    ["新產品原型開發", "高", "陳大文", "研發部", "前端", new Date(2026, 4, 5), "進行中", 70],
    ["合約審閱", "低", "待指派", "法務", "法律", new Date(2026, 4, 30), "未開始", 0],
    ["辦公室搬遷規劃", "中", "林小芬", "總務", "行政", new Date(2026, 5, 15), "進行中", 20],
    ["ERP 系統升級", "高", "待指派", "研發部", "資料庫", new Date(2026, 4, 25), "未開始", 0]
  ];
  t.getRange(2, 1, 任務.length, 8).setValues(任務);
  t.getRange("A1:H1").setBackground("#4527a0").setFontColor("#fff").setFontWeight("bold");
  t.getRange("F2:F9").setNumberFormat("yyyy/mm/dd"); // shifted to F
  t.setFrozenRows(1);
  for (var c = 1; c <= 8; c++) t.autoResizeColumn(c);

  // 員工負擔
  var e = ss.getSheetByName("員工負擔");
  if (!e) e = ss.insertSheet("員工負擔"); else e.clear();
  e.getRange(1, 1, 1, 5).setValues([["姓名", "部門", "具備技能", "目前任務數", "最大承載"]]);
  var 員工 = [
    ["王小明", "業務部", "客服,談判", 2, 5], ["李小華", "行銷部", "企劃,文案", 3, 5],
    ["張美玲", "人資部", "培訓,招募", 1, 4], ["陳大文", "研發部", "前端,後端", 4, 6],
    ["林小芬", "總務", "行政,採購", 1, 4], ["劉家豪", "研發部", "後端,資料庫", 2, 5]
  ];
  e.getRange(2, 1, 員工.length, 5).setValues(員工);
  e.getRange("A1:E1").setBackground("#00695c").setFontColor("#fff").setFontWeight("bold");
  e.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("✅ 任務資料已建立！（含 4 筆「待指派」任務）");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 智慧派工系統")
    .addItem("📦 初始化任務資料", "初始化任務資料")
    .addItem("🎯 智慧派工", "智慧派工")
    .addItem("⏰ 逾期提醒", "逾期提醒")
    .addToUi();
}
