// ============================================================
// 進階練習：智慧採購分析與供應商評比
// 對應：Session 6（filter、sort、資料摘要）
// ============================================================

/**
 * 供應商綜合評比（篩選 + 排序 + 統計）
 */
function 供應商評比() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("採購紀錄");
    if (!sheet) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];
    var 紀錄 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) obj[標題[j]] = 資料[i][j];
      紀錄.push(obj);
    }

    // 依供應商分組統計
    var 供應商統計 = {};
    紀錄.forEach(function(r) {
      var s = r["供應商"];
      if (!供應商統計[s]) {
        供應商統計[s] = { 訂單數: 0, 總金額: 0, 準時交貨: 0, 退貨次數: 0, 品質分數: [] };
      }
      供應商統計[s].訂單數++;
      供應商統計[s].總金額 += r["金額"];
      if (r["交貨狀態"] === "準時") 供應商統計[s].準時交貨++;
      if (r["退貨"] === "是") 供應商統計[s].退貨次數++;
      供應商統計[s].品質分數.push(r["品質評分"]);
    });

    // 計算綜合評分
    var 排名 = [];
    for (var name in 供應商統計) {
      var s = 供應商統計[name];
      var 準時率 = (s.準時交貨 / s.訂單數 * 100);
      var 退貨率 = (s.退貨次數 / s.訂單數 * 100);
      var 平均品質 = s.品質分數.reduce(function(a, b) { return a + b; }, 0) / s.品質分數.length;

      // 綜合評分 = 品質(40%) + 準時率(30%) + (100-退貨率)(20%) + 價格競爭力(10%)
      var 綜合 = 平均品質 * 0.4 + 準時率 * 0.3 + (100 - 退貨率) * 0.2 + 50 * 0.1;

      排名.push({
        供應商: name, 訂單數: s.訂單數, 總金額: s.總金額,
        準時率: 準時率, 退貨率: 退貨率, 品質: 平均品質, 綜合: 綜合
      });
    }

    // 依綜合評分排序
    排名.sort(function(a, b) { return b.綜合 - a.綜合; });

    // 產生評比報表
    var 評比表 = ss.getSheetByName("供應商評比");
    if (評比表) 評比表.clear(); else 評比表 = ss.insertSheet("供應商評比");

    評比表.getRange("A1").setValue("🏆 供應商綜合評比").setFontSize(16).setFontWeight("bold");
    評比表.getRange("A2").setValue("評比日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd"));

    var 表頭 = [["排名", "供應商", "訂單數", "總金額", "準時率", "退貨率", "品質", "綜合評分", "等級"]];
    評比表.getRange(4, 1, 1, 9).setValues(表頭);
    評比表.getRange(4, 1, 1, 9).setBackground("#1565c0").setFontColor("#fff").setFontWeight("bold");

    排名.forEach(function(r, idx) {
      var 等級 = r.綜合 >= 80 ? "⭐ A" : r.綜合 >= 60 ? "🔵 B" : "🔴 C";
      評比表.getRange(5 + idx, 1, 1, 9).setValues([[
        idx + 1, r.供應商, r.訂單數, r.總金額,
        (r.準時率).toFixed(1) + "%", (r.退貨率).toFixed(1) + "%",
        r.品質.toFixed(1), r.綜合.toFixed(1), 等級
      ]]);
      if (idx === 0) 評比表.getRange(5 + idx, 1, 1, 9).setBackground("#e8f5e9"); // 冠軍綠
    });

    評比表.getRange(5, 4, 排名.length, 1).setNumberFormat("#,##0");
    for (var c = 1; c <= 9; c++) 評比表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 供應商評比完成！冠軍：" + 排名[0].供應商);

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); }
}

/**
 * 採購需求預測（根據歷史消耗趨勢）
 */
function 採購預測() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("採購紀錄");
  if (!sheet) return;

  var 資料 = sheet.getDataRange().getValues();
  var 品項統計 = {};

  for (var i = 1; i < 資料.length; i++) {
    var 品項 = 資料[i][1]; // B: 品項
    var 數量 = 資料[i][3]; // D: 數量
    if (!品項統計[品項]) 品項統計[品項] = [];
    品項統計[品項].push(數量);
  }

  var 預測 = [];
  for (var name in 品項統計) {
    var 歷史 = 品項統計[name];
    
    // 判斷：必須購買超過 3 次 (包含 3 次) 才具有預測價值
    if (歷史.length >= 3) {
      var 平均 = 歷史.reduce(function(a, b) { return a + b; }, 0) / 歷史.length;
      var 建議量 = Math.ceil(平均 * 1.2); // 多 20% 安全庫存

      預測.push("🟢 " + name + " (歷史購買: " + 歷史.length + "次)：月均 " + Math.round(平均) + " → 建議採購 " + 建議量);
    } else {
      // 買太少的，給老闆一個溫馨提示
      預測.push("⚪ " + name + " (歷史購買: " + 歷史.length + "次)：數據不足，暫不預測");
    }
  }

  SpreadsheetApp.getUi().alert("📊 採購預測\n\n" + 預測.join("\n"));
}

function 初始化採購資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("採購紀錄");
  if (!sheet) sheet = ss.insertSheet("採購紀錄"); else sheet.clear();

  var 標題 = [["供應商", "品項", "單價", "數量", "金額", "交貨狀態", "退貨", "品質評分", "日期"]];
  var 供應商 = ["宏達文具", "大同辦公", "金鼎耗材", "永豐科技", "佳能事務"];
  var 品項 = ["A4影印紙", "碳粉匣", "原子筆", "資料夾", "白板筆"];
  var 資料 = [];

  for (var i = 0; i < 40; i++) {
    var s = 供應商[Math.floor(Math.random() * 供應商.length)];
    var p = 品項[Math.floor(Math.random() * 品項.length)];
    var 基礎單價 = [150, 1200, 15, 25, 45][品項.indexOf(p)];
    var 單價 = Math.round(基礎單價 * (1 + (Math.random() * 0.2 - 0.1))); // ±10% 波動
    var 數量 = Math.floor(Math.random() * 50) + 5;
    var 交貨 = Math.random() > 0.2 ? "準時" : "延遲";
    var 退貨 = Math.random() > 0.85 ? "是" : "否";
    var 品質 = Math.floor(Math.random() * 40) + 60;
    var 日期 = new Date(2026, Math.floor(Math.random() * 4), Math.floor(Math.random() * 28) + 1);

    資料.push([s, p, 單價, 數量, 單價 * 數量, 交貨, 退貨, 品質, 日期]);
  }

  sheet.getRange(1, 1, 1, 9).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 9).setValues(資料);
  sheet.getRange("A1:I1").setBackground("#6a1b9a").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("E2:E41").setNumberFormat("#,##0");
  sheet.getRange("I2:I41").setNumberFormat("yyyy/mm/dd");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 9; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 40 筆採購紀錄已建立！");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 智慧採購管理")
    .addItem("📦 初始化採購資料", "初始化採購資料")
    .addItem("🏆 供應商評比", "供應商評比")
    .addItem("📊 採購預測", "採購預測")
    .addSeparator()
    .addItem("📈 價格趨勢分析", "價格趨勢分析")
    .addToUi();
}

// ============================================================
// 進階挑戰 1：價格趨勢分析
// ============================================================

/**
 * 價格趨勢分析（同品項不同月份的價格走勢）
 */
function 價格趨勢分析() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("採購紀錄");
    if (!sheet) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];
    
    // 轉為物件陣列
    var 紀錄 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) obj[標題[j]] = 資料[i][j];
      紀錄.push(obj);
    }

    // 1. 雙重分組統計：先分品項，再分月份
    var 趨勢統計 = {};
    // 找出所有出現過的月份，確保報表欄位完整
    var 所有月份 = {}; 

    紀錄.forEach(function(r) {
      var 品項 = r["品項"];
      var 日期 = new Date(r["日期"]);
      var 月份 = (日期.getMonth() + 1); // 1~12 的純數字
      
      所有月份[月份] = true;

      if (!趨勢統計[品項]) {
        趨勢統計[品項] = {};
      }
      if (!趨勢統計[品項][月份]) {
        趨勢統計[品項][月份] = { 總單價: 0, 筆數: 0 };
      }
      
      趨勢統計[品項][月份].總單價 += r["單價"]; // 累積單價
      趨勢統計[品項][月份].筆數++;
    });

    // 2. 排序所有的月份 (為了畫表格的標題列)
    var 月份排序 = Object.keys(所有月份).map(Number).sort(function(a, b) { return a - b; });

    // 3. 準備寫入試算表的 2D 陣列
    var 報表資料 = [];
    
    // 製作第一行標題列： ["品項", "1月均價", "2月均價", ... "價格趨勢"]
    var 報表標題 = ["品項"];
    月份排序.forEach(function(m) {
      報表標題.push(m + "月均價");
    });
    報表標題.push("價格趨勢");
    報表資料.push(報表標題);

    // 製作內容列 (每項商品一行)
    var 品項清單 = Object.keys(趨勢統計);
    品項清單.forEach(function(品項) {
      var 品項列 = [品項];
      var 第一筆價格 = 0;
      var 最後一筆價格 = 0;

      月份排序.forEach(function(m) {
        var 每月紀錄 = 趨勢統計[品項][m];
        if (每月紀錄) {
          var 均價 = Math.round(每月紀錄.總單價 / 每月紀錄.筆數);
          品項列.push(均價);
          
          if (第一筆價格 === 0) 第一筆價格 = 均價;
          最後一筆價格 = 均價; // 不斷更新，最後留下的就是有購買的最後一個月
        } else {
          品項列.push("-"); // 那個月沒買
        }
      });
      
      // 計算趨勢 (比較第一筆跟最後一筆買的價格)
      if (第一筆價格 > 0 && 最後一筆價格 > 第一筆價格) {
        品項列.push("📈 價格上漲");
      } else if (第一筆價格 > 0 && 最後一筆價格 < 第一筆價格) {
        品項列.push("📉 價格下跌");
      } else {
        品項列.push("➖ 價格持平");
      }

      報表資料.push(品項列);
    });

    // 4. 寫入試算表
    var 趨勢表 = ss.getSheetByName("價格趨勢分析");
    if (趨勢表) 趨勢表.clear(); else 趨勢表 = ss.insertSheet("價格趨勢分析");

    趨勢表.getRange("A1").setValue("📈 價格趨勢交叉分析").setFontSize(16).setFontWeight("bold");
    趨勢表.getRange("A2").setValue("報表產出日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd"));

    趨勢表.getRange(4, 1, 報表資料.length, 報表資料[0].length).setValues(報表資料);
    
    // 化妝
    趨勢表.getRange(4, 1, 1, 報表資料[0].length).setBackground("#ff6f00").setFontColor("#fff").setFontWeight("bold");
    趨勢表.setFrozenRows(4);
    for (var c = 1; c <= 報表資料[0].length; c++) 趨勢表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 價格趨勢分析報表已產生！");

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); }
}
