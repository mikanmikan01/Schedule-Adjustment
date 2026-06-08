function createMonthlyScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 対象月の設定
  const today = new Date();
  const targetMonth = new Date(today.getFullYear(), today.getMonth() + 3, 1);
  const sheetName = Utilities.formatDate(targetMonth, "Asia/Tokyo", "yyyy年MM月");

  if (ss.getSheetByName(sheetName)) {
    SpreadsheetApp.getUi().alert("エラー：既に『" + sheetName + "』のシートが存在します。");
    return;
  }

  const sheet = ss.insertSheet(sheetName);
  
  // メンバーと時間帯の定義
  const members = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13"];
  const weekdaySlots = ["22:00-"];
  const holidaySlots = ["09:00-12:00", "15:00-18:00", "22:00-"];

  // ヘッダーの生成（〇の数、〇+△の数の2列を追加）
  const headers = ["日付", "時間帯", ...members, "〇の数", "〇+△合計"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 祝日判定とデータ生成
  const calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  const holidayCalendar = CalendarApp.getCalendarById(calendarId);
  const startTime = new Date(targetMonth.getFullYear(), targetMonth.getMonth(), 1);
  const endTime = new Date(targetMonth.getFullYear(), targetMonth.getMonth() + 1, 1);
  const holidays = holidayCalendar.getEvents(startTime, endTime);
  const holidayDates = holidays.map(event => event.getStartTime().getDate());

  const lastDay = new Date(targetMonth.getFullYear(), targetMonth.getMonth() + 1, 0).getDate();
  let rowData = [];
  let dateFontColors = []; 
  const kanjiDays = ["日", "月", "火", "水", "木", "金", "土"];

  for (let day = 1; day <= lastDay; day++) {
    const targetDate = new Date(targetMonth.getFullYear(), targetMonth.getMonth(), day);
    const dayOfWeek = targetDate.getDay();
    const dateStr = Utilities.formatDate(targetDate, "Asia/Tokyo", "MM/dd") + "(" + kanjiDays[dayOfWeek] + ")";
    
    // 休日判定と文字色の決定
    const isHoliday = (dayOfWeek === 0 || dayOfWeek === 6 || holidayDates.includes(day));
    const currentSlots = isHoliday ? holidaySlots : weekdaySlots;
    
    let fontColor = "#000000"; // デフォルト黒
    if (dayOfWeek === 0 || holidayDates.includes(day)) {
      fontColor = "#cc0000"; // 日・祝は赤
    } else if (dayOfWeek === 6) {
      fontColor = "#1155cc"; // 土曜は青
    }

    for (let i = 0; i < currentSlots.length; i++) {
      rowData.push([dateStr, currentSlots[i]]);
      dateFontColors.push([fontColor]);
    }
  }

  // A列・B列への書き込みと日付の文字色適用
  sheet.getRange(2, 1, rowData.length, 2).setValues(rowData);
  sheet.getRange(2, 1, rowData.length, 1).setFontColors(dateFontColors);

  // プルダウンの設定
  const dropdownRange = sheet.getRange(2, 3, rowData.length, members.length);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["〇", "△", "✕"], true)
    .setAllowInvalid(false)
    .build();
  dropdownRange.setDataValidation(rule);

  // 1. 〇の数 (P列)
  const formulaRangeO = sheet.getRange(2, 3 + members.length, rowData.length, 1);
  formulaRangeO.setFormula("=COUNTIF(C2:O2, \"〇\")");
  
  // 2. 〇+△の数 (Q列)
  const formulaRangeP = sheet.getRange(2, 4 + members.length, rowData.length, 1);
  formulaRangeP.setFormula("=COUNTIF(C2:O2, \"〇\") + COUNTIF(C2:O2, \"△\")");

  // 条件付き書式
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("〇").setBackground("#b7e1cd").setRanges([dropdownRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("△").setBackground("#fce8b2").setRanges([dropdownRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("✕").setBackground("#f3f3f3").setFontColor("#b7b7b7").setRanges([dropdownRange]).build());
  sheet.setConditionalFormatRules(rules);

  // メイン表の見栄え最適化
  const fullRange = sheet.getRange(1, 1, rowData.length + 1, headers.length);
  fullRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
  fullRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);

  // 集計列のハイライト
  sheet.getRange(2, headers.length - 1, rowData.length, 2).setBackground("#fff2cc").setFontWeight("bold");


  // ==========================================
  // トップ10抽出サマリー表の構築（S列〜W列に配置）
  // ==========================================
  const summaryStartCol = headers.length + 2; // Q列を空けてR列から開始
  const summaryHeaders = ["順位", "日付", "時間帯", "〇", "〇+△"];
  
  // サマリー用ヘッダーの書き込み
  const sumHeaderRange = sheet.getRange(1, summaryStartCol, 1, summaryHeaders.length);
  sumHeaderRange.setValues([summaryHeaders]);
  sumHeaderRange.setBackground("#1155cc").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  // 順位（1〜10）の書き込み
  const rankValues = [[1],[2],[3],[4],[5],[6],[7],[8],[9],[10]];
  sheet.getRange(2, summaryStartCol, 10, 1).setValues(rankValues).setHorizontalAlignment("center").setVerticalAlignment("middle");

  // QUERY関数による自動抽出（〇が1以上ある日程から、〇が多い順 -> 〇+△が多い順 に並び替え）
  const queryFormula = `=QUERY(A2:Q, "SELECT A, B, P, Q WHERE P > 0 OR Q > 0 ORDER BY P DESC, Q DESC LIMIT 10", 0)`;
  sheet.getRange(2, summaryStartCol + 1).setFormula(queryFormula);

  // サマリー表の枠線装飾
  const summaryDataRange = sheet.getRange(2, summaryStartCol, 10, summaryHeaders.length);
  summaryDataRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(2, summaryStartCol + 1, 10, 4).setHorizontalAlignment("center").setVerticalAlignment("middle");


  // --- 列幅の最終調整 ---
  sheet.autoResizeColumns(1, 2); // A, B列
  for (let i = 3; i <= 2 + members.length; i++) {
    sheet.setColumnWidth(i, 80); // メンバー列
  }
  sheet.setColumnWidth(headers.length - 1, 100); // 〇の数
  sheet.setColumnWidth(headers.length, 100);     // 〇+△合計
  sheet.setColumnWidth(headers.length + 1, 30);  // 余白(R列)
  sheet.setColumnWidth(summaryStartCol, 50);     // 順位
  sheet.setColumnWidth(summaryStartCol + 1, 90); // サマリー:日付
  sheet.setColumnWidth(summaryStartCol + 2, 90); // サマリー:時間帯
}