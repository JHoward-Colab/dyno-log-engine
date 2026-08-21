// =========================================================================
// 📈 SERIAL HISTORICAL ANALYZER (History.js)
// =========================================================================

function renderSerialHistory(e) {
  var ss = e ? e.source : SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SERIAL_HISTORY_VIEWER || "Serial_History_Viewer");
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);

  if (!historySheet || !logSheet) return;

  var searchRange = historySheet.getRange(CONFIG.SERIAL_HISTORY_VIEWER.RANGES.SERIAL_SEARCH_INPUT);
  var targetSerial = String(searchRange.getValue()).trim();

  // Clear previous output
  historySheet.getRange(CONFIG.SERIAL_HISTORY_VIEWER.RANGES.CLEAR_RESULTS_RANGE).clearContent().setBackground(null).setFontWeight("normal");
  historySheet.getRange("B5:E5").clearContent();

  if (!targetSerial || targetSerial === "undefined" || targetSerial === "") return;

  var logData = logSheet.getDataRange().getValues();
  if (logData.length < 2) return;

  var hMap = buildHeaderMap(logData[0]);
  var matchingRows = [];
  var detectedBaseModel = "";

  for (var r = 1; r < logData.length; r++) {
    var row = logData[r];
    var logSerial = String(row[hMap.trueSerial] || "").trim();

    if (isSerialMatch(targetSerial, logSerial)) {
      if (!detectedBaseModel && row[hMap.baseModel]) {
        detectedBaseModel = String(row[hMap.baseModel]).trim();
      }

      var timestamp   = row[hMap.timestamp];
      var serial      = logSerial;
      var rf          = safeAbsNum(row[hMap.rodForce]);
      var c1          = safeAbsNum(row[hMap.comp1]);
      var r1          = safeAbsNum(row[hMap.reb1]);
      var sl1         = row[hMap.slope1];
      var la1         = row[hMap.loopArea1];
      var c2          = safeAbsNum(row[hMap.comp2]);
      var r2          = safeAbsNum(row[hMap.reb2]);
      var la2         = row[hMap.loopArea2];
      var t1Stat      = row[hMap.test1Status];
      var t2Stat      = row[hMap.test2Status];
      var overStat    = String(row[hMap.overallStatus] || "").toUpperCase();
      var diag        = row[hMap.diagnostics];
      var evalAction  = row[hMap.evaluationAction];
      var engComments = row[hMap.engComments];

      matchingRows.push([
        timestamp, serial, rf, c1, r1, sl1, la1, c2, r2, la2,
        t1Stat, t2Stat, overStat, diag, evalAction, engComments
      ]);
    }
  }

  var totalRuns = matchingRows.length;
  if (totalRuns === 0) {
    historySheet.getRange(CONFIG.SERIAL_HISTORY_VIEWER.RANGES.BASE_MODEL).setValue("❌ Serial Not Found");
    return;
  }

  // Sort newest test runs first
  matchingRows.sort(function(a, b) {
    var dA = a[0] instanceof Date ? a[0].getTime() : 0;
    var dB = b[0] instanceof Date ? b[0].getTime() : 0;
    return dB - dA;
  });

  // Render KPI Summary
  var latestRun = matchingRows[0];
  var latestDateStr = latestRun[0] instanceof Date ? Utilities.formatDate(latestRun[0], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : String(latestRun[0]);

  historySheet.getRange(CONFIG.SERIAL_HISTORY_VIEWER.RANGES.TOTAL_RUNS).setValue(totalRuns + (totalRuns > 1 ? " Runs (Retested)" : " Run"));
  historySheet.getRange(CONFIG.SERIAL_HISTORY_VIEWER.RANGES.BASE_MODEL).setValue(detectedBaseModel || "N/A");
  historySheet.getRange(CONFIG.SERIAL_HISTORY_VIEWER.RANGES.LATEST_STATUS).setValue(latestRun[12]);
  historySheet.getRange(CONFIG.SERIAL_HISTORY_VIEWER.RANGES.LAST_TEST_DATE).setValue(latestDateStr);

  // Render Historical Log Table
  var startRow = CONFIG.SERIAL_HISTORY_VIEWER.RANGES.RESULTS_START_ROW;
  var outputRange = historySheet.getRange(startRow, 1, matchingRows.length, 16);
  outputRange.setValues(matchingRows);

  // Apply Status Colors
  var bgColors = [];
  for (var i = 0; i < matchingRows.length; i++) {
    var rowBg = new Array(16).fill("#FFFFFF");
    var t1 = String(matchingRows[i][10]).toUpperCase();
    var t2 = String(matchingRows[i][11]).toUpperCase();
    var stat = String(matchingRows[i][12]).toUpperCase();

    if (t1.includes("FAIL")) rowBg[10] = "#FADBD8";
    if (t2.includes("FAIL")) rowBg[11] = "#FCF3CF";

    if (stat.includes("PASS")) {
      rowBg[12] = "#D4EFDF";
    } else if (stat.includes("FAIL")) {
      rowBg[12] = "#FADBD8";
    } else {
      rowBg[12] = "#FCF3CF";
    }
    bgColors.push(rowBg);
  }
  outputRange.setBackgrounds(bgColors);
}