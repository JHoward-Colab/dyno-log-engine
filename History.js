// =========================================================================
// 📈 SERIAL HISTORICAL ANALYZER (History.js)
// =========================================================================

function renderSerialHistory(e) {
  try {
    var ss = e ? e.source : SpreadsheetApp.getActiveSpreadsheet();
    var tabName = (CONFIG.SHEET_NAMES && CONFIG.SHEET_NAMES.SERIAL_HISTORY_VIEWER) ? CONFIG.SHEET_NAMES.SERIAL_HISTORY_VIEWER : "SERIAL_HISTORY_VIEWER";
    var historySheet = ss.getSheetByName(tabName);
    var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);

    if (!historySheet || !logSheet) return;

    var searchCell = (CONFIG.SERIAL_HISTORY_VIEWER && CONFIG.SERIAL_HISTORY_VIEWER.RANGES) ? CONFIG.SERIAL_HISTORY_VIEWER.RANGES.SERIAL_SEARCH_INPUT : "B2";
    var targetSerial = String(historySheet.getRange(searchCell).getValue()).trim();

    // Clear previous output & summary
    var clearRange = (CONFIG.SERIAL_HISTORY_VIEWER && CONFIG.SERIAL_HISTORY_VIEWER.RANGES) ? CONFIG.SERIAL_HISTORY_VIEWER.RANGES.CLEAR_RESULTS_RANGE : "A9:P1000";
    historySheet.getRange(clearRange).clearContent().setBackground(null).setFontWeight("normal");
    historySheet.getRange("B5:E5").clearContent();

    if (!targetSerial || targetSerial === "undefined" || targetSerial === "") return;

    var logData = logSheet.getDataRange().getValues();
    if (logData.length < 2) return;

    var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG || {};
    var matchingRows = [];
    var detectedBaseModel = "";

    for (var r = 1; r < logData.length; r++) {
      var row = logData[r];
      var logSerial = String(getLogVal(row, logCols.TRUE_SERIAL, 2) || "").trim();

      if (isSerialMatch(targetSerial, logSerial)) {
        if (!detectedBaseModel) {
          var bm = getLogVal(row, logCols.BASE_MODEL, 3);
          if (bm) detectedBaseModel = String(bm).trim();
        }

        var timestamp   = getLogVal(row, logCols.TIMESTAMP, 0);
        var rf          = safeAbsNum(getLogVal(row, logCols.ROD_FORCE, 5));
        var c1          = safeAbsNum(getLogVal(row, logCols.COMP_1, 7));
        var r1          = safeAbsNum(getLogVal(row, logCols.REB_1, 8));
        var sl1         = getLogVal(row, logCols.SLOPE_1, 9);
        var la1         = getLogVal(row, logCols.LOOP_AREA_1, 10);
        var c2          = safeAbsNum(getLogVal(row, logCols.COMP_2, 12));
        var r2          = safeAbsNum(getLogVal(row, logCols.REB_2, 13));
        var la2         = getLogVal(row, logCols.LOOP_AREA_2, 14);
        var t1Stat      = getLogVal(row, logCols.TEST_1_STATUS, 18);
        var t2Stat      = getLogVal(row, logCols.TEST_2_STATUS, 19);
        var overStat    = String(getLogVal(row, logCols.OVERALL_STATUS, 20) || "").toUpperCase();
        var diag        = getLogVal(row, logCols.DIAGNOSTICS, 21);
        var evalAction  = getLogVal(row, logCols.EVALUATION_ACTION, 22);
        var engComments = getLogVal(row, logCols.ENG_COMMENTS, 23);

        matchingRows.push([
          timestamp, logSerial, rf, c1, r1, sl1, la1, c2, r2, la2,
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
    var startRow = CONFIG.SERIAL_HISTORY_VIEWER.RANGES.RESULTS_START_ROW || 9;
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

  } catch(err) {
    Logger.log("renderSerialHistory Error: " + err.toString());
  }
}