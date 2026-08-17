// =========================================================================
// 📊 SUMMARY DASHBOARD CONTROLLER (Summary.js)
// Isolated Work Order Compiler, Progress Tracker & Interactive Navigator
// =========================================================================

/**
 * Rebuilds the Summary Dashboard tab from Work Order Drive files and Master_Dyno_Log.
 */
function buildSummaryDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);

  if (!summarySheet || !logSheet) return;

  var folderId = CONFIG.FOLDERS.WORK_ORDER_FOLDER_ID;
  if (!folderId) return;

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var logData = logSheet.getDataRange().getValues();
  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG || {};
  var sumCols = CONFIG.COLUMNS.SUMMARY || {};

  var tableOutput = [];
  var bgColors = [];
  var fontColors = [];
  var fontWeights = [];

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    if (file.getMimeType() === MimeType.GOOGLE_SHEETS || file.getMimeType() === MimeType.MICROSOFT_EXCEL || fileName.endsWith(".xlsx")) {
      try {
        var woSs = SpreadsheetApp.openById(file.getId());
        var woSheet = woSs.getSheets()[0];

        var woNumber = fileName.replace(/\.[^/.]+$/, "").trim();
        var baseModel = String(woSheet.getRange("D3").getValue()).trim();
        var bomRev = String(woSheet.getRange("D4").getValue()).trim();

        // Extract expected serial numbers from Cell A12 downward
        var woLastRow = woSheet.getLastRow();
        var expectedSerials = [];
        if (woLastRow >= 12) {
          var rawSerials = woSheet.getRange(12, 1, woLastRow - 11, 1).getValues();
          for (var s = 0; s < rawSerials.length; s++) {
            var sVal = String(rawSerials[s][0] || "").trim();
            if (sVal && sVal.toLowerCase() !== "undefined" && sVal.toLowerCase() !== "null") {
              expectedSerials.push(sVal);
            }
          }
        }

        var totalQty = expectedSerials.length;
        var testedCount = 0;
        var passCount = 0;
        var failCount = 0;
        var holdCount = 0;
        var failureDetails = [];
        var lastDate = null;

        for (var es = 0; es < expectedSerials.length; es++) {
          var expS = expectedSerials[es];
          var matchedRun = null;

          for (var r = 1; r < logData.length; r++) {
            var rowSerial = String(logData[r][logCols.TRUE_SERIAL - 1] || "").trim();
            if (isSerialMatch(expS, rowSerial)) {
              matchedRun = logData[r];
              break;
            }
          }

          if (matchedRun) {
            testedCount++;
            var overall = String(matchedRun[logCols.OVERALL_STATUS - 1] || "").toUpperCase();
            var diag = String(matchedRun[logCols.DIAGNOSTICS - 1] || "");
            var runDate = matchedRun[logCols.TIMESTAMP - 1];

            if (runDate instanceof Date && (!lastDate || runDate > lastDate)) {
              lastDate = runDate;
            }

            if (overall.includes("HOLD")) {
              holdCount++;
              failureDetails.push("#" + expS.slice(-3) + " [HOLD]");
            } else if (overall.includes("FAIL")) {
              failCount++;
              var tagMatch = diag.match(/\[(.*?)\]/);
              var tag = tagMatch ? tagMatch[0] : "[FAIL]";
              failureDetails.push("#" + expS.slice(-3) + " " + tag);
            } else {
              passCount++;
            }
          }
        }

        var woStatus = "PENDING";
        var statusBg = "#FCF3CF";
        var statusFont = "#B7950B";

        if (totalQty === 0) {
          woStatus = "NO SERIALS";
          statusBg = "#F2F4F4";
          statusFont = "#5D6D7E";
        } else if (testedCount === 0) {
          woStatus = "PENDING";
          statusBg = "#FCF3CF";
          statusFont = "#B7950B";
        } else if (holdCount > 0) {
          woStatus = "HOLD";
          statusBg = "#FCF3CF";
          statusFont = "#B7950B";
        } else if (failCount > 0) {
          woStatus = "ACTION REQUIRED";
          statusBg = "#FADBD8";
          statusFont = "#C0392B";
        } else if (testedCount < totalQty) {
          woStatus = "INCOMPLETE";
          statusBg = "#FCF3CF";
          statusFont = "#B7950B";
        } else {
          woStatus = "COMPLETED";
          statusBg = "#D4EFDF";
          statusFont = "#196F3D";
        }

        var progressStr = testedCount + " / " + totalQty + " (" + (totalQty > 0 ? Math.round((testedCount / totalQty) * 100) : 0) + "%)";
        var fpyStr = testedCount > 0 ? ((passCount / testedCount) * 100).toFixed(1) + "%" : "N/A";
        var dateStr = lastDate ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "N/A";
        var detailsStr = failureDetails.length > 0 ? failureDetails.join(", ") : (testedCount === totalQty ? "✅ All Units Passed" : "⏳ Pending dyno test");

        // Dynamically build row based on Config Column Indices
        var rowData = new Array(8);
        rowData[sumCols.WORK_ORDER_STATUS - 1] = woStatus;
        rowData[sumCols.WORK_ORDER_NUMBER - 1] = woNumber;
        rowData[sumCols.BASE_MODEL - 1]        = baseModel;
        rowData[sumCols.BOM_REVISION - 1]      = bomRev;
        rowData[sumCols.TESTING_PROGRESS - 1]  = progressStr;
        rowData[sumCols.STATUS_DETAILS - 1]    = detailsStr;
        rowData[sumCols.FIRST_PASS_YIELD - 1]  = fpyStr;
        rowData[sumCols.LAST_TESTED_DATE - 1]  = dateStr;

        var rowBg = new Array(8).fill("#FFFFFF");
        rowBg[sumCols.WORK_ORDER_STATUS - 1] = statusBg;

        var rowFont = new Array(8).fill("#000000");
        rowFont[sumCols.WORK_ORDER_STATUS - 1] = statusFont;
        rowFont[sumCols.WORK_ORDER_NUMBER - 1] = "#0000FF";

        var rowWeight = new Array(8).fill("normal");
        rowWeight[sumCols.WORK_ORDER_STATUS - 1] = "bold";
        rowWeight[sumCols.WORK_ORDER_NUMBER - 1] = "bold";

        tableOutput.push(rowData);
        bgColors.push(rowBg);
        fontColors.push(rowFont);
        fontWeights.push(rowWeight);

      } catch (e) {
        Logger.log("Error processing WO file " + fileName + ": " + e.toString());
      }
    }
  }

  if (tableOutput.length > 0) {
    var maxRows = Math.max(summarySheet.getLastRow() - 1, 1);
    summarySheet.getRange(2, 1, maxRows, 8).clearContent().setBackground(null).setFontColor(null).setFontWeight("normal");

    var targetRange = summarySheet.getRange(2, 1, tableOutput.length, 8);
    targetRange.setValues(tableOutput);
    targetRange.setBackgrounds(bgColors);
    targetRange.setFontColors(fontColors);
    targetRange.setFontWeights(fontWeights);
  }
}

/**
 * Interactive Work Order Jumper.
 * Clicking any cell in Work_Order_Number column on Summary tab loads that Work Order into Operator_Station.
 */
function onSelectionChange(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sumCols = CONFIG.COLUMNS.SUMMARY || {};

  if (sheet.getName() === CONFIG.SHEET_NAMES.SUMMARY && e.range.getColumn() === (sumCols.WORK_ORDER_NUMBER || 2) && e.range.getRow() > 1) {
    var woNumber = String(e.range.getValue()).trim();
    if (!woNumber) return;

    var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
    var opSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);

    if (opSheet) {
      opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT).setValue(woNumber);
      manageOperatorStation({ source: ss, range: opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) });
      ss.setActiveSheet(opSheet);
    }
  }
}