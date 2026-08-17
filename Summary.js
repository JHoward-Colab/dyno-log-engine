// =========================================================================
// 📊 SUMMARY DASHBOARD CONTROLLER (Summary.js)
// Isolated Work Order Compiler, Progress Tracker & Interactive Navigator
// =========================================================================

/**
 * Safely extracts 4-digit/6-digit WO batch number from barcodes, short serials, or file names.
 * e.g., "43081008001979001" -> 1979
 * e.g., "001979-001"        -> 1979
 * e.g., "001979"            -> 1979
 */
function extractWoBatchNum(strVal) {
  var s = String(strVal || "").trim();
  if (!s) return 0;

  if (s.indexOf("-") !== -1) {
    s = s.split("-")[0];
  }

  var clean = s.replace(/[^0-9]/g, "");
  if (!clean) return 0;

  if (clean.length >= 10) {
    var batchPart = clean.slice(-7, -3);
    return parseInt(batchPart, 10) || 0;
  }

  return parseInt(clean, 10) || 0;
}

/**
 * Rebuilds the Summary Dashboard tab from Work Order Drive files and Master_Dyno_Log.
 */
function buildSummaryDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);

  if (!summarySheet || !logSheet) return;

  var folderId = CONFIG.FOLDERS.WORK_ORDER_FOLDER_ID;
  if (!folderId) {
    summarySheet.getRange("A2").setValue("❌ Error: WORK_ORDER_FOLDER_ID not set in Config.js");
    return;
  }

  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (err) {
    summarySheet.getRange("A2").setValue("❌ Error: Invalid Folder ID or Access Denied.");
    return;
  }

  var logData = logSheet.getDataRange().getValues();
  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG || {};
  var sumCols = CONFIG.COLUMNS.SUMMARY || {};

  // STEP 1: Determine Lowest Active WO Number in Master_Dyno_Log and Index Serials
  var minWoNumber = 999999;
  var logMapByCleanSerial = {};
  var logSerialsList = [];

  for (var r = 1; r < logData.length; r++) {
    var rawSerial = String(logData[r][(logCols.TRUE_SERIAL || 3) - 1] || "").trim();
    if (rawSerial) {
      var cSer = cleanKey(rawSerial);
      logMapByCleanSerial[cSer] = logData[r];
      logSerialsList.push({ clean: cSer, raw: rawSerial, row: logData[r] });

      var woNum = extractWoBatchNum(rawSerial);
      if (woNum > 0 && woNum < minWoNumber) {
        minWoNumber = woNum;
      }
    }
  }

  if (minWoNumber === 999999) minWoNumber = 0;

  // STEP 2: Query Drive Files
  var files = folder.searchFiles("mimeType = '" + MimeType.GOOGLE_SHEETS + "' and trashed = false");
  var cache = CacheService.getScriptCache();

  var tableOutput = [];
  var bgColors = [];
  var fontColors = [];
  var fontWeights = [];

  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    var fileName = file.getName();
    var woNumber = fileName.replace(/\.[^/.]+$/, "").trim();

    // STEP 3: Instant Cutoff Check (Skip legacy files without opening)
    var fileWoNum = extractWoBatchNum(fileName);
    if (fileWoNum > 0 && minWoNumber > 0 && fileWoNum < minWoNumber) {
      continue; // Skip legacy file in 0ms
    }

    try {
      var baseModel = "";
      var bomRev = "";
      var expectedSerials = [];

      // STEP 4: High-Speed Cache Lookup for File Metadata
      var cacheKey = "WO_META_" + fileId;
      var cachedJson = cache.get(cacheKey);

      if (cachedJson) {
        var cachedData = JSON.parse(cachedJson);
        baseModel = cachedData.baseModel;
        bomRev = cachedData.bomRev;
        expectedSerials = cachedData.expectedSerials;
      } else {
        // Open file once on cache miss and save to Cache
        var woSs = SpreadsheetApp.openById(fileId);
        var woSheet = woSs.getSheets()[0];

        baseModel = String(woSheet.getRange("D3").getValue()).trim();
        bomRev = String(woSheet.getRange("D4").getValue()).trim();

        var woLastRow = woSheet.getLastRow();
        if (woLastRow >= 12) {
          var rawSerials = woSheet.getRange(12, 1, woLastRow - 11, 1).getValues();
          for (var s = 0; s < rawSerials.length; s++) {
            var sVal = String(rawSerials[s][0] || "").trim();
            if (sVal && sVal.toLowerCase() !== "undefined" && sVal.toLowerCase() !== "null") {
              expectedSerials.push(sVal);
            }
          }
        }

        // Cache metadata for 6 hours (21,600 seconds)
        var cachePayload = {
          baseModel: baseModel,
          bomRev: bomRev,
          expectedSerials: expectedSerials
        };
        cache.put(cacheKey, JSON.stringify(cachePayload), 21600);
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
        var cleanExp = cleanKey(expS);

        if (logMapByCleanSerial[cleanExp]) {
          matchedRun = logMapByCleanSerial[cleanExp];
        } else {
          for (var lIdx = 0; lIdx < logSerialsList.length; lIdx++) {
            if (isSerialMatch(expS, logSerialsList[lIdx].raw)) {
              matchedRun = logSerialsList[lIdx].row;
              break;
            }
          }
        }

        if (matchedRun) {
          testedCount++;
          var overall = String(matchedRun[(logCols.OVERALL_STATUS || 21) - 1] || "").toUpperCase();
          var diag = String(matchedRun[(logCols.DIAGNOSTICS || 22) - 1] || "");
          var runDate = matchedRun[(logCols.TIMESTAMP || 1) - 1];

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

      var rowData = new Array(8);
      rowData[(sumCols.WORK_ORDER_STATUS || 1) - 1] = woStatus;
      rowData[(sumCols.WORK_ORDER_NUMBER || 2) - 1] = woNumber;
      rowData[(sumCols.BASE_MODEL || 3) - 1]        = baseModel;
      rowData[(sumCols.BOM_REVISION || 4) - 1]      = bomRev;
      rowData[(sumCols.TESTING_PROGRESS || 5) - 1]  = progressStr;
      rowData[(sumCols.STATUS_DETAILS || 6) - 1]    = detailsStr;
      rowData[(sumCols.FIRST_PASS_YIELD || 7) - 1]  = fpyStr;
      rowData[(sumCols.LAST_TESTED_DATE || 8) - 1]  = dateStr;

      var rowBg = new Array(8).fill("#FFFFFF");
      rowBg[(sumCols.WORK_ORDER_STATUS || 1) - 1] = statusBg;

      var rowFont = new Array(8).fill("#000000");
      rowFont[(sumCols.WORK_ORDER_STATUS || 1) - 1] = statusFont;
      rowFont[(sumCols.WORK_ORDER_NUMBER || 2) - 1] = "#0000FF";

      var rowWeight = new Array(8).fill("normal");
      rowWeight[(sumCols.WORK_ORDER_STATUS || 1) - 1] = "bold";
      rowWeight[(sumCols.WORK_ORDER_NUMBER || 2) - 1] = "bold";

      tableOutput.push(rowData);
      bgColors.push(rowBg);
      fontColors.push(rowFont);
      fontWeights.push(rowWeight);

    } catch (e) {
      Logger.log("Error processing WO file " + fileName + ": " + e.toString());
    }
  }

  var maxRows = Math.max(summarySheet.getLastRow() - 1, 1);
  summarySheet.getRange(2, 1, maxRows, 8).clearContent().setBackground(null).setFontColor(null).setFontWeight("normal");

  if (tableOutput.length > 0) {
    var targetRange = summarySheet.getRange(2, 1, tableOutput.length, 8);
    targetRange.setValues(tableOutput);
    targetRange.setBackgrounds(bgColors);
    targetRange.setFontColors(fontColors);
    targetRange.setFontWeights(fontWeights);
  } else {
    summarySheet.getRange("A2").setValue("⚠️ No Work Orders matched starting threshold (WO >= " + minWoNumber + ").");
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
    if (!woNumber || woNumber.startsWith("⚠️") || woNumber.startsWith("❌")) return;

    var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
    var opSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);

    if (opSheet) {
      opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT).setValue(woNumber);
      manageOperatorStation({ source: ss, range: opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) });
      ss.setActiveSheet(opSheet);
    }
  }
}