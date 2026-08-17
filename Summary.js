// =========================================================================
// 📊 SUMMARY DASHBOARD CONTROLLER (Summary.js)
// Isolated Work Order Compiler, Progress Tracker & Interactive Navigator
// =========================================================================

/**
 * Robustly extracts WO batch number from barcodes, short serials, or file names.
 */
function robustExtractWoBatchNum(strVal) {
  var s = String(strVal || "").trim();
  if (!s) return 0;

  var cleanDigits = s.replace(/[^0-9]/g, "");
  if (!cleanDigits) return 0;

  if (cleanDigits.length >= 14) {
    var batchPart = cleanDigits.slice(-7, -3);
    var pBatch = parseInt(batchPart, 10);
    if (!isNaN(pBatch) && pBatch > 0) return pBatch;
  }

  if (s.indexOf("-") !== -1) {
    var prefix = s.split("-")[0];
    var pDigits = prefix.replace(/[^0-9]/g, "");
    if (pDigits) return parseInt(pDigits, 10) || 0;
  }

  var match = s.match(/\b\d{4,6}\b/) || s.match(/\d{4,6}/);
  if (match) {
    return parseInt(match[0], 10) || 0;
  }

  return parseInt(cleanDigits, 10) || 0;
}

/**
 * Rebuilds the Summary Dashboard tab incrementally starting at WO 1608 baseline.
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

  // HARD BASELINE FLOOR: Ignore any Work Orders prior to WO 1608
  var BASELINE_WO_FLOOR = 1608;
  var minWoNumber = BASELINE_WO_FLOOR;

  var logMapByCleanSerial = {};
  var logSerialsList = [];

  for (var r = 1; r < logData.length; r++) {
    var rawSerial = String(logData[r][(logCols.TRUE_SERIAL || 3) - 1] || "").trim();
    if (rawSerial) {
      var cSer = cleanKey(rawSerial);
      logMapByCleanSerial[cSer] = logData[r];
      logSerialsList.push({ clean: cSer, raw: rawSerial, row: logData[r] });
    }
  }

  // STEP 1: Fast Direct Folder Scan Filtering Out Files < WO 1608
  var filesIterator = folder.getFiles();
  var fileList = [];

  while (filesIterator.hasNext()) {
    var f = filesIterator.next();
    var fName = f.getName();

    if (fName.indexOf(".xlsx") !== -1 && fName.indexOf("~") === 0) continue;

    var fWoNum = robustExtractWoBatchNum(fName);

    // Skip any files before baseline floor 1608
    if (fWoNum > 0 && fWoNum < minWoNumber) {
      continue;
    }

    fileList.push({
      file: f,
      id: f.getId(),
      name: fName,
      woNum: fWoNum
    });
  }

  // STEP 2: Sort Work Orders Ascending starting from WO 1608
  fileList.sort(function(a, b) {
    if (a.woNum !== b.woNum && a.woNum > 0 && b.woNum > 0) {
      return a.woNum - b.woNum;
    }
    return a.name.localeCompare(b.name);
  });

  var props = PropertiesService.getScriptProperties();
  var uncachedOpenedCount = 0;
  var MAX_UNCACHED_OPENS = 5; // Capped at 5 new uncached opens per run

  var tableOutput = [];
  var bgColors = [];
  var fontColors = [];
  var fontWeights = [];

  // STEP 3: Process Active Work Orders >= 1608
  for (var i = 0; i < fileList.length; i++) {
    var item = fileList[i];
    var fileId = item.id;
    var fileName = item.name;
    var woNumber = fileName.replace(/\.[^/.]+$/, "").trim();

    var fileUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/edit";
    var woLinkFormula = '=HYPERLINK("' + fileUrl + '", "' + woNumber + '")';

    try {
      var baseModel = "";
      var bomRev = "";
      var expectedSerials = [];

      var propKey = "WO_META_" + fileId;
      var cachedStr = props.getProperty(propKey);

      if (cachedStr) {
        var cachedData = JSON.parse(cachedStr);
        baseModel = cachedData.baseModel;
        bomRev = cachedData.bomRev;
        expectedSerials = cachedData.expectedSerials || [];
      } else if (uncachedOpenedCount < MAX_UNCACHED_OPENS) {
        uncachedOpenedCount++;
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

        var cachePayload = {
          baseModel: baseModel,
          bomRev: bomRev,
          expectedSerials: expectedSerials
        };
        props.setProperty(propKey, JSON.stringify(cachePayload));
      } else {
        baseModel = "PENDING CACHE";
        bomRev = "-";
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

      if (baseModel === "PENDING CACHE") {
        woStatus = "PENDING";
        statusBg = "#FCF3CF";
        statusFont = "#B7950B";
      } else if (totalQty === 0) {
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

      var progressStr = (baseModel === "PENDING CACHE") ? "Indexing..." : testedCount + " / " + totalQty + " (" + (totalQty > 0 ? Math.round((testedCount / totalQty) * 100) : 0) + "%)";
      var fpyStr = testedCount > 0 ? ((passCount / testedCount) * 100).toFixed(1) + "%" : "N/A";
      var dateStr = lastDate ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "N/A";
      var detailsStr = failureDetails.length > 0 ? failureDetails.join(", ") : (testedCount === totalQty && totalQty > 0 ? "✅ All Units Passed" : "⏳ Pending dyno test");

      var rowData = new Array(8);
      rowData[(sumCols.WORK_ORDER_STATUS || 1) - 1] = woStatus;
      rowData[(sumCols.WORK_ORDER_NUMBER || 2) - 1] = woLinkFormula;
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

  // STEP 4: Render Table Starting Clean at WO 1608
  var maxRows = Math.max(summarySheet.getLastRow() - 1, 1);
  summarySheet.getRange(2, 1, maxRows, 8).clearContent().setBackground(null).setFontColor(null).setFontWeight("normal");

  if (tableOutput.length > 0) {
    var targetRange = summarySheet.getRange(2, 1, tableOutput.length, 8);
    targetRange.setValues(tableOutput);
    targetRange.setBackgrounds(bgColors);
    targetRange.setFontColors(fontColors);
    targetRange.setFontWeights(fontWeights);
  } else {
    summarySheet.getRange("A2").setValue("⚠️ No Work Orders found matching baseline threshold (WO >= " + minWoNumber + ").");
  }
}

/**
 * Utility helper to clear persistent WO metadata properties if files are modified.
 */
function clearWoSummaryCache() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf("WO_META_") === 0) {
      props.deleteProperty(keys[i]);
    }
  }
  Logger.log("Summary persistent metadata cache cleared.");
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