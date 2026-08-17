// =========================================================================
// 📊 SUMMARY DASHBOARD CONTROLLER (Summary.js)
// Registry-Filtered, Multi-Tier Serial Engine & Interactive Navigator
// =========================================================================

/**
 * Standardized alphanumeric string cleaner.
 */
function cleanKey(str) {
  return String(str || "").toUpperCase().replace(/[^A-Z0-9]/g, "").trim();
}

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
  return match ? (parseInt(match[0], 10) || 0) : (parseInt(cleanDigits, 10) || 0);
}

/**
 * Validates whether cached JSON metadata contains real, indexed data.
 */
function isValidCache(jsonStr) {
  if (!jsonStr) return false;
  try {
    var data = JSON.parse(jsonStr);
    return data && data.bm && data.bm !== "PENDING CACHE" && data.bm !== "ERROR" && data.bm !== "";
  } catch (e) {
    return false;
  }
}

/**
 * Retrieves valid part numbers from Program_Registry to filter Work Orders.
 */
function getValidRegistryModels(ss) {
  var validBaseModels = {};
  var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY || "Program_Registry");
  if (!registrySheet) return validBaseModels;

  var regData = registrySheet.getDataRange().getValues();
  if (regData.length < 2) return validBaseModels;

  var bmColIdx = 0;
  var headers = regData[0];
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).toUpperCase();
    if (h.indexOf("BASE") !== -1 || h.indexOf("PART") !== -1 || h.indexOf("MODEL") !== -1) {
      bmColIdx = c;
      break;
    }
  }

  for (var r = 1; r < regData.length; r++) {
    var bmVal = String(regData[r][bmColIdx] || "").trim().toUpperCase();
    if (bmVal) validBaseModels[bmVal] = true;
  }
  return validBaseModels;
}

/**
 * One-Click Reset & Full Re-index
 * Flushes invalid/placeholder cache entries and forces clean indexing.
 */
function resetAndReindexAll() {
  Logger.log("🧹 Clearing placeholder cache entries...");
  clearWoSummaryCache();
  Logger.log("🚀 Starting clean indexing pass...");
  runFullSystemIndexer();
}

/**
 * Smart-Targeted Indexer with 4-Minute Safety Valve & Registry Filtering.
 */
function runFullSystemIndexer() {
  var startTime = new Date().getTime();
  var MAX_EXECUTION_TIME = 240000; // 4 minutes
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var validModels = getValidRegistryModels(ss);

  var folderId = CONFIG.FOLDERS.WORK_ORDER_FOLDER_ID;
  if (!folderId) {
    Logger.log("❌ Error: WORK_ORDER_FOLDER_ID not set.");
    return;
  }

  var folder = DriveApp.getFolderById(folderId);
  var BASELINE_WO_FLOOR = 1608;

  var propsService = PropertiesService.getScriptProperties();
  var allProps = propsService.getProperties();

  var filesIterator = folder.getFiles();
  var uncachedList = [];

  // 1. Identify Uncached Files
  while (filesIterator.hasNext()) {
    var f = filesIterator.next();
    var fName = f.getName();
    if (fName.indexOf(".xlsx") !== -1 && fName.indexOf("~") === 0) continue;

    var fWoNum = robustExtractWoBatchNum(fName);
    if (fWoNum > 0 && fWoNum < BASELINE_WO_FLOOR) continue;

    var propKey = "WO_META_" + f.getId();
    if (!isValidCache(allProps[propKey])) {
      uncachedList.push({ id: f.getId(), name: fName });
    }
  }

  Logger.log("Found " + uncachedList.length + " Work Orders needing valid indexing.");

  if (uncachedList.length === 0) {
    Logger.log("✅ All Work Orders are 100% validly indexed! Refreshing dashboard...");
    buildSummaryDashboard();
    return;
  }

  var indexedThisRun = 0;

  // 2. Process Files Until 4-Minute Safety Window
  for (var i = 0; i < uncachedList.length; i++) {
    var elapsed = new Date().getTime() - startTime;
    if (elapsed > MAX_EXECUTION_TIME) {
      Logger.log("⏱️ 4-Minute Limit Reached. Validly indexed " + indexedThisRun + " files this run.");
      Logger.log("⚠️ Click 'Run' again to continue indexing the remaining " + (uncachedList.length - indexedThisRun) + " files.");
      break; 
    }

    var item = uncachedList[i];
    try {
      var woSs = SpreadsheetApp.openById(item.id);
      var woSheet = woSs.getSheets()[0];

      var baseModel = String(woSheet.getRange("D3").getValue()).trim();
      var bomRev = String(woSheet.getRange("D4").getValue()).trim();
      var expectedSerials = [];

      var isTrackedPart = (Object.keys(validModels).length === 0) || validModels[baseModel.toUpperCase()];

      // Only extract serials if it matches the Program Registry
      if (isTrackedPart) {
        var woLastRow = woSheet.getLastRow();
        if (woLastRow >= 12) {
          var raw = woSheet.getRange(12, 1, woLastRow - 11, 1).getValues();
          for (var s = 0; s < raw.length; s++) {
            var v = String(raw[s][0] || "").trim();
            if (v && v.toLowerCase() !== "undefined" && v.toLowerCase() !== "null") {
              expectedSerials.push(v);
            }
          }
        }
      } else {
        baseModel = "IGNORED_PART"; // Tags it to be skipped permanently
        bomRev = "-";
      }

      if (baseModel && baseModel !== "undefined") {
        var payloadStr = JSON.stringify({ bm: baseModel, br: bomRev, es: expectedSerials });
        propsService.setProperty("WO_META_" + item.id, payloadStr);
        indexedThisRun++;
        Logger.log("Indexed (" + indexedThisRun + "/" + uncachedList.length + "): " + item.name + " -> " + baseModel);
      }

    } catch (err) {
      Logger.log("Failed to index " + item.name + ": " + err.toString());
    }
  }

  Logger.log("Updating Summary tab dashboard...");
  buildSummaryDashboard();
  SpreadsheetApp.flush();
}

/**
 * Standard Dashboard Builder with Dynamic Registry Filtering.
 */
function buildSummaryDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);

  if (!summarySheet || !logSheet) return;

  var validModels = getValidRegistryModels(ss);

  var folderId = CONFIG.FOLDERS.WORK_ORDER_FOLDER_ID;
  if (!folderId) return;

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

  var BASELINE_WO_FLOOR = 1608;

  // STEP 1: Auto-Detect Column Indices for Master Dyno Log
  var logHeaders = logData[0] || [];
  var colTrueSerial = (logCols.TRUE_SERIAL || 3) - 1;
  var colOverallStatus = (logCols.OVERALL_STATUS || 21) - 1;
  var colDiagnostics = (logCols.DIAGNOSTICS || 22) - 1;
  var colTimestamp = (logCols.TIMESTAMP || 1) - 1;

  for (var c = 0; c < logHeaders.length; c++) {
    var hText = String(logHeaders[c]).toUpperCase().trim();
    if (hText.indexOf("TRUE_SERIAL") !== -1 || hText === "SERIAL") colTrueSerial = c;
    if (hText.indexOf("OVERALL") !== -1) colOverallStatus = c;
    if (hText.indexOf("DIAG") !== -1) colDiagnostics = c;
    if (hText.indexOf("TIME") !== -1 || hText.indexOf("DATE") !== -1) colTimestamp = c;
  }

  // STEP 2: Index Dyno Runs by Full Key AND Trailing Short Suffixes
  var allRunsBySerial = {};
  for (var r = 1; r < logData.length; r++) {
    var rawSerial = String(logData[r][colTrueSerial] || "").trim();
    if (rawSerial) {
      var cSer = cleanKey(rawSerial);
      if (!allRunsBySerial[cSer]) allRunsBySerial[cSer] = [];
      allRunsBySerial[cSer].push(logData[r]);

      if (cSer.length >= 3) {
        var s3 = cSer.slice(-3);
        var shortKey = "SHORT_" + s3;
        if (!allRunsBySerial[shortKey]) allRunsBySerial[shortKey] = [];
        allRunsBySerial[shortKey].push(logData[r]);
      }
    }
  }

  // STEP 3: Load Persistent Cache
  var propsService = PropertiesService.getScriptProperties();
  var allProps = propsService.getProperties();

  // STEP 4: Scan Drive Folder Files
  var filesIterator = folder.getFiles();
  var fileList = [];

  while (filesIterator.hasNext()) {
    var f = filesIterator.next();
    var fName = f.getName();
    if (fName.indexOf(".xlsx") !== -1 && fName.indexOf("~") === 0) continue;

    var fWoNum = robustExtractWoBatchNum(fName);
    if (fWoNum > 0 && fWoNum < BASELINE_WO_FLOOR) continue;

    fileList.push({ id: f.getId(), name: fName, woNum: fWoNum });
  }

  fileList.sort(function(a, b) {
    return (a.woNum !== b.woNum && a.woNum > 0 && b.woNum > 0) ? (a.woNum - b.woNum) : a.name.localeCompare(b.name);
  });

  var tableOutput = [];
  var bgColors = [];
  var fontColors = [];
  var fontWeights = [];

  // STEP 5: Process Summary Table Matrix
  for (var i = 0; i < fileList.length; i++) {
    var item = fileList[i];
    var fileId = item.id;
    var fileName = item.name;
    var woNumber = fileName.replace(/\.[^/.]+$/, "").trim();

    var fileUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/edit";
    var woLinkFormula = '=HYPERLINK("' + fileUrl + '", "' + woNumber + '")';

    var baseModel = "";
    var bomRev = "";
    var expectedSerials = [];

    var propKey = "WO_META_" + fileId;
    var cachedStr = allProps[propKey];

    if (isValidCache(cachedStr)) {
      var cachedData = JSON.parse(cachedStr);
      baseModel = cachedData.bm || "";
      bomRev = cachedData.br || "";
      expectedSerials = cachedData.es || [];
    } else {
      baseModel = "PENDING CACHE";
      bomRev = "-";
    }

    // 🔥 REGISTRY FILTER: Skip kits, service parts, and removed parts dynamically
    if (baseModel === "IGNORED_PART") continue;
    if (baseModel !== "PENDING CACHE" && Object.keys(validModels).length > 0 && !validModels[baseModel.toUpperCase()]) {
      continue;
    }

    var totalQty = expectedSerials.length;
    var testedCount = 0;
    var firstPassCount = 0;
    var activeFailCount = 0;
    var activeHoldCount = 0;
    var activeFailureDetails = [];
    var lastDate = null;

    for (var es = 0; es < expectedSerials.length; es++) {
      var expS = expectedSerials[es];
      var cExp = cleanKey(expS);

      // Multi-Tier Matching Engine
      var runs = allRunsBySerial[cExp];

      if (!runs || runs.length === 0) {
        var woNumClean = cleanKey(item.woNum || woNumber);
        runs = allRunsBySerial[woNumClean + cExp] || allRunsBySerial["WO" + woNumClean + cExp];
      }

      if (!runs || runs.length === 0) {
        var expS3 = cExp.slice(-3);
        if (expS3.length === 3) {
          var potentialRuns = allRunsBySerial["SHORT_" + expS3] || [];
          if (potentialRuns.length > 0) {
            var woSearch = String(item.woNum || woNumber).toUpperCase();
            runs = potentialRuns.filter(function(runRow) {
              var rowStr = runRow.join(" ").toUpperCase();
              return rowStr.indexOf(woSearch) !== -1;
            });
            if (runs.length === 0) runs = potentialRuns; // Fallback to short suffix match
          }
        }
      }

      if (runs && runs.length > 0) {
        testedCount++;

        var firstRun = runs[0];
        if (String(firstRun[colOverallStatus] || "").toUpperCase().includes("PASS")) {
          firstPassCount++;
        }

        var latestRun = runs[runs.length - 1];
        var latestOverall = String(latestRun[colOverallStatus] || "").toUpperCase();
        var latestDiag = String(latestRun[colDiagnostics] || "");
        var runDate = latestRun[colTimestamp];

        if (runDate instanceof Date && (!lastDate || runDate > lastDate)) {
          lastDate = runDate;
        }

        if (latestOverall.includes("HOLD")) {
          activeHoldCount++;
          activeFailureDetails.push("#" + expS.slice(-3) + " [HOLD]");
        } else if (latestOverall.includes("FAIL")) {
          activeFailCount++;
          var tagMatch = latestDiag.match(/\[(.*?)\]/);
          activeFailureDetails.push("#" + expS.slice(-3) + " " + (tagMatch ? tagMatch[0] : "[FAIL]"));
        }
      }
    }

    var woStatus = "PENDING";
    var statusBg = "#FCF3CF";
    var statusFont = "#B7950B";

    if (baseModel === "PENDING CACHE") {
      woStatus = "PENDING"; statusBg = "#FCF3CF"; statusFont = "#B7950B";
    } else if (totalQty === 0) {
      woStatus = "NO SERIALS"; statusBg = "#F2F4F4"; statusFont = "#5D6D7E";
    } else if (testedCount === 0) {
      woStatus = "PENDING"; statusBg = "#FCF3CF"; statusFont = "#B7950B";
    } else if (activeHoldCount > 0) {
      woStatus = "HOLD"; statusBg = "#FCF3CF"; statusFont = "#B7950B";
    } else if (activeFailCount > 0) {
      woStatus = "ACTION REQUIRED"; statusBg = "#FADBD8"; statusFont = "#C0392B";
    } else if (testedCount < totalQty) {
      woStatus = "INCOMPLETE"; statusBg = "#FCF3CF"; statusFont = "#B7950B";
    } else {
      woStatus = "COMPLETED"; statusBg = "#D4EFDF"; statusFont = "#196F3D";
    }

    var progressStr = (baseModel === "PENDING CACHE") ? "Indexing..." : testedCount + " / " + totalQty + " (" + (totalQty > 0 ? Math.round((testedCount / totalQty) * 100) : 0) + "%)";
    var fpyStr = testedCount > 0 ? ((firstPassCount / testedCount) * 100).toFixed(1) + "%" : "N/A";
    var dateStr = lastDate ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "N/A";
    var detailsStr = activeFailureDetails.length > 0 ? activeFailureDetails.join(", ") : (testedCount === totalQty && totalQty > 0 ? "✅ All Units Passed" : "⏳ Pending dyno test");

    var rowData = new Array(8);
    rowData[(sumCols.WORK_ORDER_STATUS || 1) - 1] = woStatus;
    rowData[(sumCols.WORK_ORDER_NUMBER || 2) - 1] = woLinkFormula;
    rowData[(sumCols.BASE_MODEL || 3) - 1]        = baseModel;
    rowData[(sumCols.BOM_REVISION || 4) - 1]      = bomRev;
    rowData[(sumCols.TESTING_PROGRESS || 5) - 1]  = progressStr;
    rowData[(sumCols.STATUS_DETAILS || 6) - 1]    = detailsStr;
    rowData[(sumCols.FIRST_PASS_YIELD || 7) - 1]  = fpyStr;
    rowData[(sumCols.LAST_TESTED_DATE || 8) - 1]  = dateStr;

    tableOutput.push(rowData);
    bgColors.push([statusBg, "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF"]);
    fontColors.push([statusFont, "#0000FF", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000"]);
    fontWeights.push(["bold", "bold", "normal", "normal", "normal", "normal", "normal", "normal"]);
  }

  // STEP 6: Single Bulk Render to Summary Sheet
  var maxRows = Math.max(summarySheet.getLastRow() - 1, 1);
  summarySheet.getRange(2, 1, maxRows, 8).clearContent().setBackground(null).setFontColor(null).setFontWeight("normal");

  if (tableOutput.length > 0) {
    var targetRange = summarySheet.getRange(2, 1, tableOutput.length, 8);
    targetRange.setValues(tableOutput);
    targetRange.setBackgrounds(bgColors);
    targetRange.setFontColors(fontColors);
    targetRange.setFontWeights(fontWeights);
  }
}

/**
 * Flushes all cached metadata.
 */
function clearWoSummaryCache() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf("WO_META_") === 0) props.deleteProperty(keys[i]);
  }
  Logger.log("Summary persistent metadata cache cleared.");
}

/**
 * Interactive Work Order Jumper.
 * Formats zero-padded file names (e.g., WO-002038 -> WO-2038) before pushing to Operator_Station.
 */
function onSelectionChange(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (sheet.getName() === CONFIG.SHEET_NAMES.SUMMARY && (col === 1 || col === 2) && row > 1) {
    var rawVal = String(sheet.getRange(row, 2).getValue()).trim();
    if (!rawVal || rawVal.startsWith("⚠️") || rawVal.startsWith("❌")) return;

    // Strip leading zeroes (e.g. "WO-002038" -> 2038 -> "WO-2038")
    var woBatch = robustExtractWoBatchNum(rawVal);
    var formattedVal = woBatch > 0 ? "WO-" + woBatch : rawVal;

    var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
    var opSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);

    if (opSheet) {
      opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT).setValue(formattedVal);
      manageOperatorStation({ source: ss, range: opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) });
      ss.setActiveSheet(opSheet);
    }
  }
}