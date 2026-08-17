// =========================================================================
// 📊 SUMMARY DASHBOARD CONTROLLER (Summary.js)
// Registry-Filtered, Conditional-Pass Aware, Strict-Match Serial Engine
// =========================================================================

function cleanKey(str) {
  return String(str || "").toUpperCase().replace(/[^A-Z0-9]/g, "").trim();
}

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

function isValidCache(jsonStr) {
  if (!jsonStr) return false;
  try {
    var data = JSON.parse(jsonStr);
    return data && data.bm && data.bm !== "PENDING CACHE" && data.bm !== "ERROR" && data.bm !== "";
  } catch (e) {
    return false;
  }
}

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

function resetAndReindexAll() {
  Logger.log("🧹 Clearing placeholder cache entries...");
  clearWoSummaryCache();
  Logger.log("🚀 Starting clean indexing pass...");
  runFullSystemIndexer();
}

function runFullSystemIndexer() {
  var startTime = new Date().getTime();
  var MAX_EXECUTION_TIME = 240000;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var validModels = getValidRegistryModels(ss);

  var folderId = CONFIG.FOLDERS.WORK_ORDER_FOLDER_ID;
  if (!folderId) return;

  var folder = DriveApp.getFolderById(folderId);
  var BASELINE_WO_FLOOR = 1608;

  var propsService = PropertiesService.getScriptProperties();
  var allProps = propsService.getProperties();

  var filesIterator = folder.getFiles();
  var uncachedList = [];

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

  if (uncachedList.length === 0) {
    buildSummaryDashboard();
    return;
  }

  var indexedThisRun = 0;

  for (var i = 0; i < uncachedList.length; i++) {
    var elapsed = new Date().getTime() - startTime;
    if (elapsed > MAX_EXECUTION_TIME) break; 

    var item = uncachedList[i];
    try {
      var woSs = SpreadsheetApp.openById(item.id);
      var woSheet = woSs.getSheets()[0];

      var baseModel = String(woSheet.getRange("D3").getValue()).trim();
      var bomRev = String(woSheet.getRange("D4").getValue()).trim();
      var expectedSerials = [];

      var isTrackedPart = (Object.keys(validModels).length === 0) || validModels[baseModel.toUpperCase()];

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
        baseModel = "IGNORED_PART";
        bomRev = "-";
      }

      if (baseModel && baseModel !== "undefined") {
        var payloadStr = JSON.stringify({ bm: baseModel, br: bomRev, es: expectedSerials });
        propsService.setProperty("WO_META_" + item.id, payloadStr);
        indexedThisRun++;
      }

    } catch (err) {}
  }

  buildSummaryDashboard();
  SpreadsheetApp.flush();
}

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

  var propsService = PropertiesService.getScriptProperties();
  var allProps = propsService.getProperties();

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

    if (baseModel === "IGNORED_PART") continue;
    if (baseModel !== "PENDING CACHE" && Object.keys(validModels).length > 0 && !validModels[baseModel.toUpperCase()]) {
      continue;
    }

    var totalQty = expectedSerials.length;
    var testedCount = 0;
    var firstPassCount = 0;
    var activeFailCount = 0;
    var activeHoldCount = 0;
    var activeCondCount = 0;
    var activeFailureDetails = [];
    var lastDate = null;
    var woNumClean = cleanKey(item.woNum || woNumber);

    for (var es = 0; es < expectedSerials.length; es++) {
      var expS = expectedSerials[es];
      var cExp = cleanKey(expS);

      var runs = allRunsBySerial[cExp];

      if (!runs || runs.length === 0) {
        runs = allRunsBySerial[woNumClean + cExp] || allRunsBySerial["WO" + woNumClean + cExp];
      }

      if (!runs || runs.length === 0) {
        var expS3 = cExp.slice(-3);
        if (expS3.length === 3) {
          var potentialRuns = allRunsBySerial["SHORT_" + expS3] || [];
          if (potentialRuns.length > 0) {
            runs = potentialRuns.filter(function(runRow) {
              var rowStr = runRow.join(" ").toUpperCase();
              return rowStr.indexOf(woNumClean) !== -1;
            });
          }
        }
      }

      if (runs && runs.length > 0) {
        testedCount++;

        // First Pass Yield evaluation (Strict clean pass without conditional flags)
        var firstRun = runs[0];
        var firstOverall = String(firstRun[colOverallStatus] || "").toUpperCase();
        if (firstOverall.includes("PASS") && !firstOverall.includes("COND")) {
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
        } else if (latestOverall.includes("COND")) {
          activeCondCount++;
          activeFailureDetails.push("#" + expS.slice(-3) + " [COND PASS]");
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
    } else if (activeCondCount > 0) {
      woStatus = "CONDITIONAL PASS"; statusBg = "#FDEBD0"; statusFont = "#B9770E";
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

function clearWoSummaryCache() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf("WO_META_") === 0) props.deleteProperty(keys[i]);
  }
  Logger.log("Summary persistent metadata cache cleared.");
}

function onSelectionChange(e) {
  if (!e || !e.range) return;
  var range = e.range;
  
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  var sheet = range.getSheet();
  if (sheet.getName() === CONFIG.SHEET_NAMES.SUMMARY) {
    var col = range.getColumn();
    var row = range.getRow();

    if (col === 1 && row > 1) {
      var rawVal = String(sheet.getRange(row, 2).getValue()).trim();
      if (!rawVal || rawVal.startsWith("⚠️") || rawVal.startsWith("❌")) return;

      var woBatch = robustExtractWoBatchNum(rawVal);
      if (woBatch === 0) return;

      var formattedBarcode = "WO-" + woBatch;

      var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
      var opSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);

      if (opSheet) {
        var targetCellKey = (CONFIG.OPERATOR_STATION && CONFIG.OPERATOR_STATION.RANGES && CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) ? CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT : "C3";
        var targetRange = opSheet.getRange(targetCellKey);
        
        ss.setActiveSheet(opSheet);
        targetRange.setValue(formattedBarcode);
        SpreadsheetApp.flush();
        
        if (typeof manageOperatorStation === "function") {
          try {
            manageOperatorStation({
              source: ss,
              range: targetRange,
              value: formattedBarcode,
              oldValue: ""
            });
          } catch (err) {
            Logger.log("Operator Station execution alert: " + err.toString());
          }
        }
      }
    }
  }
}