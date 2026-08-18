// =========================================================================
// 📊 SUMMARY DASHBOARD CONTROLLER (Summary.js)
// Registry-Filtered & Exact UI.js Engine Mirroring Controller
// =========================================================================

/**
 * Normalizes string keys by stripping non-alphanumeric characters and lowercasing.
 * Exact copy from UI.js.
 */
function cleanKey(val) {
  if (val === null || val === undefined) return "";
  return String(val).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Robustly extracts pure WO batch number from barcodes or file names.
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
 * Robust Serial Matcher between Work Order barcodes and Dyno Log serials.
 * Exact copy from UI.js.
 */
function isSerialMatch(expSerial, logSerial) {
  var cExp = cleanKey(expSerial);
  var cLog = cleanKey(logSerial);
  if (!cExp || !cLog) return false;

  if (cExp === cLog) return true;
  if (cExp.endsWith(cLog) || cLog.endsWith(cExp)) return true;

  var cLogNoZero = cLog.replace(/^0+/, "");
  var cExpNoZero = cExp.replace(/^0+/, "");
  if (cExp.endsWith(cLogNoZero) || cLog.endsWith(cExpNoZero)) return true;

  if (cExp.length >= 6 && cLog.length >= 3) {
    var expUnit = cExp.slice(-3);
    var logUnit = cLog.slice(-3);
    if (expUnit === logUnit) {
      var logBatch = cLog.slice(0, -3).replace(/^0+/, "");
      if (logBatch && cExp.indexOf(logBatch) !== -1) return true;
    }
  }

  return false;
}

/**
 * Builds header map for Summary Dashboard columns specifically.
 * Renamed buildSummaryHeaderMap to avoid global function collision with Engine.js.
 */
function buildSummaryHeaderMap(headers) {
  var map = {
    trueSerial: 2,      // Col C (default)
    baseModel: 3,       // Col D (default)
    test1Status: 6,     // Col G (default)
    test2Status: 7,     // Col H (default)
    overallStatus: 8,   // Col I (default)
    diagnostics: 10,    // Col K (default)
    timestamp: 0        // Col A (default)
  };

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().replace(/[^a-z0-9]/g, "");
    if (h.includes("trueserial") || (h.includes("serial") && !h.includes("type"))) map.trueSerial = i;
    if (h.includes("basemodel") || h.includes("partnumber")) map.baseModel = i;
    if (h.includes("test1") || h.includes("comp1status") || h.includes("t1status")) map.test1Status = i;
    if (h.includes("test2") || h.includes("comp2status") || h.includes("t2status")) map.test2Status = i;
    if (h.includes("overall") || h.includes("overallstatus")) map.overallStatus = i;
    if (h.includes("diag") || h.includes("diagnostics")) map.diagnostics = i;
    if (h.includes("time") || h.includes("date")) map.timestamp = i;
  }
  return map;
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
    if (bmVal) validBaseModels[cleanKey(bmVal)] = true;
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

      var isTrackedPart = (Object.keys(validModels).length === 0) || validModels[cleanKey(baseModel)];

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

/**
 * Builds Summary Dashboard mirroring UI.js cell A8 decision logic and renderOperatorTableWithFormatting().
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
  var sumCols = CONFIG.COLUMNS.SUMMARY || {};
  var BASELINE_WO_FLOOR = 1608;

  var hMap = buildSummaryHeaderMap(logData[0] || []);

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
    if (baseModel !== "PENDING CACHE" && Object.keys(validModels).length > 0 && !validModels[cleanKey(baseModel)]) {
      continue;
    }

    var searchBarcode = String(item.woNum || woNumber).trim();
    var cleanBarcodeStr = cleanKey(searchBarcode);
    var cleanPartStr = cleanKey(baseModel);

    var latestLogBySerial = {};

    if (logData.length > 1) {
      for (var r = 1; r < logData.length; r++) {
        var row = logData[r];
        var trueSerial = String(row[hMap.trueSerial] || "").trim();
        var logBaseModel = cleanKey(row[hMap.baseModel]);
        var cleanSerial = cleanKey(trueSerial);

        var isMatch = false;
        if (cleanBarcodeStr !== "" && cleanSerial.includes(cleanBarcodeStr)) {
          isMatch = true;
        } else if (cleanPartStr !== "" && logBaseModel === cleanPartStr && (cleanBarcodeStr === "" || cleanBarcodeStr === "undefined")) {
          isMatch = true;
        }

        if (isMatch && trueSerial !== "") {
          latestLogBySerial[cleanSerial] = {
            data: row,
            trueSerial: trueSerial
          };
        }
      }
    }

    var totalQty = expectedSerials.length;
    var testedCount = 0;
    var untestedCount = 0;
    var test1FailCount = 0;
    var test2FailCount = 0;
    var holdCount = 0;
    var activeFailureDetails = [];
    var lastDate = null;

    for (var es = 0; es < expectedSerials.length; es++) {
      var expS = expectedSerials[es];
      var matchedLogItem = null;

      var logSerialKeys = Object.keys(latestLogBySerial);
      for (var lIdx = 0; lIdx < logSerialKeys.length; lIdx++) {
        var logKey = logSerialKeys[lIdx];
        var logItem = latestLogBySerial[logKey];
        if (isSerialMatch(expS, logItem.trueSerial)) {
          matchedLogItem = logItem;
          break;
        }
      }

      if (matchedLogItem) {
        testedCount++;
        var row = matchedLogItem.data;

        var t1Status    = String(row[hMap.test1Status] || "").trim().toUpperCase();
        var t2Status    = String(row[hMap.test2Status] || "").trim().toUpperCase();
        var overallStat = String(row[hMap.overallStatus] || "").trim().toUpperCase();
        var diagnostics = String(row[hMap.diagnostics] || "").trim();
        var runDate     = row[hMap.timestamp];

        if (runDate instanceof Date && (!lastDate || runDate > lastDate)) {
          lastDate = runDate;
        }

        if (overallStat.includes("HOLD")) {
          holdCount++;
          activeFailureDetails.push("#" + expS.slice(-3) + " [HOLD]");
        }
        if (t1Status.includes("FAIL")) {
          test1FailCount++;
          var tagMatch = diagnostics.match(/\[(.*?)\]/);
          activeFailureDetails.push("#" + expS.slice(-3) + " " + (tagMatch ? tagMatch[0] : "[T1 FAIL]"));
        }
        if (t2Status.includes("FAIL")) {
          test2FailCount++;
          var tagMatch2 = diagnostics.match(/\[(.*?)\]/);
          activeFailureDetails.push("#" + expS.slice(-3) + " " + (tagMatch2 ? tagMatch2[0] : "[T2 OUTLIER]"));
        }
      } else {
        untestedCount++;
      }
    }

    var woStatus = "PENDING";
    var statusBg = "#FCF3CF";
    var statusFont = "#B7950B";

    var totalCount = totalQty;

    if (baseModel === "PENDING CACHE") {
      woStatus = "PENDING"; statusBg = "#FCF3CF"; statusFont = "#B7950B";
    } else if (totalCount === 0) {
      woStatus = "NO SERIALS"; statusBg = "#F2F4F4"; statusFont = "#5D6D7E";
    } else if (testedCount === 0) {
      woStatus = "PENDING"; statusBg = "#FCF3CF"; statusFont = "#B7950B";
    } else if (holdCount > 0) {
      woStatus = "HOLD"; statusBg = "#FCF3CF"; statusFont = "#B7950B";
    } else if (test1FailCount > 0) {
      woStatus = "ACTION REQUIRED"; statusBg = "#FADBD8"; statusFont = "#C0392B";
    } else if (test2FailCount > 0) {
      woStatus = "CONDITIONAL PASS"; statusBg = "#FDEBD0"; statusFont = "#B9770E";
    } else if (untestedCount > 0) {
      woStatus = "INCOMPLETE"; statusBg = "#FCF3CF"; statusFont = "#B7950B";
    } else {
      woStatus = "COMPLETED"; statusBg = "#D4EFDF"; statusFont = "#196F3D";
    }

    var progressStr = (baseModel === "PENDING CACHE") ? "Indexing..." : testedCount + " / " + totalQty + " (" + (totalQty > 0 ? Math.round((testedCount / totalQty) * 100) : 0) + "%)";
    var fpyStr = testedCount > 0 ? (((testedCount - test1FailCount) / testedCount) * 100).toFixed(1) + "%" : "N/A";
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

/**
 * Interactive Column 1 Click Navigator.
 * Populates pure numeric batch digits (e.g. 1905, 1634) into cell C3.
 */
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

      var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
      var opSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);

      if (opSheet) {
        var targetCellKey = (CONFIG.OPERATOR_STATION && CONFIG.OPERATOR_STATION.RANGES && CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) ? CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT : "C3";
        var targetRange = opSheet.getRange(targetCellKey);
        
        var pureBatchStr = String(woBatch);

        ss.setActiveSheet(opSheet);
        targetRange.setValue(pureBatchStr);
        SpreadsheetApp.flush();
        
        if (typeof manageOperatorStation === "function") {
          try {
            manageOperatorStation({
              source: ss,
              range: targetRange,
              value: pureBatchStr,
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