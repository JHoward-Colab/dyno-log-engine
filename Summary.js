// =========================================================================
// 📊 SUMMARY DASHBOARD CONTROLLER (Summary.js)
// High-Speed Built-In Advanced Service Summary Engine
// =========================================================================

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
 * High-speed metadata reader using Advanced Sheets API (50ms per file).
 */
function fetchWoMetadataFast(fileId) {
  try {
    if (typeof Sheets !== "undefined") {
      var res = Sheets.Spreadsheets.Values.batchGet(fileId, {
        ranges: ["D3:D4", "A12:A100"]
      });
      var vRanges = res.valueRanges || [];
      var d3d4 = (vRanges[0] && vRanges[0].values) ? vRanges[0].values : [];
      var a12a100 = (vRanges[1] && vRanges[1].values) ? vRanges[1].values : [];

      var baseModel = (d3d4[0] && d3d4[0][0]) ? String(d3d4[0][0]).trim() : "";
      var bomRev = (d3d4[1] && d3d4[1][0]) ? String(d3d4[1][0]).trim() : "";
      var expectedSerials = [];

      for (var s = 0; s < a12a100.length; s++) {
        var sVal = (a12a100[s] && a12a100[s][0]) ? String(a12a100[s][0]).trim() : "";
        if (sVal && sVal.toLowerCase() !== "undefined" && sVal.toLowerCase() !== "null") {
          expectedSerials.push(sVal);
        }
      }
      return { bm: baseModel, br: bomRev, es: expectedSerials };
    }
  } catch (e) {
    Logger.log("Advanced Service bypass for " + fileId + ": " + e.toString());
  }

  // Fallback if Advanced Service is disabled
  var woSs = SpreadsheetApp.openById(fileId);
  var woSheet = woSs.getSheets()[0];
  var bm = String(woSheet.getRange("D3").getValue()).trim();
  var br = String(woSheet.getRange("D4").getValue()).trim();
  var woLastRow = woSheet.getLastRow();
  var serials = [];
  if (woLastRow >= 12) {
    var raw = woSheet.getRange(12, 1, woLastRow - 11, 1).getValues();
    for (var i = 0; i < raw.length; i++) {
      var v = String(raw[i][0] || "").trim();
      if (v && v.toLowerCase() !== "undefined" && v.toLowerCase() !== "null") serials.push(v);
    }
  }
  return { bm: bm, br: br, es: serials };
}

function buildSummaryDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);

  if (!summarySheet || !logSheet) return;

  var folderId = CONFIG.FOLDERS.WORK_ORDER_FOLDER_ID;
  if (!folderId) return;

  var folder = DriveApp.getFolderById(folderId);
  var logData = logSheet.getDataRange().getValues();
  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG || {};
  var sumCols = CONFIG.COLUMNS.SUMMARY || {};

  var BASELINE_WO_FLOOR = 1608;

  // Step 1: Group Dyno Runs Chronologically
  var allRunsBySerial = {};
  for (var r = 1; r < logData.length; r++) {
    var rawSerial = String(logData[r][(logCols.TRUE_SERIAL || 3) - 1] || "").trim();
    if (rawSerial) {
      var cSer = cleanKey(rawSerial);
      if (!allRunsBySerial[cSer]) allRunsBySerial[cSer] = [];
      allRunsBySerial[cSer].push(logData[r]);
    }
  }

  // Step 2: Load In-Memory Cache
  var propsService = PropertiesService.getScriptProperties();
  var allProps = propsService.getProperties();
  var newPropsToSave = {};

  // Step 3: Fast File Enumeration
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

  // Step 4: Sort Ascending
  fileList.sort(function(a, b) {
    return (a.woNum !== b.woNum && a.woNum > 0 && b.woNum > 0) ? (a.woNum - b.woNum) : a.name.localeCompare(b.name);
  });

  var tableOutput = [];
  var bgColors = [];
  var fontColors = [];
  var fontWeights = [];

  // Step 5: Process Matrix
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

    if (cachedStr) {
      var cachedData = JSON.parse(cachedStr);
      baseModel = cachedData.bm || "";
      bomRev = cachedData.br || "";
      expectedSerials = cachedData.es || [];
    } else {
      var meta = fetchWoMetadataFast(fileId);
      baseModel = meta.bm;
      bomRev = meta.br;
      expectedSerials = meta.es;

      var payloadStr = JSON.stringify(meta);
      newPropsToSave[propKey] = payloadStr;
      allProps[propKey] = payloadStr;
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
      var runs = allRunsBySerial[cExp] || [];

      if (runs.length > 0) {
        testedCount++;

        var firstRun = runs[0];
        if (String(firstRun[(logCols.OVERALL_STATUS || 21) - 1] || "").toUpperCase().includes("PASS")) {
          firstPassCount++;
        }

        var latestRun = runs[runs.length - 1];
        var latestOverall = String(latestRun[(logCols.OVERALL_STATUS || 21) - 1] || "").toUpperCase();
        var latestDiag = String(latestRun[(logCols.DIAGNOSTICS || 22) - 1] || "");
        var runDate = latestRun[(logCols.TIMESTAMP || 1) - 1];

        if (runDate instanceof Date && (!lastDate || runDate > lastDate)) lastDate = runDate;

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

    if (totalQty === 0) {
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

    var progressStr = testedCount + " / " + totalQty + " (" + (totalQty > 0 ? Math.round((testedCount / totalQty) * 100) : 0) + "%)";
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

  if (Object.keys(newPropsToSave).length > 0) {
    propsService.setProperties(newPropsToSave, false);
  }

  // Step 6: Single Bulk Output
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
  var sheet = e.range.getSheet();
  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (sheet.getName() === CONFIG.SHEET_NAMES.SUMMARY && (col === 1 || col === 2) && row > 1) {
    var rawVal = String(sheet.getRange(row, 2).getValue()).trim();
    if (!rawVal || rawVal.startsWith("⚠️") || rawVal.startsWith("❌")) return;

    var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
    var opSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);

    if (opSheet) {
      opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT).setValue(rawVal);
      manageOperatorStation({ source: ss, range: opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) });
      ss.setActiveSheet(opSheet);
    }
  }
}