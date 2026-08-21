// =========================================================================
// 🖥️ USER INTERFACE & CONTROLLERS (UI.js)
// Workspace Rendering, Button Actions & Triggers
// =========================================================================

/**
 * Master Sync Action triggered by button click on Operator Station.
 */
function clickMasterSyncButton() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(err) {}

  try { 
    processDynoFiles(); 
  } catch(e) { 
    Logger.log("Watch folder alert: " + e.toString());
    if (ui) ui.alert("⚠️ Sync Warning", "Process Dyno Files Error: " + e.message, ui.ButtonSet.OK);
  }

  try { 
    retroactiveLogRecalculate(); 
  } catch(e) { 
    Logger.log("Reference Matrix Recalculation Alert: " + e.toString()); 
    if (ui) ui.alert("⚠️ Recalculation Warning", "Retroactive Recalculate Error: " + e.message, ui.ButtonSet.OK);
  }

  try {
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);
    if (sheet) {
      manageOperatorStation({ source: ss, range: sheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) });
    }
  } catch(e) { 
    Logger.log("Console screen alert: " + e.toString()); 
  }
}

/**
 * Helper to safely extract positive absolute numbers.
 */
function safeAbsNum(val) {
  if (val === "" || val === null || val === undefined) return "";
  var n = parseFloat(val);
  return isNaN(n) ? val : Math.abs(n);
}

/**
 * Normalizes string keys by stripping non-alphanumeric characters and lowercasing.
 */
function cleanKey(val) {
  if (val === null || val === undefined) return "";
  return String(val).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Helper to convert long WO barcodes (43081008001979001) into standard short serials (001979-001).
 */
function formatShortSerial(rawSerial, searchBarcode) {
  var str = String(rawSerial || "").trim();
  if (!str) return "";

  if (str.includes("-")) return str;

  if (str.length >= 3) {
    var unit = str.slice(-3);
    var batch = searchBarcode ? String(searchBarcode).trim() : "";
    if (batch.includes("-")) batch = batch.split("-")[0];
    if (batch.includes("_")) batch = batch.split("_")[0];
    if (!isNaN(batch) && batch.length === 4) batch = "00" + batch;

    if (batch) {
      return batch + "-" + unit;
    }
  }
  return str;
}

/**
 * Robust Serial Matcher between Work Order barcodes (e.g. 43081008001979001) and Dyno Log serials (e.g. 001979-001).
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
 * Safely fetches column value from a log row using CONFIG key or 0-indexed fallback.
 */
function getLogVal(row, colConfigProp, defaultIndex) {
  if (colConfigProp && !isNaN(colConfigProp) && colConfigProp > 0) {
    return row[colConfigProp - 1];
  }
  return row[defaultIndex];
}

/**
 * Sets cell A8 text and dynamically formats background & text color based on status.
 */
function setA8Status(sheet, statusMessage) {
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  var a8 = sheet.getRange(ranges.STATUS_BANNER);
  a8.setValue(statusMessage);

  var upper = String(statusMessage).toUpperCase();

  if (upper.includes("COMPLETED AND PASSING")) {
    a8.setBackground("#00C853").setFontColor("#FFFFFF").setFontWeight("bold");
  } else if (upper.includes("FAIL") || upper.includes("ACTION REQUIRED") || upper.includes("NOT FOUND")) {
    a8.setBackground("#D50000").setFontColor("#FFFFFF").setFontWeight("bold");
  } else if (upper.includes("CONDITIONAL") || upper.includes("ATTENTION") || upper.includes("PENDING") || upper.includes("HOLD") || upper.includes("INCOMPLETE")) {
    a8.setBackground("#FFD600").setFontColor("#000000").setFontWeight("bold");
  } else {
    a8.setBackground(null).setFontColor(null).setFontWeight("normal");
  }
}

/**
 * Operator Station Event Manager for Barcode Scanning & Work Order Lookup.
 */
function manageOperatorStation(e) {
  var ss = e ? e.source : SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION); 
  if (!sheet) return;
  
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  var range = e ? e.range : sheet.getRange(ranges.BARCODE_INPUT); 
  var barcode = String(sheet.getRange(ranges.BARCODE_INPUT).getValue()).trim();  
    
  if (range.getA1Notation() === ranges.BARCODE_INPUT) {  
    sheet.getRange(ranges.CLEAR_METADATA_RANGE).clearContent(); 
    sheet.getRange(ranges.CLEAR_PANEL_RANGE).clearContent(); 
    sheet.getRange(ranges.CLEAR_LIMITS_RANGE).clearContent();
    sheet.getRange(ranges.STATUS_BANNER).clearContent().setBackground(null).setFontColor(null);
    
    var tableRange = sheet.getRange(ranges.CLEAR_RESULTS_RANGE);   
    tableRange.clearContent();   
    tableRange.setBackground("#FFFFFF").setFontColor("#000000").setFontWeight("normal");   

    if (!barcode || barcode === "undefined" || barcode === "null") return;  
      
    var searchBarcode = barcode;  
    if (barcode.includes("-")) { searchBarcode = barcode.split("-")[0].trim(); }   
    else if (barcode.includes("_")) { searchBarcode = barcode.split("_")[0].trim(); }  
    if (!isNaN(searchBarcode) && searchBarcode.length === 4) { searchBarcode = "00" + searchBarcode; }  
      
    var searchCriteria = "title contains '" + searchBarcode + "' and (mimeType = 'application/vnd.google-apps.spreadsheet' or mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') and trashed = false";  
    var files = DriveApp.searchFiles(searchCriteria);  
    if (!files.hasNext()) { 
      sheet.getRange(ranges.FILE_LINK_OUTPUT).setValue("❌ Work Order File Not Found: " + searchBarcode); 
      sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("CROSS-CHECK FAILED");
      setA8Status(sheet, "WORK ORDER FILE NOT FOUND");
      return; 
    }  
    
    var file = files.next();  
    var verifiedFileIdStr = file.getId(); 
    var realFileName = file.getName();   
    
    sheet.getRange(ranges.FILE_LINK_OUTPUT).setFormula('=HYPERLINK("' + file.getUrl() + '", "🔗 Open ' + realFileName + '")');   
    sheet.getRange(ranges.CACHED_FILE_ID).setValue(verifiedFileIdStr);  
      
    try {  
      var woSpreadsheet = SpreadsheetApp.openById(verifiedFileIdStr); 
      var woSheet = woSpreadsheet.getSheets()[0];   
      var woPartNumber = String(woSheet.getRange("D3").getValue()).trim(); 
      var woBomRevision = String(woSheet.getRange("D4").getValue()).trim();   
      
      var expectedSerials = [];
      var woLastRow = woSheet.getLastRow();
      if (woLastRow >= 12) {
        var rawSerials = woSheet.getRange(12, 1, woLastRow - 11, 1).getValues();
        for (var sIdx = 0; sIdx < rawSerials.length; sIdx++) {
          var sVal = String(rawSerials[sIdx][0] || "").trim();
          if (sVal && sVal.toLowerCase() !== "undefined" && sVal.toLowerCase() !== "null") {
            expectedSerials.push(sVal);
          }
        }
      }

      sheet.getRange(ranges.BOM_REV_OUTPUT).setValue(woBomRevision); 
      sheet.getRange(ranges.BASE_MODEL_OUTPUT).setValue(woPartNumber);   

      var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);  
      var matchedProgramName = "";
      var matchedDynamicKey = "";
      var matchedBaseModel = "";

      if (registrySheet && woPartNumber) {  
        var regValues = registrySheet.getDataRange().getValues();  
        var cleanWoPart = cleanKey(woPartNumber);  
        var regCols = CONFIG.COLUMNS.PROGRAM_REGISTRY;
        
        var bestMatchRow = null;
        var highestMatchScore = 0;

        for (var rR = 1; rR < regValues.length; rR++) {  
          var regRow = regValues[rR];
          var regProgName = String(regRow[regCols.PROGRAM_NAME - 1] || "").trim();  
          var regDynamicKey = String(regRow[regCols.DYNAMIC_KEY - 1] || "").trim();
          var regBaseModel = String(regRow[regCols.BASE_MODEL - 1] || "").trim();
          
          var cleanProgName = cleanKey(regProgName);
          var cleanDynKey = cleanKey(regDynamicKey);
          var cleanBase = cleanKey(regBaseModel);

          if (!regProgName && !regDynamicKey && !regBaseModel) continue;

          var currentScore = 0;

          // Priority 1: Exact match on Dynamic Key or Program Name
          if ((cleanDynKey && cleanDynKey === cleanWoPart) || (cleanProgName && cleanProgName === cleanWoPart)) {
            currentScore = 3;
          }
          // Priority 2: Partial match on Dynamic Key or Program Name
          else if ((cleanDynKey && cleanWoPart.indexOf(cleanDynKey) !== -1) || (cleanProgName && cleanWoPart.indexOf(cleanProgName) !== -1)) {
            currentScore = 2;
          }
          // Priority 3: Fallback match on Base Model (later rows override earlier rows)
          else if (cleanBase && (cleanBase === cleanWoPart || cleanWoPart.indexOf(cleanBase) !== -1 || cleanBase.indexOf(cleanWoPart) !== -1)) {
            currentScore = 1;
          }

          if (currentScore > 0 && currentScore >= highestMatchScore) {
            highestMatchScore = currentScore;
            bestMatchRow = regRow;
          }
        }

        if (bestMatchRow) {
          matchedProgramName = String(bestMatchRow[regCols.PROGRAM_NAME - 1] || "").trim();
          matchedDynamicKey = String(bestMatchRow[regCols.DYNAMIC_KEY - 1] || "").trim();
          matchedBaseModel = String(bestMatchRow[regCols.BASE_MODEL - 1] || "").trim();

          sheet.getRange(ranges.CUSTOMER_ACCOUNT_OUTPUT).setValue(bestMatchRow[regCols.CUSTOMER_ACCOUNT - 1] || "");
          sheet.getRange(ranges.VEHICLE_SPEC_OUTPUT).setValue(bestMatchRow[regCols.VEHICLE_SPEC - 1] || "");
          sheet.getRange(ranges.PROGRAM_NAME_OUTPUT).setValue(matchedProgramName);
          sheet.getRange(ranges.VALVING_VERSION_OUTPUT).setValue(bestMatchRow[regCols.VALVING_VERSION - 1] || "");
          sheet.getRange(ranges.ADJUSTER_TARGETS_OUTPUT).setValue(bestMatchRow[regCols.ADJUSTER_SETTINGS - 1] || "");
          
          if (matchedBaseModel) {
            sheet.getRange(ranges.BASE_MODEL_OUTPUT).setValue(matchedBaseModel);
          }

          sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("Validated in Registry");
        }  
      }  

      if (!matchedProgramName) {
        sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("Registry Match Pending");
      }

      populateSpecLimits(ss, sheet, matchedDynamicKey || matchedProgramName || woPartNumber);
      renderOperatorTableWithFormatting(ss, sheet, searchBarcode, woPartNumber, expectedSerials);

    } catch(e) {
      Logger.log("WO Lookup Error: " + e.toString());
      sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("Extraction Error");
    }
  }  
}  

/**
 * Direct lookup helper for spec limits from Part_Reference_Matrix.
 */
function getSpecLimitsFromMatrix(ss, dynamicKeyOrPart) {
  var limits = { c1Min: NaN, c1Max: NaN, r1Min: NaN, r1Max: NaN, c2Min: NaN, c2Max: NaN, r2Min: NaN, r2Max: NaN, slopeMin: NaN };
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);
  if (!refSheet || !dynamicKeyOrPart) return limits;

  var refData = refSheet.getDataRange().getValues();
  var refCols = CONFIG.COLUMNS.PART_REFERENCE_MATRIX;
  var targetCleanKey = cleanKey(dynamicKeyOrPart);

  var refRow = null;
  for (var i = 1; i < refData.length; i++) {
    var matrixKey = cleanKey(refData[i][refCols.DYNAMIC_KEY - 1]);
    if (matrixKey === targetCleanKey || matrixKey.includes(targetCleanKey) || targetCleanKey.includes(matrixKey)) {
      refRow = refData[i];
      break;
    }
  }

  if (!refRow) return limits;

  function getAbsPair(minVal, maxVal) {
    if (minVal === "" || maxVal === "" || minVal === null || maxVal === null) return { min: NaN, max: NaN };
    var a = Math.abs(parseFloat(minVal));
    var b = Math.abs(parseFloat(maxVal));
    if (isNaN(a) || isNaN(b)) return { min: Math.min(a, b), max: Math.max(a, b) };
    return { min: Math.min(a, b), max: Math.max(a, b) };
  }

  var c1 = getAbsPair(refRow[refCols.COMP_1_MIN - 1], refRow[refCols.COMP_1_MAX - 1]);
  var r1 = getAbsPair(refRow[refCols.REB_1_MIN - 1], refRow[refCols.REB_1_MAX - 1]);
  var c2 = getAbsPair(refRow[refCols.COMP_2_MIN - 1], refRow[refCols.COMP_2_MAX - 1]);
  var r2 = getAbsPair(refRow[refCols.REB_2_MIN - 1], refRow[refCols.REB_2_MAX - 1]);
  var slope = parseFloat(refRow[refCols.SLOPE_1_MIN - 1]);

  return {
    c1Min: c1.min, c1Max: c1.max,
    r1Min: r1.min, r1Max: r1.max,
    c2Min: c2.min, c2Max: c2.max,
    r2Min: r2.min, r2Max: r2.max,
    slopeMin: isNaN(slope) ? NaN : Math.abs(slope)
  };
}

/**
 * Populates spec limit cells (B22:F23) with positive absolute values.
 */
function populateSpecLimits(ss, sheet, dynamicKeyOrPart) {
  var limits = getSpecLimitsFromMatrix(ss, dynamicKeyOrPart);
  var ranges = CONFIG.OPERATOR_STATION.RANGES;

  sheet.getRange(ranges.LIMIT_COMP_1_MIN).setValue(isNaN(limits.c1Min) ? "" : limits.c1Min);
  sheet.getRange(ranges.LIMIT_COMP_1_MAX).setValue(isNaN(limits.c1Max) ? "" : limits.c1Max);
  sheet.getRange(ranges.LIMIT_REB_1_MIN).setValue(isNaN(limits.r1Min) ? "" : limits.r1Min);
  sheet.getRange(ranges.LIMIT_REB_1_MAX).setValue(isNaN(limits.r1Max) ? "" : limits.r1Max);

  sheet.getRange(ranges.LIMIT_COMP_2_MIN).setValue(isNaN(limits.c2Min) ? "" : limits.c2Min);
  sheet.getRange(ranges.LIMIT_COMP_2_MAX).setValue(isNaN(limits.c2Max) ? "" : limits.c2Max);
  sheet.getRange(ranges.LIMIT_REB_2_MIN).setValue(isNaN(limits.r2Min) ? "" : limits.r2Min);
  sheet.getRange(ranges.LIMIT_REB_2_MAX).setValue(isNaN(limits.r2Max) ? "" : limits.r2Max);

  sheet.getRange(ranges.LIMIT_SLOPE_1_MIN).setValue(isNaN(limits.slopeMin) ? "No Limit" : limits.slopeMin.toFixed(1));
}

/**
 * Queries Master_Dyno_Log, cross-references against Work Order expected serials, and renders table A27 downward.
 */
function renderOperatorTableWithFormatting(ss, sheet, searchBarcode, partNumber, expectedSerials) {
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  if (!logSheet) return;

  var logData = logSheet.getDataRange().getValues();
  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG || {};
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  var logSheetId = logSheet.getSheetId();

  var hMap = buildHeaderMap(logData[0] || []);

  var limits = {
    c1Min: parseFloat(sheet.getRange(ranges.LIMIT_COMP_1_MIN).getValue()),
    c1Max: parseFloat(sheet.getRange(ranges.LIMIT_COMP_1_MAX).getValue()),
    r1Min: parseFloat(sheet.getRange(ranges.LIMIT_REB_1_MIN).getValue()),
    r1Max: parseFloat(sheet.getRange(ranges.LIMIT_REB_1_MAX).getValue()),
    c2Min: parseFloat(sheet.getRange(ranges.LIMIT_COMP_2_MIN).getValue()),
    c2Max: parseFloat(sheet.getRange(ranges.LIMIT_COMP_2_MAX).getValue()),
    r2Min: parseFloat(sheet.getRange(ranges.LIMIT_REB_2_MIN).getValue()),
    r2Max: parseFloat(sheet.getRange(ranges.LIMIT_REB_2_MAX).getValue())
  };

  if (isNaN(limits.c1Min) || isNaN(limits.c1Max)) {
    var directLimits = getSpecLimitsFromMatrix(ss, partNumber || searchBarcode);
    if (!isNaN(directLimits.c1Min)) limits = directLimits;
  }

  var cleanBarcodeStr = cleanKey(searchBarcode);
  var cleanPartStr = cleanKey(partNumber);

  var latestLogBySerial = {};

  if (logData.length > 1) {
    for (var r = 1; r < logData.length; r++) {
      var row = logData[r];
      var trueSerial = String(getLogVal(row, logCols.TRUE_SERIAL, 2) || "").trim();
      var baseModel = cleanKey(getLogVal(row, logCols.BASE_MODEL, 3));
      var cleanSerial = cleanKey(trueSerial);

      var isMatch = false;
      if (cleanBarcodeStr !== "" && cleanSerial.includes(cleanBarcodeStr)) {
        isMatch = true;
      } else if (cleanPartStr !== "" && baseModel === cleanPartStr && (cleanBarcodeStr === "" || cleanBarcodeStr === "undefined")) {
        isMatch = true;
      }

      if (isMatch && trueSerial !== "") {
        latestLogBySerial[cleanSerial] = {
          rowIdx: r + 1,
          data: row,
          trueSerial: trueSerial
        };
      }
    }
  }

  var itemsToProcess = [];
  var matchedLogKeys = {};

  if (expectedSerials && expectedSerials.length > 0) {
    for (var eIdx = 0; eIdx < expectedSerials.length; eIdx++) {
      var expSerial = expectedSerials[eIdx];
      var matchedLogItem = null;

      var logSerialKeys = Object.keys(latestLogBySerial);
      for (var lIdx = 0; lIdx < logSerialKeys.length; lIdx++) {
        var logKey = logSerialKeys[lIdx];
        var logItem = latestLogBySerial[logKey];
        if (isSerialMatch(expSerial, logItem.trueSerial)) {
          matchedLogItem = logItem;
          matchedLogKeys[logKey] = true;
          break;
        }
      }

      if (matchedLogItem) {
        itemsToProcess.push({
          rawSerial: matchedLogItem.trueSerial,
          isTested: true,
          data: matchedLogItem.data,
          rowIdx: matchedLogItem.rowIdx
        });
      } else {
        itemsToProcess.push({
          rawSerial: expSerial,
          isTested: false,
          data: null,
          rowIdx: -1
        });
      }
    }
  }

  var logSerialKeys = Object.keys(latestLogBySerial);
  for (var lIdx = 0; lIdx < logSerialKeys.length; lIdx++) {
    var k = logSerialKeys[lIdx];
    if (!matchedLogKeys[k]) {
      itemsToProcess.push({
        rawSerial: latestLogBySerial[k].trueSerial,
        isTested: true,
        data: latestLogBySerial[k].data,
        rowIdx: latestLogBySerial[k].rowIdx
      });
    }
  }

  // Pre-format short display serial for each item prior to sorting
  for (var i = 0; i < itemsToProcess.length; i++) {
    itemsToProcess[i].displaySerial = formatShortSerial(itemsToProcess[i].rawSerial, searchBarcode);
  }

  // Natural numeric sort on standardized displaySerial (e.g. 001979-001, 001979-002, 001979-010)
  itemsToProcess.sort(function(a, b) {
    var sA = a.displaySerial || "";
    var sB = b.displaySerial || "";
    var cmp = sA.localeCompare(sB, undefined, { numeric: true, sensitivity: "base" });
    if (cmp !== 0) return cmp;
    var mA = sA.match(/(\d+)$/);
    var mB = sB.match(/(\d+)$/);
    var uA = mA ? parseInt(mA[1], 10) : 0;
    var uB = mB ? parseInt(mB[1], 10) : 0;
    return uA - uB;
  });

  var rowsToDisplay = [];
  var test1FailCount = 0;
  var test2FailCount = 0;
  var holdCount = 0;
  var testedCount = 0;
  var untestedCount = 0;

  for (var i = 0; i < itemsToProcess.length; i++) {
    var item = itemsToProcess[i];
    var displaySerial = item.displaySerial || formatShortSerial(item.rawSerial, searchBarcode);

    if (item.isTested) {
      testedCount++;
      var row = item.data;
      var actualSheetRow = item.rowIdx;

      var rowLink = "#gid=" + logSheetId + "&range=A" + actualSheetRow;
      var serialHyperlinkFormula = '=HYPERLINK("' + rowLink + '", "' + displaySerial + '")';

      var t1Status    = String(row[hMap.test1Status] || "").trim();
      var t2Status    = String(row[hMap.test2Status] || "").trim();
      var overallStat = String(row[hMap.overallStatus] || "").trim();
      var diagnostics = String(row[hMap.diagnostics] || "").trim();
      var evalAction  = String(row[hMap.evaluationAction] || "").trim();
      var engComm     = String(row[hMap.engComments] || "").trim();

      if (overallStat.toUpperCase().includes("HOLD")) holdCount++;
      if (t1Status.toUpperCase().includes("FAIL")) test1FailCount++;
      if (t2Status.toUpperCase().includes("FAIL")) test2FailCount++;

      var mappedRow = [
        serialHyperlinkFormula,                           // Col A (1): Standardized Short Serial
        safeAbsNum(getLogVal(row, logCols.ROD_FORCE, 5)), // Col B (2): Rod Force
        safeAbsNum(getLogVal(row, logCols.COMP_1, 7)),    // Col C (3): Low Speed Comp
        safeAbsNum(getLogVal(row, logCols.REB_1, 8)),     // Col D (4): Low Speed Reb
        safeAbsNum(getLogVal(row, logCols.COMP_2, 12)),   // Col E (5): Med Speed Comp
        safeAbsNum(getLogVal(row, logCols.REB_2, 13)),    // Col F (6): Med Speed Reb
        t1Status,                                         // Col G (7): Test 1 Status
        t2Status,                                         // Col H (8): Test 2 Status
        overallStat,                                      // Col I (9): Overall Status
        evalAction,                                       // Col J (10): Evaluation Action
        diagnostics,                                      // Col K (11): Diagnostics
        engComm                                           // Col L (12): Engineering Comments
      ];
      rowsToDisplay.push(mappedRow);
    } else {
      untestedCount++;
      var mappedRow = [
        displaySerial,                                    // Col A (1): Standardized Short Serial
        "",                                               // Col B (2): Rod Force
        "",                                               // Col C (3): Low Speed Comp
        "",                                               // Col D (4): Low Speed Reb
        "",                                               // Col E (5): Med Speed Comp
        "",                                               // Col F (6): Med Speed Reb
        "NOT TESTED YET",                                 // Col G (7): Test 1 Status
        "NOT TESTED YET",                                 // Col H (8): Test 2 Status
        "NOT TESTED YET",                                 // Col I (9): Overall Status
        "",                                               // Col J (10): Evaluation Action
        "⏳ Unit pending dyno test.",                     // Col K (11): Diagnostics
        ""                                                // Col L (12): Engineering Comments
      ];
      rowsToDisplay.push(mappedRow);
    }
  }

  var statusMessage = "";
  var totalCount = itemsToProcess.length;

  if (totalCount === 0) {
    statusMessage = "PENDING TESTING";
  } else if (testedCount === 0) {
    statusMessage = "PENDING TESTING: 0 of " + totalCount + " Tested";
  } else if (untestedCount > 0) {
    statusMessage = "INCOMPLETE WORK ORDER: " + testedCount + " of " + totalCount + " Tested (" + untestedCount + " Unit(s) Missing)";
  } else if (holdCount > 0) {
    statusMessage = "HOLD: Action Required (" + holdCount + " unit(s) marked for Retest/Teardown)";
  } else if (test1FailCount > 0) {
    statusMessage = "ACTION REQUIRED: Test 1 Failure Detected (" + test1FailCount + " unit(s))";
  } else if (test2FailCount > 0) {
    statusMessage = "CONDITIONAL PASS: Attention Required (Test 2 Outlier Detected - " + test2FailCount + " unit(s))";
  } else {
    statusMessage = "WORK ORDER COMPLETED AND PASSING";
  }

  setA8Status(sheet, statusMessage);

  if (rowsToDisplay.length === 0) return;

  var startRow = ranges.RESULTS_START_ROW;
  var numRows = rowsToDisplay.length;
  var numCols = ranges.RESULTS_COL_COUNT;
  var outputRange = sheet.getRange(startRow, ranges.RESULTS_START_COL, numRows, numCols);

  outputRange.setValues(rowsToDisplay);
  sheet.getRange(startRow, 2, numRows, 5).setNumberFormat("0.0");

  var bgColors = [];
  var fontColors = [];
  var fontWeights = [];

  for (var rIdx = 0; rIdx < numRows; rIdx++) {
    var rowBg = ["#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF"];
    var rowFont = ["#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000"];
    var rowWeight = ["normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal"];
    var rowData = rowsToDisplay[rIdx];

    var t1StatusStr    = String(rowData[6] || "").toUpperCase();
    var t2StatusStr    = String(rowData[7] || "").toUpperCase();
    var overallStatStr = String(rowData[8] || "").toUpperCase();
    var diagnosticsStr = String(rowData[10] || "");

    // Helper: Test 1 Red Highlight (Blueprint / Spec Limit Fails)
    var applyRedHighlight = function(colIndex) {
      rowBg[colIndex] = "#FADBD8";   // Soft Red
      rowFont[colIndex] = "#C0392B"; // Dark Red
      rowWeight[colIndex] = "bold";
    };

    // Helper: Test 2 Yellow Highlight (Outlier / Cohort Fails)
    var applyYellowHighlight = function(colIndex) {
      if (rowBg[colIndex] !== "#FADBD8") {
        rowBg[colIndex] = "#FCF3CF";   // Soft Yellow
        rowFont[colIndex] = "#B9770E"; // Dark Yellow/Gold
        rowWeight[colIndex] = "bold";
      }
    };

    // 1. Direct Numeric Measurement vs Target Range Checks (Test 1 Spec Limits)
    var c1Val = parseFloat(rowData[2]);
    var r1Val = parseFloat(rowData[3]);
    var c2Val = parseFloat(rowData[4]);
    var r2Val = parseFloat(rowData[5]);

    if (!isNaN(c1Val) && ((!isNaN(limits.c1Min) && c1Val < limits.c1Min) || (!isNaN(limits.c1Max) && c1Val > limits.c1Max))) applyRedHighlight(2);
    if (!isNaN(r1Val) && ((!isNaN(limits.r1Min) && r1Val < limits.r1Min) || (!isNaN(limits.r1Max) && r1Val > limits.r1Max))) applyRedHighlight(3);
    if (!isNaN(c2Val) && ((!isNaN(limits.c2Min) && c2Val < limits.c2Min) || (!isNaN(limits.c2Max) && c2Val > limits.c2Max))) applyRedHighlight(4);
    if (!isNaN(r2Val) && ((!isNaN(limits.r2Min) && r2Val < limits.r2Min) || (!isNaN(limits.r2Max) && r2Val > limits.r2Max))) applyRedHighlight(5);

    // 2. Diagnostic Tag Highlighting
    if (!diagnosticsStr.includes("✅") && !diagnosticsStr.includes("⏳")) {
      var isT1Fail = t1StatusStr.includes("FAIL");
      var isT2Fail = t2StatusStr.includes("FAIL");
      var activeHighlightFunc = isT1Fail ? applyRedHighlight : (isT2Fail ? applyYellowHighlight : applyRedHighlight);

      if (diagnosticsStr.indexOf("[RF_FAIL]") !== -1)    activeHighlightFunc(1);
      if (diagnosticsStr.indexOf("[C1_FAIL]") !== -1)    activeHighlightFunc(2);
      if (diagnosticsStr.indexOf("[R1_FAIL]") !== -1)    activeHighlightFunc(3);
      if (diagnosticsStr.indexOf("[C2_FAIL]") !== -1)    activeHighlightFunc(4);
      if (diagnosticsStr.indexOf("[R2_FAIL]") !== -1)    activeHighlightFunc(5);
      if (diagnosticsStr.indexOf("[SLOPE_FAIL]") !== -1) { activeHighlightFunc(2); activeHighlightFunc(3); }
    }

    // 3. Test 1 Status Column (Col G / Index 6)
    if (t1StatusStr.includes("FAIL")) {
      applyRedHighlight(6);
    } else if (t1StatusStr.includes("PASS")) {
      rowBg[6] = "#D4EFDF"; rowFont[6] = "#196F3D";
    } else if (t1StatusStr.includes("NOT TESTED")) {
      rowBg[6] = "#F2F4F4"; rowFont[6] = "#5D6D7E";
    }

    // 4. Test 2 Status Column (Col H / Index 7) - Yellow for Test 2 Outliers
    if (t2StatusStr.includes("FAIL")) {
      rowBg[7] = "#FCF3CF"; rowFont[7] = "#B9770E"; rowWeight[7] = "bold";
    } else if (t2StatusStr.includes("PASS")) {
      rowBg[7] = "#D4EFDF"; rowFont[7] = "#196F3D";
    } else if (t2StatusStr.includes("NOT TESTED")) {
      rowBg[7] = "#F2F4F4"; rowFont[7] = "#5D6D7E";
    }

    // 5. Overall Status Column (Col I / Index 8)
    if (overallStatStr.includes("FAIL")) {
      if (t1StatusStr.includes("FAIL")) {
        rowBg[8] = "#C0392B"; rowFont[8] = "#FFFFFF"; rowWeight[8] = "bold"; // Solid Red
      } else {
        rowBg[8] = "#FCF3CF"; rowFont[8] = "#B9770E"; rowWeight[8] = "bold"; // Soft Yellow
      }
    } else if (overallStatStr.includes("HOLD") || overallStatStr.includes("NOT TESTED")) {
      rowBg[8] = "#FCF3CF"; rowFont[8] = "#B7950B"; rowWeight[8] = "bold";
    } else if (overallStatStr.includes("OVERRIDE") || overallStatStr.includes("PASS")) {
      rowBg[8] = "#D4EFDF"; rowFont[8] = "#196F3D"; rowWeight[8] = "bold";
    }

    bgColors.push(rowBg);
    fontColors.push(rowFont);
    fontWeights.push(rowWeight);
  }

  outputRange.setBackgrounds(bgColors).setFontColors(fontColors).setFontWeights(fontWeights);
}

/**
 * Standard Simple Trigger for Google Sheet edits.
 */
function onEdit(e) {
  installableOnEdit(e);
}

/**
 * Main Controller Engine for spreadsheet edits.
 */
function installableOnEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  if (sheetName === (CONFIG.SHEET_NAMES.SERIAL_HISTORY_VIEWER || "Serial_History_Viewer")) {
    var searchCell = (CONFIG.SERIAL_HISTORY_VIEWER && CONFIG.SERIAL_HISTORY_VIEWER.RANGES) ? CONFIG.SERIAL_HISTORY_VIEWER.RANGES.SERIAL_SEARCH_INPUT : "B2";
    if (e.range.getA1Notation() === searchCell) {
      if (typeof renderSerialHistory === "function") {
        renderSerialHistory(e);
      }
    }
    return;
  }

  if (sheetName === CONFIG.SHEET_NAMES.OPERATOR_STATION) {
    try { manageOperatorStation(e); } catch(err) { Logger.log("Operator station edit error: " + err.toString()); }
    return;
  }

  if (sheetName === CONFIG.SHEET_NAMES.MASTER_DYNO_LOG) {
    var editedCol = e.range.getColumn();

    if (editedCol >= 22 && editedCol <= 26) {
      retroactiveLogRecalculate();

      var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
      var opSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);
      if (opSheet) {
        try {
          manageOperatorStation({
            source: ss,
            range: opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT)
          });
        } catch(err) {
          Logger.log("Notice: UI refresh skipped.");
        }
      }
    }
  }
}

/**
 * Automated Trigger Setup Helper.
 */
function setupTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var triggers = ScriptApp.getUserTriggers(ss);
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "installableOnEdit") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("installableOnEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log("Successfully installed OnEdit trigger for installableOnEdit.");
}

/**
 * Trigger wrapper for scheduled background syncs.
 */
function syncDynoAndRefreshWO() {
  clickMasterSyncButton();
}