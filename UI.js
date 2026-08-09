// =========================================================================
// 🖥️ USER INTERFACE & CONTROLLERS (UI.js)
// Workspace Rendering, Button Actions & Triggers
// =========================================================================

/**
 * Master Sync Action triggered by button click on Operator Station.
 */
function clickMasterSyncButton() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();[cite: 1]
  try { processDynoFiles(); } catch(e) { Logger.log("Watch folder alert: " + e.toString()); }[cite: 1]
  try { retroactiveLogRecalculate(); } catch(e) { Logger.log("Reference Matrix Recalculation Alert: " + e.toString()); }[cite: 1]
  try {
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);[cite: 1]
    if (sheet) {
      manageOperatorStation({ source: ss, range: sheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) });[cite: 1]
    }
  } catch(e) { Logger.log("Console screen alert: " + e.toString()); }[cite: 1]
}

/**
 * Helper to safely extract positive absolute numbers.
 */
function safeAbsNum(val) {
  if (val === "" || val === null || val === undefined) return "";[cite: 1]
  var n = parseFloat(val);[cite: 1]
  return isNaN(n) ? val : Math.abs(n);[cite: 1]
}

/**
 * Normalizes string keys by stripping non-alphanumeric characters and lowercasing.
 */
function cleanKey(val) {
  if (val === null || val === undefined) return "";[cite: 1]
  return String(val).toLowerCase().replace(/[^a-z0-9]/g, "");[cite: 1]
}

/**
 * Safely fetches column value from a log row using CONFIG key or 0-indexed fallback.
 */
function getLogVal(row, colConfigProp, defaultIndex) {
  if (colConfigProp && !isNaN(colConfigProp) && colConfigProp > 0) {[cite: 1]
    return row[colConfigProp - 1];[cite: 1]
  }
  return row[defaultIndex];[cite: 1]
}

/**
 * Sets cell A8 text and dynamically formats background & text color based on status.
 */
function setA8Status(sheet, statusMessage) {
  var ranges = CONFIG.OPERATOR_STATION.RANGES;[cite: 1]
  var a8 = sheet.getRange(ranges.STATUS_BANNER);[cite: 1]
  a8.setValue(statusMessage);[cite: 1]

  var upper = String(statusMessage).toUpperCase();

  if (upper.includes("COMPLETED AND PASSING")) {[cite: 1]
    a8.setBackground("#00C853").setFontColor("#FFFFFF").setFontWeight("bold");[cite: 1]
  } else if (upper.includes("FAIL") || upper.includes("ACTION REQUIRED") || upper.includes("NOT FOUND")) {[cite: 1]
    a8.setBackground("#D50000").setFontColor("#FFFFFF").setFontWeight("bold");[cite: 1]
  } else if (upper.includes("CONDITIONAL") || upper.includes("ATTENTION") || upper.includes("PENDING") || upper.includes("IN PROGRESS")) {[cite: 1]
    a8.setBackground("#FFD600").setFontColor("#000000").setFontWeight("bold");[cite: 1]
  } else {
    a8.setBackground(null).setFontColor(null).setFontWeight("normal");[cite: 1]
  }
}

/**
 * Operator Station Event Manager for Barcode Scanning & Work Order Lookup.
 */
function manageOperatorStation(e) {
  var ss = e ? e.source : SpreadsheetApp.getActiveSpreadsheet(); [cite: 1]
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION); [cite: 1]
  if (!sheet) return;
  
  var ranges = CONFIG.OPERATOR_STATION.RANGES;[cite: 1]
  var range = e ? e.range : sheet.getRange(ranges.BARCODE_INPUT); [cite: 1]
  var barcode = String(sheet.getRange(ranges.BARCODE_INPUT).getValue()).trim();  [cite: 1]
    
  if (range.getA1Notation() === ranges.BARCODE_INPUT) {  [cite: 1]
    sheet.getRange(ranges.CLEAR_METADATA_RANGE).clearContent(); [cite: 1]
    sheet.getRange(ranges.CLEAR_PANEL_RANGE).clearContent(); [cite: 1]
    sheet.getRange(ranges.CLEAR_LIMITS_RANGE).clearContent();[cite: 1]
    sheet.getRange(ranges.STATUS_BANNER).clearContent().setBackground(null).setFontColor(null);[cite: 1]
    
    var tableRange = sheet.getRange(ranges.CLEAR_RESULTS_RANGE);   [cite: 1]
    tableRange.clearContent();   [cite: 1]
    tableRange.setBackground("#FFFFFF").setFontColor("#000000").setFontWeight("normal");   [cite: 1]

    if (!barcode || barcode === "undefined" || barcode === "null") return;  [cite: 1]
      
    var searchBarcode = barcode;  [cite: 1]
    if (barcode.includes("-")) { searchBarcode = barcode.split("-")[0].trim(); }   [cite: 1]
    else if (barcode.includes("_")) { searchBarcode = barcode.split("_")[0].trim(); }  [cite: 1]
    if (!isNaN(searchBarcode) && searchBarcode.length === 4) { searchBarcode = "00" + searchBarcode; }  [cite: 1]
      
    var searchCriteria = "title contains '" + searchBarcode + "' and (mimeType = 'application/vnd.google-apps.spreadsheet' or mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') and trashed = false";   [cite: 1]
    var files = DriveApp.searchFiles(searchCriteria);  [cite: 1]
    if (!files.hasNext()) { [cite: 1]
      sheet.getRange(ranges.FILE_LINK_OUTPUT).setValue("❌ Work Order File Not Found: " + searchBarcode); [cite: 1]
      sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("CROSS-CHECK FAILED");[cite: 1]
      setA8Status(sheet, "WORK ORDER FILE NOT FOUND");[cite: 1]
      return; [cite: 1]
    }  
    
    var file = files.next();  [cite: 1]
    var verifiedFileIdStr = file.getId();  [cite: 1]
    var realFileName = file.getName();   [cite: 1]
    
    sheet.getRange(ranges.FILE_LINK_OUTPUT).setFormula('=HYPERLINK("' + file.getUrl() + '", "🔗 Open ' + realFileName + '")');   [cite: 1]
    sheet.getRange(ranges.CACHED_FILE_ID).setValue(verifiedFileIdStr);  [cite: 1]
      
    try {  [cite: 1]
      var woSpreadsheet = SpreadsheetApp.openById(verifiedFileIdStr); [cite: 1]
      var woSheet = woSpreadsheet.getSheets()[0];   [cite: 1]
      var woPartNumber = String(woSheet.getRange("D3").getValue()).trim(); [cite: 1]
      var woBomRevision = String(woSheet.getRange("D4").getValue()).trim();   [cite: 1]
      
      sheet.getRange(ranges.BOM_REV_OUTPUT).setValue(woBomRevision); [cite: 1]
      sheet.getRange(ranges.BASE_MODEL_OUTPUT).setValue(woPartNumber);   [cite: 1]

      var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);  [cite: 1]
      var matchedProgramName = "";
      var matchedDynamicKey = "";

      if (registrySheet && woPartNumber) {  [cite: 1]
        var regValues = registrySheet.getDataRange().getValues();  [cite: 1]
        var cleanWoPart = cleanKey(woPartNumber);  [cite: 1]
        var regCols = CONFIG.COLUMNS.PROGRAM_REGISTRY;[cite: 1]
        
        for (var rR = 1; rR < regValues.length; rR++) {  [cite: 1]
          var regRow = regValues[rR];[cite: 1]
          var regPartClean = cleanKey(regRow[regCols.BASE_MODEL - 1]);  [cite: 1]
          var regProgName = String(regRow[regCols.PROGRAM_NAME - 1] || "").trim();  [cite: 1]
          var regDynamicKey = String(regRow[regCols.DYNAMIC_KEY - 1] || "").trim();[cite: 1]
          
          if (regPartClean === cleanWoPart && regProgName) {  [cite: 1]
            matchedProgramName = regProgName;
            matchedDynamicKey = regDynamicKey;
            
            sheet.getRange(ranges.CUSTOMER_ACCOUNT_OUTPUT).setValue(regRow[regCols.CUSTOMER_ACCOUNT - 1] || "");[cite: 1]
            sheet.getRange(ranges.VEHICLE_SPEC_OUTPUT).setValue(regRow[regCols.VEHICLE_SPEC - 1] || "");[cite: 1]
            sheet.getRange(ranges.PROGRAM_NAME_OUTPUT).setValue(regProgName);[cite: 1]
            sheet.getRange(ranges.VALVING_VERSION_OUTPUT).setValue(regRow[regCols.VALVING_VERSION - 1] || "");[cite: 1]
            sheet.getRange(ranges.ADJUSTER_TARGETS_OUTPUT).setValue(regRow[regCols.ADJUSTER_SETTINGS - 1] || "");[cite: 1]
            
            sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("Validated in Registry");[cite: 1]
            break;
          }  
        }  
      }  

      if (!matchedProgramName) {
        sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("Registry Match Pending");[cite: 1]
      }

      populateSpecLimits(ss, sheet, matchedDynamicKey || matchedProgramName || woPartNumber);[cite: 1]
      renderOperatorTableWithFormatting(ss, sheet, searchBarcode, woPartNumber);[cite: 1]

    } catch(e) {
      Logger.log("WO Lookup Error: " + e.toString());[cite: 1]
      sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("Extraction Error");[cite: 1]
    }
  }  
}  

/**
 * Direct lookup helper for spec limits from Part_Reference_Matrix.
 */
function getSpecLimitsFromMatrix(ss, dynamicKeyOrPart) {
  var limits = { c1Min: NaN, c1Max: NaN, r1Min: NaN, r1Max: NaN, c2Min: NaN, c2Max: NaN, r2Min: NaN, r2Max: NaN, slopeMin: NaN };[cite: 1]
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);[cite: 1]
  if (!refSheet || !dynamicKeyOrPart) return limits;[cite: 1]

  var refData = refSheet.getDataRange().getValues();[cite: 1]
  var refCols = CONFIG.COLUMNS.PART_REFERENCE_MATRIX;[cite: 1]
  var targetCleanKey = cleanKey(dynamicKeyOrPart);[cite: 1]

  var refRow = null;
  for (var i = 1; i < refData.length; i++) {[cite: 1]
    var matrixKey = cleanKey(refData[i][refCols.DYNAMIC_KEY - 1]);[cite: 1]
    if (matrixKey === targetCleanKey || matrixKey.includes(targetCleanKey) || targetCleanKey.includes(matrixKey)) {[cite: 1]
      refRow = refData[i];
      break;
    }
  }

  if (!refRow) return limits;[cite: 1]

  function getAbsPair(minVal, maxVal) {
    if (minVal === "" || maxVal === "" || minVal === null || maxVal === null) return { min: NaN, max: NaN };[cite: 1]
    var a = Math.abs(parseFloat(minVal));[cite: 1]
    var b = Math.abs(parseFloat(maxVal));[cite: 1]
    if (isNaN(a) || isNaN(b)) return { min: NaN, max: NaN };[cite: 1]
    return { min: Math.min(a, b), max: Math.max(a, b) };[cite: 1]
  }

  var c1 = getAbsPair(refRow[refCols.COMP_1_MIN - 1], refRow[refCols.COMP_1_MAX - 1]);[cite: 1]
  var r1 = getAbsPair(refRow[refCols.REB_1_MIN - 1], refRow[refCols.REB_1_MAX - 1]);[cite: 1]
  var c2 = getAbsPair(refRow[refCols.COMP_2_MIN - 1], refRow[refCols.COMP_2_MAX - 1]);[cite: 1]
  var r2 = getAbsPair(refRow[refCols.REB_2_MIN - 1], refRow[refCols.REB_2_MAX - 1]);[cite: 1]
  var slope = parseFloat(refRow[refCols.SLOPE_1_MIN - 1]);[cite: 1]

  return {
    c1Min: c1.min, c1Max: c1.max,[cite: 1]
    r1Min: r1.min, r1Max: r1.max,[cite: 1]
    c2Min: c2.min, c2Max: c2.max,[cite: 1]
    r2Min: r2.min, r2Max: r2.max,[cite: 1]
    slopeMin: isNaN(slope) ? NaN : Math.abs(slope)[cite: 1]
  };
}

/**
 * Populates spec limit cells (B22:F23) with positive absolute values.
 */
function populateSpecLimits(ss, sheet, dynamicKeyOrPart) {
  var limits = getSpecLimitsFromMatrix(ss, dynamicKeyOrPart);[cite: 1]
  var ranges = CONFIG.OPERATOR_STATION.RANGES;[cite: 1]

  sheet.getRange(ranges.LIMIT_COMP_1_MIN).setValue(isNaN(limits.c1Min) ? "" : limits.c1Min);[cite: 1]
  sheet.getRange(ranges.LIMIT_COMP_1_MAX).setValue(isNaN(limits.c1Max) ? "" : limits.c1Max);[cite: 1]
  sheet.getRange(ranges.LIMIT_REB_1_MIN).setValue(isNaN(limits.r1Min) ? "" : limits.r1Min);[cite: 1]
  sheet.getRange(ranges.LIMIT_REB_1_MAX).setValue(isNaN(limits.r1Max) ? "" : limits.r1Max);[cite: 1]

  sheet.getRange(ranges.LIMIT_COMP_2_MIN).setValue(isNaN(limits.c2Min) ? "" : limits.c2Min);[cite: 1]
  sheet.getRange(ranges.LIMIT_COMP_2_MAX).setValue(isNaN(limits.c2Max) ? "" : limits.c2Max);[cite: 1]
  sheet.getRange(ranges.LIMIT_REB_2_MIN).setValue(isNaN(limits.r2Min) ? "" : limits.r2Min);[cite: 1]
  sheet.getRange(ranges.LIMIT_REB_2_MAX).setValue(isNaN(limits.r2Max) ? "" : limits.r2Max);[cite: 1]

  sheet.getRange(ranges.LIMIT_SLOPE_1_MIN).setValue(isNaN(limits.slopeMin) ? "No Limit" : limits.slopeMin.toFixed(1));[cite: 1]
}

/**
 * Queries Master_Dyno_Log and renders records in the exact 12-column UI table layout (A26:L100).
 */
function renderOperatorTableWithFormatting(ss, sheet, searchBarcode, partNumber) {
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);[cite: 1]
  if (!logSheet) return;

  var logData = logSheet.getDataRange().getValues();[cite: 1]
  if (logData.length <= 1) {
    setA8Status(sheet, "PENDING TESTING");[cite: 1]
    return;
  }

  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG || {};[cite: 1]
  var ranges = CONFIG.OPERATOR_STATION.RANGES;[cite: 1]
  var logSheetId = logSheet.getSheetId();[cite: 1]

  var limits = {
    c1Min: parseFloat(sheet.getRange(ranges.LIMIT_COMP_1_MIN).getValue()),[cite: 1]
    c1Max: parseFloat(sheet.getRange(ranges.LIMIT_COMP_1_MAX).getValue()),[cite: 1]
    r1Min: parseFloat(sheet.getRange(ranges.LIMIT_REB_1_MIN).getValue()),[cite: 1]
    r1Max: parseFloat(sheet.getRange(ranges.LIMIT_REB_1_MAX).getValue()),[cite: 1]
    c2Min: parseFloat(sheet.getRange(ranges.LIMIT_COMP_2_MIN).getValue()),[cite: 1]
    c2Max: parseFloat(sheet.getRange(ranges.LIMIT_COMP_2_MAX).getValue()),[cite: 1]
    r2Min: parseFloat(sheet.getRange(ranges.LIMIT_REB_2_MIN).getValue()),[cite: 1]
    r2Max: parseFloat(sheet.getRange(ranges.LIMIT_REB_2_MAX).getValue())[cite: 1]
  };

  if (isNaN(limits.c1Min) || isNaN(limits.c1Max)) {[cite: 1]
    var directLimits = getSpecLimitsFromMatrix(ss, partNumber || searchBarcode);[cite: 1]
    if (!isNaN(directLimits.c1Min)) limits = directLimits;[cite: 1]
  }

  var cleanBarcodeStr = cleanKey(searchBarcode);[cite: 1]
  var cleanPartStr = cleanKey(partNumber);[cite: 1]

  var latestLogBySerial = {};

  for (var r = 1; r < logData.length; r++) {[cite: 1]
    var row = logData[r];[cite: 1]
    var trueSerial = String(getLogVal(row, logCols.TRUE_SERIAL, 2) || "").trim();[cite: 1]
    var baseModel = cleanKey(getLogVal(row, logCols.BASE_MODEL, 3));[cite: 1]
    var cleanSerial = cleanKey(trueSerial);[cite: 1]

    var isMatch = false;
    if (cleanBarcodeStr !== "" && cleanSerial.includes(cleanBarcodeStr)) {[cite: 1]
      isMatch = true;
    } else if (cleanPartStr !== "" && baseModel === cleanPartStr && (cleanBarcodeStr === "" || cleanBarcodeStr === "undefined")) {[cite: 1]
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

  var rowsToDisplay = [];
  var test1FailCount = 0;
  var test2FailCount = 0;

  var serialKeys = Object.keys(latestLogBySerial).sort(function(a, b) {
    var mA = a.match(/(\d+)$/);
    var mB = b.match(/(\d+)$/);
    var uA = mA ? parseInt(mA[1], 10) : 0;
    var uB = mB ? parseInt(mB[1], 10) : 0;
    return uA - uB;
  });

  for (var k = 0; k < serialKeys.length; k++) {
    var item = latestLogBySerial[serialKeys[k]];
    var row = item.data;
    var actualSheetRow = item.rowIdx;
    var trueSerial = item.trueSerial;

    var rowLink = "#gid=" + logSheetId + "&range=A" + actualSheetRow;[cite: 1]
    var serialHyperlinkFormula = '=HYPERLINK("' + rowLink + '", "' + trueSerial + '")';[cite: 1]

    var t1Status    = String(getLogVal(row, logCols.TEST_1_STATUS, 18) || "").trim();[cite: 1]
    var t2Status    = String(getLogVal(row, logCols.TEST_2_STATUS, 19) || "").trim();[cite: 1]
    var overallStat = String(getLogVal(row, logCols.OVERALL_STATUS, 20) || "").trim();[cite: 1]
    var diagnostics = String(getLogVal(row, logCols.DIAGNOSTICS, 21) || "").trim();[cite: 1]
    var evalAction  = String(getLogVal(row, logCols.EVALUATION_ACTION, 22) || "").trim();[cite: 1]
    var engComm     = String(getLogVal(row, logCols.ENG_COMMENTS, 23) || "").trim();[cite: 1]

    if (t1Status.toUpperCase().includes("FAIL")) test1FailCount++;
    if (t2Status.toUpperCase().includes("FAIL")) test2FailCount++;

    var mappedRow = [
      serialHyperlinkFormula,                           // Col A (1)[cite: 1]
      safeAbsNum(getLogVal(row, logCols.ROD_FORCE, 5)), // Col B (2)[cite: 1]
      safeAbsNum(getLogVal(row, logCols.COMP_1, 7)),    // Col C (3)[cite: 1]
      safeAbsNum(getLogVal(row, logCols.REB_1, 8)),     // Col D (4)[cite: 1]
      safeAbsNum(getLogVal(row, logCols.COMP_2, 12)),   // Col E (5)[cite: 1]
      safeAbsNum(getLogVal(row, logCols.REB_2, 13)),    // Col F (6)[cite: 1]
      t1Status,                                         // Col G (7)[cite: 1]
      t2Status,                                         // Col H (8)[cite: 1]
      overallStat,                                      // Col I (9)[cite: 1]
      evalAction,                                       // Col J (10)[cite: 1]
      diagnostics,                                      // Col K (11)[cite: 1]
      engComm                                           // Col L (12)[cite: 1]
    ];

    rowsToDisplay.push(mappedRow);
  }

  var statusMessage = "";
  if (rowsToDisplay.length === 0) {
    statusMessage = "PENDING TESTING";[cite: 1]
  } else if (test1FailCount > 0) {
    statusMessage = "ACTION REQUIRED: Test 1 Failure Detected (" + test1FailCount + " unit(s))";[cite: 1]
  } else if (test2FailCount > 0) {
    statusMessage = "CONDITIONAL PASS: Attention Required (Test 2 Outlier Detected - " + test2FailCount + " unit(s))";[cite: 1]
  } else {
    statusMessage = "WORK ORDER COMPLETED AND PASSING";[cite: 1]
  }

  setA8Status(sheet, statusMessage);[cite: 1]

  if (rowsToDisplay.length === 0) return;

  var startRow = ranges.RESULTS_START_ROW;[cite: 1]
  var numRows = rowsToDisplay.length;
  var numCols = ranges.RESULTS_COL_COUNT;[cite: 1]
  var outputRange = sheet.getRange(startRow, ranges.RESULTS_START_COL, numRows, numCols);[cite: 1]

  outputRange.setValues(rowsToDisplay);[cite: 1]
  sheet.getRange(startRow, 2, numRows, 5).setNumberFormat("0.0");[cite: 1]

  var bgColors = [];
  var fontColors = [];
  var fontWeights = [];

  for (var rIdx = 0; rIdx < numRows; rIdx++) {
    var rowBg = ["#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF"];[cite: 1]
    var rowFont = ["#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000"];[cite: 1]
    var rowWeight = ["normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal"];[cite: 1]
    var rowData = rowsToDisplay[rIdx];

    var t1StatusStr    = String(rowData[6] || "").toUpperCase();[cite: 1]
    var t2StatusStr    = String(rowData[7] || "").toUpperCase();[cite: 1]
    var overallStatStr = String(rowData[8] || "").toUpperCase();[cite: 1]
    var diagnosticsStr = String(rowData[10] || "");[cite: 1]

    var applyFaultHighlight = function(colIndex) {
      rowFont[colIndex] = "#FF0000";      [cite: 1]
      rowWeight[colIndex] = "bold";       [cite: 1]
      rowBg[colIndex] = "#FADBD8";        [cite: 1]
    };

    // Always highlight failing cells regardless of override status
    if (!diagnosticsStr.includes("✅")) {[cite: 1]
      if (diagnosticsStr.indexOf("[RF_FAIL]") !== -1)    applyFaultHighlight(1);[cite: 1]
      if (diagnosticsStr.indexOf("[C1_FAIL]") !== -1)    applyFaultHighlight(2);[cite: 1]
      if (diagnosticsStr.indexOf("[R1_FAIL]") !== -1)    applyFaultHighlight(3);[cite: 1]
      if (diagnosticsStr.indexOf("[C2_FAIL]") !== -1)    applyFaultHighlight(4);[cite: 1]
      if (diagnosticsStr.indexOf("[R2_FAIL]") !== -1)    applyFaultHighlight(5);[cite: 1]
      if (diagnosticsStr.indexOf("[SLOPE_FAIL]") !== -1) { applyFaultHighlight(2); applyFaultHighlight(3); }[cite: 1]
    }

    var c1Val = parseFloat(rowData[2]);[cite: 1]
    var r1Val = parseFloat(rowData[3]);[cite: 1]
    var c2Val = parseFloat(rowData[4]);[cite: 1]
    var r2Val = parseFloat(rowData[5]);[cite: 1]

    if (!isNaN(c1Val) && ((!isNaN(limits.c1Min) && c1Val < limits.c1Min) || (!isNaN(limits.c1Max) && c1Val > limits.c1Max))) applyFaultHighlight(2);[cite: 1]
    if (!isNaN(r1Val) && ((!isNaN(limits.r1Min) && r1Val < limits.r1Min) || (!isNaN(limits.r1Max) && r1Val > limits.r1Max))) applyFaultHighlight(3);[cite: 1]
    if (!isNaN(c2Val) && ((!isNaN(limits.c2Min) && c2Val < limits.c2Min) || (!isNaN(limits.c2Max) && c2Val > limits.c2Max))) applyFaultHighlight(4);[cite: 1]
    if (!isNaN(r2Val) && ((!isNaN(limits.r2Min) && r2Val < limits.r2Min) || (!isNaN(limits.r2Max) && r2Val > limits.r2Max))) applyFaultHighlight(5);[cite: 1]

    if (t1StatusStr.includes("FAIL")) {[cite: 1]
      rowBg[6] = "#FADBD8"; rowFont[6] = "#C0392B"; rowWeight[6] = "bold";[cite: 1]
    } else if (t1StatusStr.includes("PASS")) {[cite: 1]
      rowBg[6] = "#D4EFDF"; rowFont[6] = "#196F3D";[cite: 1]
    }

    if (t2StatusStr.includes("FAIL")) {[cite: 1]
      rowBg[7] = "#FADBD8"; rowFont[7] = "#C0392B"; rowWeight[7] = "bold";[cite: 1]
    } else if (t2StatusStr.includes("PASS")) {[cite: 1]
      rowBg[7] = "#D4EFDF"; rowFont[7] = "#196F3D";[cite: 1]
    }

    if (overallStatStr.includes("FAIL")) {[cite: 1]
      rowBg[8] = "#C0392B"; rowFont[8] = "#FFFFFF"; rowWeight[8] = "bold";[cite: 1]
    } else if (overallStatStr.includes("OVERRIDE") || overallStatStr.includes("PASS")) {[cite: 1]
      rowBg[8] = "#D4EFDF"; rowFont[8] = "#196F3D"; rowWeight[8] = "bold";[cite: 1]
    }

    bgColors.push(rowBg);[cite: 1]
    fontColors.push(rowFont);[cite: 1]
    fontWeights.push(rowWeight);[cite: 1]
  }

  outputRange.setBackgrounds(bgColors).setFontColors(fontColors).setFontWeights(fontWeights);[cite: 1]
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

  // 1. Operator Station Edits
  if (sheetName === CONFIG.SHEET_NAMES.OPERATOR_STATION) {[cite: 1]
    try { manageOperatorStation(e); } catch(err) { Logger.log("Operator station edit error: " + err.toString()); }
    return;
  }

  // 2. Master Dyno Log Edits (Columns V through Z / Cols 22 through 26)
  if (sheetName === CONFIG.SHEET_NAMES.MASTER_DYNO_LOG) {[cite: 1]
    var editedCol = e.range.getColumn();

    if (editedCol >= 22 && editedCol <= 26) {
      retroactiveLogRecalculate();

      var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();[cite: 1]
      var opSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);[cite: 1]
      if (opSheet) {
        try {
          manageOperatorStation({
            source: ss,
            range: opSheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT)[cite: 1]
          });
        } catch(err) {
          Logger.log("Notice: UI refresh skipped (Installable Trigger required for DriveApp lookup).");
        }
      }
    }
  }
}

/**
 * Automated Trigger Setup Helper.
 */
function setupTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();[cite: 1]
  
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