// =========================================================================
// 🖥️ USER INTERFACE & CONTROLLERS (UI.js)
// Workspace Rendering, Button Actions & Triggers
// =========================================================================

/**
 * Master Sync Action triggered by button click on Operator Station.
 */
function clickMasterSyncButton() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try { processDynoFiles(); } catch(e) { Logger.log("Watch folder alert: " + e.toString()); }
  try { retroactiveLogRecalculate(); } catch(e) { Logger.log("Reference Matrix Recalculation Alert: " + e.toString()); }
  try {
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION);
    if (sheet) {
      manageOperatorStation({ source: ss, range: sheet.getRange(CONFIG.OPERATOR_STATION.RANGES.BARCODE_INPUT) });
    }
  } catch(e) { Logger.log("Console screen alert: " + e.toString()); }
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
  } else if (upper.includes("CONDITIONAL") || upper.includes("ATTENTION") || upper.includes("PENDING") || upper.includes("IN PROGRESS")) {
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
    // 1. Clear Metadata, Panel, Limits, and Results
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
      
      sheet.getRange(ranges.BOM_REV_OUTPUT).setValue(woBomRevision); 
      sheet.getRange(ranges.BASE_MODEL_OUTPUT).setValue(woPartNumber);   

      var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);  
      var matchedProgramName = "";
      var matchedDynamicKey = "";

      if (registrySheet && woPartNumber) {  
        var regValues = registrySheet.getDataRange().getValues();  
        var cleanWoPart = cleanKey(woPartNumber);  
        
        var regCols = CONFIG.COLUMNS.PROGRAM_REGISTRY;
        
        for (var rR = 1; rR < regValues.length; rR++) {  
          var regRow = regValues[rR];
          var regPartClean = cleanKey(regRow[regCols.BASE_MODEL - 1]);  
          var regProgName = String(regRow[regCols.PROGRAM_NAME - 1] || "").trim();  
          var regDynamicKey = String(regRow[regCols.DYNAMIC_KEY - 1] || "").trim();
          
          if (regPartClean === cleanWoPart && regProgName) {  
            matchedProgramName = regProgName;
            matchedDynamicKey = regDynamicKey;
            
            // Populate Metadata Panel
            sheet.getRange(ranges.CUSTOMER_ACCOUNT_OUTPUT).setValue(regRow[regCols.CUSTOMER_ACCOUNT - 1] || "");
            sheet.getRange(ranges.VEHICLE_SPEC_OUTPUT).setValue(regRow[regCols.VEHICLE_SPEC - 1] || "");
            sheet.getRange(ranges.PROGRAM_NAME_OUTPUT).setValue(regProgName);
            sheet.getRange(ranges.VALVING_VERSION_OUTPUT).setValue(regRow[regCols.VALVING_VERSION - 1] || "");
            sheet.getRange(ranges.ADJUSTER_TARGETS_OUTPUT).setValue(regRow[regCols.ADJUSTER_SETTINGS - 1] || "");
            
            sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("Validated in Registry");
            break;
          }  
        }  
      }  

      if (!matchedProgramName) {
        sheet.getRange(ranges.CROSS_CHECK_OUTPUT).setValue("Registry Match Pending");
      }

      // Populate Spec Limits using Dynamic Key or Program Name
      populateSpecLimits(ss, sheet, matchedDynamicKey || matchedProgramName || woPartNumber);

      // Render Dyno Log Records Table
      renderOperatorTableWithFormatting(ss, sheet, searchBarcode, woPartNumber);

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
    if (isNaN(a) || isNaN(b)) return { min: NaN, max: NaN };
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
 * Queries Master_Dyno_Log and renders records in the exact 12-column UI table layout (A26:L100).
 */
function renderOperatorTableWithFormatting(ss, sheet, searchBarcode, partNumber) {
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  if (!logSheet) return;

  var logData = logSheet.getDataRange().getValues();
  if (logData.length <= 1) {
    setA8Status(sheet, "PENDING TESTING");
    return;
  }

  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG || {};
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  var logSheetId = logSheet.getSheetId();

  // Read active spec limits directly from cells B22:F23
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

    var rowLink = "#gid=" + logSheetId + "&range=A" + actualSheetRow;
    var serialHyperlinkFormula = '=HYPERLINK("' + rowLink + '", "' + trueSerial + '")';

    var t1Status    = String(getLogVal(row, logCols.TEST_1_STATUS, 18) || "").trim();     // Col S (19)
    var t2Status    = String(getLogVal(row, logCols.TEST_2_STATUS, 19) || "").trim();     // Col T (20)
    var overallStat = String(getLogVal(row, logCols.OVERALL_STATUS, 20) || "").trim();    // Col U (21)
    var diagnostics = String(getLogVal(row, logCols.DIAGNOSTICS, 21) || "").trim();        // Col V (22)
    var evalAction  = String(getLogVal(row, logCols.EVALUATION_ACTION, 22) || "").trim(); // Col W (23)
    var engComm     = String(getLogVal(row, logCols.ENG_COMMENTS, 23) || "").trim();       // Col X (24)

    if (t1Status.toUpperCase().includes("FAIL")) test1FailCount++;
    if (t2Status.toUpperCase().includes("FAIL")) test2FailCount++;

    // EXACT 12-COLUMN TABLE MAPPING (A26:L26 Headers)
    var mappedRow = [
      serialHyperlinkFormula,                           // Col A (1): Serial Number
      safeAbsNum(getLogVal(row, logCols.ROD_FORCE, 5)), // Col B (2): Rod Force
      safeAbsNum(getLogVal(row, logCols.COMP_1, 7)),    // Col C (3): Low Speed Comp
      safeAbsNum(getLogVal(row, logCols.REB_1, 8)),     // Col D (4): Low Speed Reb
      safeAbsNum(getLogVal(row, logCols.COMP_2, 12)),   // Col E (5): Med Speed Comp
      safeAbsNum(getLogVal(row, logCols.REB_2, 13)),    // Col F (6): Med Speed Reb
      t1Status,                                         // Col G (7): Test 1: Global Gate
      t2Status,                                         // Col H (8): Test 2: Batch Gate
      overallStat,                                      // Col I (9): Overall Status
      evalAction,                                       // Col J (10): Evaluation Action (Log Col W)
      diagnostics,                                      // Col K (11): Diagnostics and Troubleshooting (Log Col V)
      engComm                                           // Col L (12): Diagnostic Notes / Engineering Comments (Log Col X)
    ];

    rowsToDisplay.push(mappedRow);
  }

  var statusMessage = "";
  if (rowsToDisplay.length === 0) {
    statusMessage = "PENDING TESTING";
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

  // Set decimal format on numeric force columns B through F
  sheet.getRange(startRow, 2, numRows, 5).setNumberFormat("0.0"); // Cols B..F

  var bgColors = [];
  var fontColors = [];
  var fontWeights = [];

  for (var rIdx = 0; rIdx < numRows; rIdx++) {
    var rowBg = [];
    var rowFont = [];
    var rowWeight = [];
    var rowData = rowsToDisplay[rIdx];

    for (var cIdx = 0; cIdx < numCols; cIdx++) {
      var val = parseFloat(rowData[cIdx]);
      var strVal = String(rowData[cIdx] || "").toUpperCase();
      var isOut = false;

      // Evaluate numeric force columns against active spec limits
      if (!isNaN(val)) {
        if (cIdx === 2 && ((!isNaN(limits.c1Min) && val < limits.c1Min) || (!isNaN(limits.c1Max) && val > limits.c1Max))) isOut = true; // Low Comp (Col C)
        if (cIdx === 3 && ((!isNaN(limits.r1Min) && val < limits.r1Min) || (!isNaN(limits.r1Max) && val > limits.r1Max))) isOut = true; // Low Reb (Col D)
        if (cIdx === 4 && ((!isNaN(limits.c2Min) && val < limits.c2Min) || (!isNaN(limits.c2Max) && val > limits.c2Max))) isOut = true; // Med Comp (Col E)
        if (cIdx === 5 && ((!isNaN(limits.r2Min) && val < limits.r2Min) || (!isNaN(limits.r2Max) && val > limits.r2Max))) isOut = true; // Med Reb (Col F)
      }

      // Highlight status columns (Cols G, H, I) if they contain FAIL
      if ((cIdx === 6 || cIdx === 7 || cIdx === 8) && strVal.includes("FAIL")) {
        isOut = true;
      }

      if (isOut) {
        rowBg.push("#FFCCCC");
        rowFont.push("#990000");
        rowWeight.push("bold");
      } else {
        rowBg.push("#FFFFFF");
        rowFont.push("#000000");
        rowWeight.push("normal");
      }
    }
    bgColors.push(rowBg);
    fontColors.push(rowFont);
    fontWeights.push(rowWeight);
  }

  outputRange.setBackgrounds(bgColors).setFontColors(fontColors).setFontWeights(fontWeights);
}