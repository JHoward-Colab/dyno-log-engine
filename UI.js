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
 * Normalizes string keys by stripping all non-alphanumeric characters and lowercasing.
 */
function cleanKey(val) {
  if (val === null || val === undefined) return "";
  return String(val).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Returns sorted positive absolute min and max pair from matrix values.
 */
function getAbsMinMax(val1, val2) {
  var num1 = Math.abs(parseFloat(val1));
  var num2 = Math.abs(parseFloat(val2));
  if (isNaN(num1) && isNaN(num2)) return { min: NaN, max: NaN };
  if (isNaN(num1)) return { min: num2, max: num2 };
  if (isNaN(num2)) return { min: num1, max: num1 };
  return { min: Math.min(num1, num2), max: Math.max(num1, num2) };
}

/**
 * Sets cell A8 text and dynamically formats background & text color based on the status.
 */
function setA8Status(sheet, statusMessage) {
  var a8 = sheet.getRange("A8");
  a8.setValue(statusMessage);

  var upper = String(statusMessage).toUpperCase();

  if (upper.includes("COMPLETED AND PASSING")) {
    a8.setBackground("#00C853").setFontColor("#FFFFFF").setFontWeight("bold"); // Bright Green
  } else if (upper.includes("FAIL") || upper.includes("ACTION REQUIRED") || upper.includes("NOT FOUND")) {
    a8.setBackground("#D50000").setFontColor("#FFFFFF").setFontWeight("bold"); // Bright Red
  } else if (upper.includes("CONDITIONAL") || upper.includes("ATTENTION") || upper.includes("PENDING") || upper.includes("IN PROGRESS")) {
    a8.setBackground("#FFD600").setFontColor("#000000").setFontWeight("bold"); // Bright Yellow
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
    // 1. Clear Value Cells & Formatting Only (Preserving static label cells A14:A18)
    sheet.getRange(ranges.CLEAR_METADATA_RANGE).clearContent(); 
    sheet.getRange("C6").clearContent(); 
    sheet.getRange("A8").clearContent().setBackground(null).setFontColor(null);
    sheet.getRange("C14:C18").clearContent();  
    sheet.getRange("E14:E18").clearContent();  
    sheet.getRange("B22:F23").clearContent();
    
    // 2. Clear Results Table & Formatting (A27:L100)
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
      sheet.getRange("C6").setValue("CROSS-CHECK FAILED");
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
      sheet.getRange(ranges.PART_NO_OUTPUT).setValue(woPartNumber);   

      // Cell C6 Cross-Check Verification
      sheet.getRange("C6").setValue(searchBarcode);
      
      var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);  
      var matchedProgramName = "";

      if (registrySheet && woPartNumber) {  
        var regValues = registrySheet.getDataRange().getValues();  
        var cleanWoPart = cleanKey(woPartNumber);  
        
        var regProgIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.PROGRAM_NAME - 1; 
        var regBaseModelIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.BASE_MODEL - 1; 
        
        for (var rR = 1; rR < regValues.length; rR++) {  
          var regRow = regValues[rR];
          var regPartClean = cleanKey(regRow[regBaseModelIdx]);  
          var regProgName = String(regRow[regProgIdx] || "").trim();  
          
          if (regPartClean === cleanWoPart && regProgName) {  
            matchedProgramName = regProgName;
            
            // EXACT METADATA PANEL POPULATION (C14:C18)
            sheet.getRange("C14").setValue(regRow[3] || ""); // Col D: Customer Account
            sheet.getRange("C15").setValue(regRow[4] || ""); // Col E: Vehicle/Model Spec
            sheet.getRange("C16").setValue(regRow[0] || ""); // Col A: LABA7 Program Name
            sheet.getRange("C17").setValue(regRow[6] || ""); // Col G: Target Valving Spec
            sheet.getRange("C18").setValue(regRow[7] || ""); // Col H: Required Clicker Targets
            break;
          }  
        }  
      }  

      // 3. Populate Limit Cells B22:F23 with absolute positive values
      populateSpecLimits(ss, sheet, matchedProgramName || woPartNumber);

      // 4. Render dyno log records directly by matching scanned barcode / part number
      renderOperatorTableWithFormatting(ss, sheet, searchBarcode, woPartNumber);

    } catch(e) {
      Logger.log("WO Lookup Error: " + e.toString());
    }
  }  
}  

/**
 * Direct lookup helper for spec limits from Part_Reference_Matrix using correct column indices.
 */
function getSpecLimitsFromMatrix(ss, programOrPart) {
  var limits = { c1Min: NaN, c1Max: NaN, r1Min: NaN, r1Max: NaN, c2Min: NaN, c2Max: NaN, r2Min: NaN, r2Max: NaN };
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);
  if (!refSheet || !programOrPart) return limits;

  var refData = refSheet.getDataRange().getValues();
  var refCols = CONFIG.COLUMNS.PART_REFERENCE_MATRIX;
  var targetCleanKey = cleanKey(programOrPart);

  var refRow = null;
  for (var i = 1; i < refData.length; i++) {
    var progName = cleanKey(refData[i][refCols.PROGRAM_NAME - 1]);
    if (progName === targetCleanKey || progName.includes(targetCleanKey) || targetCleanKey.includes(progName)) {
      refRow = refData[i];
      break;
    }
  }

  if (!refRow) return limits;

  var c1 = getAbsMinMax(refRow[refCols.COMP_1_MIN - 1], refRow[refCols.COMP_1_MAX - 1]);
  var r1 = getAbsMinMax(refRow[refCols.REB_1_MIN - 1], refRow[refCols.REB_1_MAX - 1]);
  var c2 = getAbsMinMax(refRow[refCols.COMP_2_MIN - 1], refRow[refCols.COMP_2_MAX - 1]);
  var r2 = getAbsMinMax(refRow[refCols.REB_2_MIN - 1], refRow[refCols.REB_2_MAX - 1]);

  return {
    c1Min: c1.min, c1Max: c1.max,
    r1Min: r1.min, r1Max: r1.max,
    c2Min: c2.min, c2Max: c2.max,
    r2Min: r2.min, r2Max: r2.max
  };
}

/**
 * Populates spec limit cells (B22:F23) with positive absolute values.
 */
function populateSpecLimits(ss, sheet, programOrPart) {
  var limits = getSpecLimitsFromMatrix(ss, programOrPart);
  var ranges = CONFIG.OPERATOR_STATION.RANGES;

  sheet.getRange(ranges.LIMIT_C1_MIN).setValue(isNaN(limits.c1Min) ? "" : limits.c1Min);
  sheet.getRange(ranges.LIMIT_C1_MAX).setValue(isNaN(limits.c1Max) ? "" : limits.c1Max);
  sheet.getRange(ranges.LIMIT_R1_MIN).setValue(isNaN(limits.r1Min) ? "" : limits.r1Min);
  sheet.getRange(ranges.LIMIT_R1_MAX).setValue(isNaN(limits.r1Max) ? "" : limits.r1Max);

  sheet.getRange(ranges.LIMIT_C2_MIN).setValue(isNaN(limits.c2Min) ? "" : limits.c2Min);
  sheet.getRange(ranges.LIMIT_C2_MAX).setValue(isNaN(limits.c2Max) ? "" : limits.c2Max);
  sheet.getRange(ranges.LIMIT_R2_MIN).setValue(isNaN(limits.r2Min) ? "" : limits.r2Min);
  sheet.getRange(ranges.LIMIT_R2_MAX).setValue(isNaN(limits.r2Max) ? "" : limits.r2Max);
}

/**
 * Queries Master_Dyno_Log directly for records matching the scanned barcode prefix.
 * Maps Col X (Overall Status) to Col I, Col Z (Evaluation Action) to Col J,
 * highlights specific out-of-spec force values in BOLD RED, and updates Cell A8.
 */
function renderOperatorTableWithFormatting(ss, sheet, searchBarcode, partNumber) {
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  if (!logSheet) return;

  var logData = logSheet.getDataRange().getValues();
  if (logData.length <= 1) {
    setA8Status(sheet, "PENDING TESTING");
    return;
  }

  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG;
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  var logSheetId = logSheet.getSheetId();

  // Read active spec limits directly from Matrix lookup
  var limits = getSpecLimitsFromMatrix(ss, partNumber || searchBarcode);

  var cleanBarcodeStr = cleanKey(searchBarcode);
  var cleanPartStr = cleanKey(partNumber);

  // 1. First Pass: Collect matching entries and keep only the LATEST entry per serial
  var latestLogBySerial = {};

  for (var r = 1; r < logData.length; r++) {
    var row = logData[r];
    var trueSerial = String(row[logCols.TRUE_SERIAL - 1] || "").trim();
    var baseModel = cleanKey(row[logCols.BASE_MODEL - 1]);
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

  // Natural numeric sort on unit suffixes (e.g. 1522-001, 1522-006, 1522-009, 1522-015)
  var serialKeys = Object.keys(latestLogBySerial).sort(function(a, b) {
    var mA = a.match(/(\d+)$/);
    var mB = b.match(/(\d+)$/);
    var uA = mA ? parseInt(mA[1], 10) : 0;
    var uB = mB ? parseInt(mB[1], 10) : 0;
    return uA - uB;
  });

  // 2. Second Pass: Build table rows and evaluate test outcomes
  for (var k = 0; k < serialKeys.length; k++) {
    var item = latestLogBySerial[serialKeys[k]];
    var row = item.data;
    var actualSheetRow = item.rowIdx;
    var trueSerial = item.trueSerial;

    var rowLink = "#gid=" + logSheetId + "&range=A" + actualSheetRow;
    var serialHyperlinkFormula = '=HYPERLINK("' + rowLink + '", "' + trueSerial + '")';

    var t1Status = String(row[logCols.TEST_1_STATUS - 1] || "").trim(); // Log Col V (idx 21)
    var t2Status = String(row[logCols.TEST_2_STATUS - 1] || "").trim(); // Log Col W (idx 22)

    var t1Upper = t1Status.toUpperCase();
    var t2Upper = t2Status.toUpperCase();

    if (t1Upper.includes("FAIL")) test1FailCount++;
    if (t2Upper.includes("FAIL")) test2FailCount++;

    // Explicit Log Column Bindings for UI Table
    var overallStat = String(row[logCols.OVERALL_STATUS - 1] || "").trim();    // Log Col X (idx 23) -> UI Col I
    var evalAction  = String(row[logCols.EVALUATION_ACTION - 1] || "").trim(); // Log Col Z (idx 25) -> UI Col J
    var teardownAct = String(row[logCols.TEARDOWN_ACTION - 1] || "").trim();   // Log Col Z (idx 25)

    var fullEvalAction = evalAction || teardownAct;

    // 12-COLUMN TABLE MAPPING
    var mappedRow = [
      serialHyperlinkFormula,                          // Col A (1): True Serial (Hyperlink)
      safeAbsNum(row[logCols.ROD_FORCE - 1]),          // Col B (2): Rod Force (ABS)
      safeAbsNum(row[logCols.COMP_1 - 1]),             // Col C (3): Low Speed Comp (ABS)
      safeAbsNum(row[logCols.REB_1 - 1]),              // Col D (4): Low Speed Reb (ABS)
      safeAbsNum(row[logCols.COMP_2 - 1]),             // Col E (5): Med Speed Comp (ABS)
      safeAbsNum(row[logCols.REB_2 - 1]),              // Col F (6): Med Speed Reb (ABS)
      t1Status,                                        // Col G (7): Test 1 Status (Log Col V)
      t2Status,                                        // Col H (8): Test 2 Status (Log Col W)
      overallStat,                                     // Col I (9): Overall Status (Log Col X)
      fullEvalAction,                                  // Col J (10): Evaluation Action (Log Col Z)
      safeAbsNum(row[logCols.COMP_3 - 1]),             // Col K (11): High Speed Comp (ABS)
      safeAbsNum(row[logCols.REB_3 - 1])              // Col L (12): High Speed Reb (ABS)
    ];

    rowsToDisplay.push(mappedRow);
  }

  // 3. Accurate Cell A8 Work Order Status Evaluation
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

  var startRow = ranges.RESULTS_START_ROW; // Row 27
  var numRows = rowsToDisplay.length;
  var numCols = ranges.RESULTS_COL_COUNT;   // 12 columns
  var outputRange = sheet.getRange(startRow, ranges.RESULTS_START_COL, numRows, numCols);

  // 4. Write Mapped Data Values & Formulas
  outputRange.setValues(rowsToDisplay);

  // 5. Set Decimal Formatting (X.X format on force columns B..F and K..L)
  sheet.getRange(startRow, 2, numRows, 5).setNumberFormat("0.0");  // Cols B..F
  sheet.getRange(startRow, 11, numRows, 2).setNumberFormat("0.0"); // Cols K..L

  // 6. Build Explicit Formatting Arrays (Failed forces & statuses -> BOLD Dark Red on Light Red)
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
        rowBg.push("#FFCCCC");   // Light Red Background
        rowFont.push("#990000"); // Dark Red Text
        rowWeight.push("bold");  // BOLD font for failing values
      } else {
        rowBg.push("#FFFFFF");   // Clean White Background
        rowFont.push("#000000"); // Black Text
        rowWeight.push("normal");
      }
    }
    bgColors.push(rowBg);
    fontColors.push(rowFont);
    fontWeights.push(rowWeight);
  }

  // 7. Apply explicit formatting matrices in a single API call
  outputRange.setBackgrounds(bgColors).setFontColors(fontColors).setFontWeights(fontWeights);
}