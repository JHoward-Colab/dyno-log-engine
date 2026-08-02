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
 * Sets cell A8 text and dynamically formats background & text color based on the status.
 */
function setA8Status(sheet, statusMessage) {
  var a8 = sheet.getRange("A8");
  a8.setValue(statusMessage);

  var upper = String(statusMessage).toUpperCase();

  if (upper.includes("COMPLETED AND PASSING")) {
    // Bright Green (#00C853) / White Text
    a8.setBackground("#00C853").setFontColor("#FFFFFF").setFontWeight("bold");
  } else if (upper.includes("FAIL") || upper.includes("ACTION REQUIRED") || upper.includes("NOT FOUND")) {
    // Bright Red (#D50000) / White Text
    a8.setBackground("#D50000").setFontColor("#FFFFFF").setFontWeight("bold");
  } else if (upper.includes("CONDITIONAL") || upper.includes("ATTENTION") || upper.includes("PENDING") || upper.includes("IN PROGRESS")) {
    // Bright Yellow (#FFD600) / Black Text
    a8.setBackground("#FFD600").setFontColor("#000000").setFontWeight("bold");
  } else {
    // Neutral fallback reset
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
    sheet.getRange("A8").clearContent().setBackground(null).setFontColor(null); // Reset Cell A8
    sheet.getRange("C14:C18").clearContent();  
    sheet.getRange("E14:E18").clearContent();  
    sheet.getRange("B22:F23").clearContent();
    
    // 2. Clear Results Table & Formatting (A27:L100)
    var tableRange = sheet.getRange(ranges.CLEAR_RESULTS_RANGE);   
    tableRange.clearContent();   
    tableRange.setBackground(null).setFontColor(null).setFontWeight("normal");   

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
        var cleanWoPart = woPartNumber.toLowerCase().replace(/[-_\s]/g, "");  
        
        var regProgIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.PROGRAM_NAME - 1; 
        var regBaseModelIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.BASE_MODEL - 1; 
        
        for (var rR = 1; rR < regValues.length; rR++) {  
          var regRow = regValues[rR];
          var regPartClean = String(regRow[regBaseModelIdx] || "").trim().toLowerCase().replace(/[-_\s]/g, "");  
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
 * Populates spec limit cells (B22:F23) with positive absolute values.
 */
function populateSpecLimits(ss, sheet, programOrPart) {
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);
  if (!refSheet || !programOrPart) return;

  var refData = refSheet.getDataRange().getValues();
  var refCols = CONFIG.COLUMNS.PART_REFERENCE_MATRIX;
  var cleanKey = String(programOrPart).trim().toLowerCase();

  var refRow = null;
  for (var i = 1; i < refData.length; i++) {
    var progName = String(refData[i][refCols.PROGRAM_NAME - 1] || "").trim().toLowerCase();
    if (progName === cleanKey) {
      refRow = refData[i];
      break;
    }
  }

  if (!refRow) return;

  var ranges = CONFIG.OPERATOR_STATION.RANGES;

  function getAbsPair(minVal, maxVal) {
    if (minVal === "" || maxVal === "") return { min: "", max: "" };
    var a = Math.abs(parseFloat(minVal));
    var b = Math.abs(parseFloat(maxVal));
    return { min: Math.min(a, b), max: Math.max(a, b) };
  }

  var c1 = getAbsPair(refRow[refCols.COMP_1_MIN - 1], refRow[refCols.COMP_1_MAX - 1]);
  var r1 = getAbsPair(refRow[refCols.REB_1_MIN - 1], refRow[refCols.REB_1_MAX - 1]);
  var c2 = getAbsPair(refRow[refCols.COMP_2_MIN - 1], refRow[refCols.COMP_2_MAX - 1]);
  var r2 = getAbsPair(refRow[refCols.REB_2_MIN - 1], refRow[refCols.REB_2_MAX - 1]);

  sheet.getRange(ranges.LIMIT_C1_MIN).setValue(c1.min);
  sheet.getRange(ranges.LIMIT_C1_MAX).setValue(c1.max);
  sheet.getRange(ranges.LIMIT_R1_MIN).setValue(r1.min);
  sheet.getRange(ranges.LIMIT_R1_MAX).setValue(r1.max);

  sheet.getRange(ranges.LIMIT_C2_MIN).setValue(c2.min);
  sheet.getRange(ranges.LIMIT_C2_MAX).setValue(c2.max);
  sheet.getRange(ranges.LIMIT_R2_MIN).setValue(r2.min);
  sheet.getRange(ranges.LIMIT_R2_MAX).setValue(r2.max);

  sheet.getRange(ranges.LIMIT_SLOPE).setValue(safeAbsNum(refRow[refCols.SLOPE_1_MIN - 1]));
}

/**
 * Queries Master_Dyno_Log directly for records matching the scanned barcode prefix.
 * Deduplicates by True Serial (keeping the latest test entry), maps Col V & W to Col G & H,
 * and dynamically formats cell A8 Work Order status with color indicators.
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

  // Read active spec limits directly from sheet cells B22:F23
  var c1Min = parseFloat(sheet.getRange(ranges.LIMIT_C1_MIN).getValue());
  var c1Max = parseFloat(sheet.getRange(ranges.LIMIT_C1_MAX).getValue());
  var r1Min = parseFloat(sheet.getRange(ranges.LIMIT_R1_MIN).getValue());
  var r1Max = parseFloat(sheet.getRange(ranges.LIMIT_R1_MAX).getValue());

  var c2Min = parseFloat(sheet.getRange(ranges.LIMIT_C2_MIN).getValue());
  var c2Max = parseFloat(sheet.getRange(ranges.LIMIT_C2_MAX).getValue());
  var r2Min = parseFloat(sheet.getRange(ranges.LIMIT_R2_MIN).getValue());
  var r2Max = parseFloat(sheet.getRange(ranges.LIMIT_R2_MAX).getValue());

  var cleanBarcode = String(searchBarcode).trim().toLowerCase();
  var cleanPart = String(partNumber).trim().toLowerCase();

  // Determine Col V (22) and Col W (23) column indices (0-indexed)
  var colTest1Idx = (logCols && logCols.TEST_1_STATUS) ? (logCols.TEST_1_STATUS - 1) : 21; // Col V
  var colTest2Idx = (logCols && logCols.TEST_2_STATUS) ? (logCols.TEST_2_STATUS - 1) : 22; // Col W

  // 1. First Pass: Collect matching entries and keep only the LATEST entry per serial
  var latestLogBySerial = {};

  for (var r = 1; r < logData.length; r++) {
    var row = logData[r];
    var trueSerial = String(row[logCols.TRUE_SERIAL - 1] || "").trim();
    var baseModel = String(row[logCols.BASE_MODEL - 1] || "").trim().toLowerCase();

    var isMatch = false;
    if (cleanBarcode !== "" && trueSerial.toLowerCase().includes(cleanBarcode)) {
      isMatch = true;
    } else if (cleanPart !== "" && baseModel === cleanPart && (cleanBarcode === "" || cleanBarcode === "undefined")) {
      isMatch = true;
    }

    if (isMatch && trueSerial !== "") {
      latestLogBySerial[trueSerial.toLowerCase()] = {
        rowIdx: r + 1,
        data: row,
        trueSerial: trueSerial
      };
    }
  }

  var rowsToDisplay = [];
  var test1FailCount = 0;
  var test2FailCount = 0;

  // Sort serials sequentially
  var serialKeys = Object.keys(latestLogBySerial).sort();

  // 2. Second Pass: Build table rows and evaluate test outcomes
  for (var k = 0; k < serialKeys.length; k++) {
    var item = latestLogBySerial[serialKeys[k]];
    var row = item.data;
    var actualSheetRow = item.rowIdx;
    var trueSerial = item.trueSerial;

    var rowLink = "#gid=" + logSheetId + "&range=A" + actualSheetRow;
    var serialHyperlinkFormula = '=HYPERLINK("' + rowLink + '", "' + trueSerial + '")';

    var t1Status = String(row[colTest1Idx] || "").trim();
    var t2Status = String(row[colTest2Idx] || "").trim();

    var t1Upper = t1Status.toUpperCase();
    var t2Upper = t2Status.toUpperCase();

    // Flexible string check for any variation of "FAIL"
    if (t1Upper.includes("FAIL")) {
      test1FailCount++;
    }
    if (t2Upper.includes("FAIL")) {
      test2FailCount++;
    }

    var overallStat = String(row[logCols.OVERALL_STATUS - 1] || "").trim();
    var teardownAct = String(row[logCols.TEARDOWN_ACTION - 1] || "").trim();
    var evalActionCombo = overallStat + (teardownAct ? " / " + teardownAct : "");

    var diagNotes = String(row[logCols.DIAGNOSTICS - 1] || "").trim();
    var engComm = String(row[logCols.ENG_COMMENTS - 1] || "").trim();
    var diagCombo = diagNotes + (engComm ? " | " + engComm : "");

    // 12-COLUMN TABLE MAPPING
    var mappedRow = [
      serialHyperlinkFormula,                     // Col A (1): True Serial (Hyperlink)
      safeAbsNum(row[logCols.ROD_FORCE - 1]),     // Col B (2): Rod Force (ABS)
      safeAbsNum(row[logCols.COMP_1 - 1]),        // Col C (3): Low Speed Comp (ABS)
      safeAbsNum(row[logCols.REB_1 - 1]),         // Col D (4): Low Speed Reb (ABS)
      safeAbsNum(row[logCols.COMP_2 - 1]),        // Col E (5): Med Speed Comp (ABS)
      safeAbsNum(row[logCols.REB_2 - 1]),         // Col F (6): Med Speed Reb (ABS)
      t1Status,                                   // Col G (7): Test 1 Status (Log Col V)
      t2Status,                                   // Col H (8): Test 2 Status (Log Col W)
      safeAbsNum(row[logCols.COMP_3 - 1]),        // Col I (9): High Speed Comp (ABS)
      safeAbsNum(row[logCols.REB_3 - 1]),         // Col J (10): High Speed Reb (ABS)
      evalActionCombo,                            // Col K (11): Overall Status / Eval Action
      diagCombo                                   // Col L (12): Diagnostic Notes
    ];

    rowsToDisplay.push(mappedRow);
  }

  // 3. Accurate Cell A8 Work Order Status Evaluation with dynamic formatting
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

  // 5. Set Decimal Formatting (X.X format on force columns B..F and I..J, skipping G & H)
  sheet.getRange(startRow, 2, numRows, 5).setNumberFormat("0.0"); // Cols B..F
  sheet.getRange(startRow, 9, numRows, 2).setNumberFormat("0.0"); // Cols I..J

  // 6. Build Formatting Arrays for Batch Highlighting
  var bgColors = [];
  var fontColors = [];

  for (var rIdx = 0; rIdx < numRows; rIdx++) {
    var rowBg = [];
    var rowFont = [];
    var rowData = rowsToDisplay[rIdx];

    for (var cIdx = 0; cIdx < numCols; cIdx++) {
      var val = parseFloat(rowData[cIdx]);
      var isOut = false;

      if (!isNaN(val)) {
        if (cIdx === 2 && ((!isNaN(c1Min) && val < c1Min) || (!isNaN(c1Max) && val > c1Max))) isOut = true; // Low Comp
        if (cIdx === 3 && ((!isNaN(r1Min) && val < r1Min) || (!isNaN(r1Max) && val > r1Max))) isOut = true; // Low Reb
        if (cIdx === 4 && ((!isNaN(c2Min) && val < c2Min) || (!isNaN(c2Max) && val > c2Max))) isOut = true; // Med Comp
        if (cIdx === 5 && ((!isNaN(r2Min) && val < r2Min) || (!isNaN(r2Max) && val > r2Max))) isOut = true; // Med Reb
      }

      if (isOut) {
        rowBg.push("#FFCCCC");   // Production Light Red
        rowFont.push("#990000"); // Production Dark Red
      } else {
        rowBg.push(null);        
        rowFont.push(null);
      }
    }
    bgColors.push(rowBg);
    fontColors.push(rowFont);
  }

  // 7. Batch apply background and text formatting in a single API call
  outputRange.setBackgrounds(bgColors).setFontColors(fontColors);
}