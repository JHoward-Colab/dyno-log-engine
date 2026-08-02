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
 * Helper to safely extract positive absolute numbers formatted or null.
 */
function safeAbsNum(val) {
  if (val === "" || val === null || val === undefined) return "";
  var n = parseFloat(val);
  return isNaN(n) ? val : Math.abs(n);
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
    // 1. Clear Value Cells Only (Preserving static label cells A14:A18)
    sheet.getRange(ranges.CLEAR_METADATA_RANGE).clearContent(); 
    sheet.getRange("C6").clearContent(); 
    sheet.getRange("A8").clearContent(); // Work Order Status
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
      sheet.getRange("A8").setValue("WORK ORDER FILE NOT FOUND");
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

      // 4. Render dyno records using Work Order serial list as master truth
      renderOperatorTableWithWOAlignment(ss, sheet, woSheet, searchBarcode, woPartNumber);

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
 * Extracts all expected serials directly from the Work Order spreadsheet.
 * Cross-references Master_Dyno_Log for test results, maps Test 1 to Col G & Test 2 to Col H,
 * and updates cell A8 with Overall Work Order Status.
 */
function renderOperatorTableWithWOAlignment(ss, sheet, woSheet, searchBarcode, partNumber) {
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  var cleanBarcode = String(searchBarcode).trim().toLowerCase();

  // 1. Extract master serial list directly from the Work Order File
  var woValues = woSheet.getDataRange().getValues();
  var woSerials = [];

  for (var r = 0; r < woValues.length; r++) {
    for (var c = 0; c < woValues[r].length; c++) {
      var cellVal = String(woValues[r][c] || "").trim();
      if (cellVal.toLowerCase().includes(cleanBarcode) && cellVal !== searchBarcode) {
        if (!woSerials.includes(cellVal)) {
          woSerials.push(cellVal);
        }
      }
    }
  }

  // Sort WO serials sequentially by unit suffix
  woSerials.sort(function(a, b) {
    var mA = a.match(/[-_](\d+)$/);
    var mB = b.match(/[-_](\d+)$/);
    var uA = mA ? parseInt(mA[1], 10) : 0;
    var uB = mB ? parseInt(mB[1], 10) : 0;
    return uA - uB;
  });

  // 2. Query Master_Dyno_Log for test results
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG;
  var logSheetId = logSheet ? logSheet.getSheetId() : null;

  var latestTestsBySerial = {};

  if (logSheet) {
    var logData = logSheet.getDataRange().getValues();
    for (var i = 1; i < logData.length; i++) {
      var row = logData[i];
      var trueSerial = String(row[logCols.TRUE_SERIAL - 1] || "").trim();
      if (trueSerial.toLowerCase().includes(cleanBarcode)) {
        latestTestsBySerial[trueSerial] = {
          rowIdx: i + 1,
          data: row
        };
      }
    }
  }

  // Read active spec limits directly from sheet cells B22:F23
  var c1Min = parseFloat(sheet.getRange(ranges.LIMIT_C1_MIN).getValue());
  var c1Max = parseFloat(sheet.getRange(ranges.LIMIT_C1_MAX).getValue());
  var r1Min = parseFloat(sheet.getRange(ranges.LIMIT_R1_MIN).getValue());
  var r1Max = parseFloat(sheet.getRange(ranges.LIMIT_R1_MAX).getValue());

  var c2Min = parseFloat(sheet.getRange(ranges.LIMIT_C2_MIN).getValue());
  var c2Max = parseFloat(sheet.getRange(ranges.LIMIT_C2_MAX).getValue());
  var r2Min = parseFloat(sheet.getRange(ranges.LIMIT_R2_MIN).getValue());
  var r2Max = parseFloat(sheet.getRange(ranges.LIMIT_R2_MAX).getValue());

  var rowsToDisplay = [];
  var testedCount = 0;
  var test1PassCount = 0;
  var test1FailCount = 0;
  var test2PassCount = 0;
  var test2FailCount = 0;

  // 3. Build mapped rows using WO Serials as Master Truth
  for (var s = 0; s < woSerials.length; s++) {
    var sNum = woSerials[s];
    var testRecord = latestTestsBySerial[sNum];

    if (testRecord) {
      testedCount++;
      var row = testRecord.data;
      var actualRow = testRecord.rowIdx;
      var rowLink = "#gid=" + logSheetId + "&range=A" + actualRow;
      var serialHyperlinkFormula = '=HYPERLINK("' + rowLink + '", "' + sNum + '")';

      var t1Status = String(row[logCols.TEST_1_STATUS - 1] || "").trim();
      var t2Status = String(row[logCols.TEST_2_STATUS - 1] || "").trim();

      if (t1Status.toUpperCase() === "PASS") test1PassCount++;
      else if (t1Status.toUpperCase() === "FAIL") test1FailCount++;

      if (t2Status.toUpperCase() === "PASS") test2PassCount++;
      else if (t2Status.toUpperCase() === "FAIL") test2FailCount++;

      var overallStat = String(row[logCols.OVERALL_STATUS - 1] || "").trim();
      var teardownAct = String(row[logCols.TEARDOWN_ACTION - 1] || "").trim();
      var evalActionCombo = overallStat + (teardownAct ? " / " + teardownAct : "");

      var diagNotes = String(row[logCols.DIAGNOSTICS - 1] || "").trim();
      var engComm = String(row[logCols.ENG_COMMENTS - 1] || "").trim();
      var diagCombo = diagNotes + (engComm ? " | " + engComm : "");

      // ✅ CORRECTED 12-COLUMN MAPPING (Col G = Test 1, Col H = Test 2)
      var mappedRow = [
        serialHyperlinkFormula,                     // Col A (1): True Serial (Hyperlink)
        safeAbsNum(row[logCols.ROD_FORCE - 1]),     // Col B (2): Rod Force (ABS)
        safeAbsNum(row[logCols.COMP_1 - 1]),        // Col C (3): Low Speed Comp (ABS)
        safeAbsNum(row[logCols.REB_1 - 1]),         // Col D (4): Low Speed Reb (ABS)
        safeAbsNum(row[logCols.COMP_2 - 1]),        // Col E (5): Med Speed Comp (ABS)
        safeAbsNum(row[logCols.REB_2 - 1]),         // Col F (6): Med Speed Reb (ABS)
        t1Status,                                   // Col G (7): Test 1 Status
        t2Status,                                   // Col H (8): Test 2 Status
        safeAbsNum(row[logCols.COMP_3 - 1]),        // Col I (9): High Speed Comp (ABS)
        safeAbsNum(row[logCols.REB_3 - 1]),         // Col J (10): High Speed Reb (ABS)
        evalActionCombo,                            // Col K (11): Overall Status / Eval Action
        diagCombo                                   // Col L (12): Diagnostic Notes
      ];
      rowsToDisplay.push(mappedRow);
    } else {
      // Missing test result: display serial with empty data slots
      rowsToDisplay.push([sNum, "", "", "", "", "", "", "", "", "", "", ""]);
    }
  }

  // 4. Update Work Order Status in Cell A8
  var totalSerials = woSerials.length;
  var statusMessage = "";

  if (totalSerials === 0) {
    statusMessage = "NO SERIALS FOUND IN WORK ORDER";
  } else if (testedCount === 0) {
    statusMessage = "PENDING TESTING";
  } else if (test1FailCount > 0) {
    statusMessage = "ACTION REQUIRED: Test 1 Failure Detected (" + test1FailCount + " unit(s))";
  } else if (test2FailCount > 0) {
    statusMessage = "CONDITIONAL PASS: Attention Required (Test 2 Outlier Detected)";
  } else if (testedCount === totalSerials && test1FailCount === 0 && test2FailCount === 0) {
    statusMessage = "WORK ORDER COMPLETED AND PASSING";
  } else {
    statusMessage = "IN PROGRESS: " + testedCount + "/" + totalSerials + " Units Tested (All Passing)";
  }

  sheet.getRange("A8").setValue(statusMessage);

  if (rowsToDisplay.length === 0) return;

  var startRow = ranges.RESULTS_START_ROW; // Row 27
  var numRows = rowsToDisplay.length;
  var numCols = ranges.RESULTS_COL_COUNT;   // 12 columns
  var outputRange = sheet.getRange(startRow, ranges.RESULTS_START_COL, numRows, numCols);

  // 5. Write Mapped Data Values & Formulas
  outputRange.setValues(rowsToDisplay);

  // 6. Set Decimal Formatting (X.X format on force columns B through F, I, J)
  sheet.getRange(startRow, 2, numRows, 5).setNumberFormat("0.0");
  sheet.getRange(startRow, 9, numRows, 2).setNumberFormat("0.0");

  // 7. Build Formatting Arrays for Batch Highlighting
  var bgColors = [];
  var fontColors = [];

  for (var rIdx = 0; rIdx < numRows; rIdx++) {
    var rowBg = [];
    var rowFont = [];
    var rowData = rowsToDisplay[rIdx];

    for (var cIdx = 0; cIdx < numCols; cIdx++) {
      var val = parseFloat(rowData[cIdx]);
      var isOut = false;

      if (!isNaN(val) && rowData[0] !== "") {
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

  // 8. Apply background and text formatting in a single API call
  outputRange.setBackgrounds(bgColors).setFontColors(fontColors);
}