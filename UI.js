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
 * Converts raw serial strings into standard XXXX-YYY format (e.g. "1654007" -> "1654-007").
 */
function formatSerialWithDash(raw) {
  if (raw === null || raw === undefined) return "";
  var str = String(raw).trim();
  
  // Match 4 digits followed by an optional separator and 3 digits (e.g. 1654007 or 1654-007)
  var match = str.match(/(\d{4})[-_ ]?(\d{3})/);
  if (match) {
    return match[1] + "-" + match[2];
  }
  
  // Fallback: Extract all digits, take last 7 digits, insert dash
  var digits = str.replace(/\D/g, "");
  if (digits.length >= 7) {
    var last7 = digits.slice(-7);
    return last7.slice(0, 4) + "-" + last7.slice(4);
  }
  
  return str;
}

/**
 * Normalizes serial strings for key matching.
 */
function normalizeSerialKey(str) {
  if (str === null || str === undefined) return "";
  var formatted = formatSerialWithDash(str);
  return String(formatted).trim().toLowerCase().replace(/[^a-z0-9]/g, "");
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
    sheet.getRange("A8").clearContent(); 
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

      // 4. Render dyno records using formatted dash serial matching
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
 * Dynamically resolves Master_Dyno_Log column header positions.
 */
function getLogHeaderIndices(headerRow) {
  var cols = CONFIG.COLUMNS.MASTER_DYNO_LOG;
  var map = {
    trueSerial: cols.TRUE_SERIAL - 1,
    baseModel: cols.BASE_MODEL - 1,
    rodForce: cols.ROD_FORCE - 1,
    comp1: cols.COMP_1 - 1,
    reb1: cols.REB_1 - 1,
    comp2: cols.COMP_2 - 1,
    reb2: cols.REB_2 - 1,
    comp3: cols.COMP_3 - 1,
    reb3: cols.REB_3 - 1,
    test1Status: cols.TEST_1_STATUS - 1,
    test2Status: cols.TEST_2_STATUS - 1,
    overallStatus: cols.OVERALL_STATUS - 1,
    teardownAction: cols.TEARDOWN_ACTION - 1,
    diagnostics: cols.DIAGNOSTICS - 1,
    engComments: cols.ENG_COMMENTS - 1
  };

  if (!headerRow) return map;

  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i] || "").trim().toLowerCase();
    if (h.includes("true serial") || h.includes("serial number")) map.trueSerial = i;
    else if (h.includes("base model")) map.baseModel = i;
    else if (h.includes("rod force")) map.rodForce = i;
    else if (h.includes("comp peak 1") || h.includes("comp 1")) map.comp1 = i;
    else if (h.includes("reb peak 1") || h.includes("reb 1")) map.reb1 = i;
    else if (h.includes("comp peak 2") || h.includes("comp 2")) map.comp2 = i;
    else if (h.includes("reb peak 2") || h.includes("reb 2")) map.reb2 = i;
    else if (h.includes("comp peak 3") || h.includes("comp 3")) map.comp3 = i;
    else if (h.includes("reb peak 3") || h.includes("reb 3")) map.reb3 = i;
    else if (h.includes("test 1 status") || h.includes("test1 status")) map.test1Status = i;
    else if (h.includes("test 2 status") || h.includes("test2 status")) map.test2Status = i;
    else if (h.includes("overall status") || h.includes("status")) map.overallStatus = i;
    else if (h.includes("teardown")) map.teardownAction = i;
    else if (h.includes("diagnostics")) map.diagnostics = i;
    else if (h.includes("comments")) map.engComments = i;
  }
  return map;
}

/**
 * Extracts serials from WO file, formats to XXXX-YYY, and cross-references Master_Dyno_Log.
 */
function renderOperatorTableWithWOAlignment(ss, sheet, woSheet, searchBarcode, partNumber) {
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  var cleanBarcodeDigits = String(searchBarcode).replace(/\D/g, "");

  // 1. Extract and format master serial list directly from the Work Order File
  var woValues = woSheet.getDataRange().getValues();
  var woSerials = [];

  for (var r = 0; r < woValues.length; r++) {
    for (var c = 0; c < woValues[r].length; c++) {
      var cellVal = String(woValues[r][c] || "").trim();
      if (!cellVal) continue;

      var cellDigits = cellVal.replace(/\D/g, "");

      // Match cells containing WO prefix digits and having unit serial length (7+ digits)
      if (cleanBarcodeDigits !== "" && cellDigits.includes(cleanBarcodeDigits) && cellDigits.length >= 7) {
        var formattedSerial = formatSerialWithDash(cellVal);
        if (formattedSerial && formattedSerial !== searchBarcode && !woSerials.includes(formattedSerial)) {
          woSerials.push(formattedSerial);
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

  // 2. Index Master_Dyno_Log using normalized serial keys
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  var logSheetId = logSheet ? logSheet.getSheetId() : null;
  var latestTestsByNormKey = {};
  var logMap = null;

  if (logSheet) {
    var logData = logSheet.getDataRange().getValues();
    if (logData.length > 0) {
      logMap = getLogHeaderIndices(logData[0]);
      for (var i = 1; i < logData.length; i++) {
        var row = logData[i];
        var trueSerial = String(row[logMap.trueSerial] || "").trim();
        var normKey = normalizeSerialKey(trueSerial);
        if (normKey !== "") {
          latestTestsByNormKey[normKey] = {
            rowIdx: i + 1,
            data: row
          };
        }
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

  // 3. Build mapped rows matching formatted WO serials against normalized Log Keys
  for (var s = 0; s < woSerials.length; s++) {
    var sNum = woSerials[s]; // Formatted as XXXX-YYY
    var normSKey = normalizeSerialKey(sNum);
    var testRecord = latestTestsByNormKey[normSKey];

    if (testRecord && logMap) {
      testedCount++;
      var row = testRecord.data;
      var actualRow = testRecord.rowIdx;
      var rowLink = "#gid=" + logSheetId + "&range=A" + actualRow;
      var serialHyperlinkFormula = '=HYPERLINK("' + rowLink + '", "' + sNum + '")';

      var t1Status = String(row[logMap.test1Status] || "").trim();
      var t2Status = String(row[logMap.test2Status] || "").trim();

      if (t1Status.toUpperCase() === "PASS") test1PassCount++;
      else if (t1Status.toUpperCase() === "FAIL") test1FailCount++;

      if (t2Status.toUpperCase() === "PASS") test2PassCount++;
      else if (t2Status.toUpperCase() === "FAIL") test2FailCount++;

      var overallStat = String(row[logMap.overallStatus] || "").trim();
      var teardownAct = String(row[logMap.teardownAction] || "").trim();
      var evalActionCombo = overallStat + (teardownAct ? " / " + teardownAct : "");

      var diagNotes = String(row[logMap.diagnostics] || "").trim();
      var engComm = String(row[logMap.engComments] || "").trim();
      var diagCombo = diagNotes + (engComm ? " | " + engComm : "");

      // 12-COLUMN TABLE MAPPING (Col G = Test 1, Col H = Test 2)
      var mappedRow = [
        serialHyperlinkFormula,               // Col A (1): True Serial (Hyperlink)
        safeAbsNum(row[logMap.rodForce]),     // Col B (2): Rod Force (ABS)
        safeAbsNum(row[logMap.comp1]),        // Col C (3): Low Speed Comp (ABS)
        safeAbsNum(row[logMap.reb1]),         // Col D (4): Low Speed Reb (ABS)
        safeAbsNum(row[logMap.comp2]),        // Col E (5): Med Speed Comp (ABS)
        safeAbsNum(row[logMap.reb2]),         // Col F (6): Med Speed Reb (ABS)
        t1Status,                             // Col G (7): Test 1 Status
        t2Status,                             // Col H (8): Test 2 Status
        safeAbsNum(row[logMap.comp3]),        // Col I (9): High Speed Comp (ABS)
        safeAbsNum(row[logMap.reb3]),         // Col J (10): High Speed Reb (ABS)
        evalActionCombo,                      // Col K (11): Overall Status / Eval Action
        diagCombo                             // Col L (12): Diagnostic Notes
      ];
      rowsToDisplay.push(mappedRow);
    } else {
      // Un-tested unit: displays formatted serial with blank data slots
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