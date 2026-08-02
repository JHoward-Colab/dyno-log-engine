// =========================================================================
// 🖥️ USER INTERFACE & CONTROLLERS (UI.js)
// Workspace Rendering, Button Actions & Triggers
// =========================================================================

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

function safeAbsNum(val) {
  if (val === "" || val === null || val === undefined) return "";
  var n = parseFloat(val);
  return isNaN(n) ? val : Math.abs(n);
}

function cleanKey(val) {
  if (val === null || val === undefined) return "";
  return String(val).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Safely converts raw signed limits into sorted positive absolute bounds.
 */
function getAbsPair(val1, val2) {
  var a = Math.abs(parseFloat(val1));
  var b = Math.abs(parseFloat(val2));
  if (isNaN(a) && isNaN(b)) return { min: NaN, max: NaN };
  if (isNaN(a)) return { min: b, max: b };
  if (isNaN(b)) return { min: a, max: a };
  return { min: Math.min(a, b), max: Math.max(a, b) };
}

function setA8Status(sheet, statusMessage) {
  var a8 = sheet.getRange("A8");
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

function manageOperatorStation(e) {
  var ss = e ? e.source : SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OPERATOR_STATION); 
  if (!sheet) return;
  
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  var range = e ? e.range : sheet.getRange(ranges.BARCODE_INPUT); 
  var barcode = String(sheet.getRange(ranges.BARCODE_INPUT).getValue()).trim();  
    
  if (range.getA1Notation() === ranges.BARCODE_INPUT) {  
    sheet.getRange(ranges.CLEAR_METADATA_RANGE).clearContent(); 
    sheet.getRange("C6").clearContent(); 
    sheet.getRange("A8").clearContent().setBackground(null).setFontColor(null);
    sheet.getRange("C14:C18").clearContent();  
    sheet.getRange("E14:E18").clearContent();  
    sheet.getRange("B22:F23").clearContent();
    
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
            sheet.getRange("C14").setValue(regRow[3] || ""); 
            sheet.getRange("C15").setValue(regRow[4] || ""); 
            sheet.getRange("C16").setValue(regRow[0] || ""); 
            sheet.getRange("C17").setValue(regRow[6] || ""); 
            sheet.getRange("C18").setValue(regRow[7] || ""); 
            break;
          }  
        }  
      }  

      populateSpecLimits(ss, sheet, matchedProgramName || woPartNumber);
      renderOperatorTableWithFormatting(ss, sheet, searchBarcode, woPartNumber);

    } catch(e) {
      Logger.log("WO Lookup Error: " + e.toString());
    }
  }  
}  

function populateSpecLimits(ss, sheet, programOrPart) {
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);
  if (!refSheet || !programOrPart) return;

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

  if (!refRow) return;

  var ranges = CONFIG.OPERATOR_STATION.RANGES;

  var c1 = getAbsPair(refRow[refCols.COMP_1_MIN - 1], refRow[refCols.COMP_1_MAX - 1]);
  var r1 = getAbsPair(refRow[refCols.REB_1_MIN - 1], refRow[refCols.REB_1_MAX - 1]);
  var c2 = getAbsPair(refRow[refCols.COMP_2_MIN - 1], refRow[refCols.COMP_2_MAX - 1]);
  var r2 = getAbsPair(refRow[refCols.REB_2_MIN - 1], refRow[refCols.REB_2_MAX - 1]);

  sheet.getRange(ranges.LIMIT_C1_MIN).setValue(isNaN(c1.min) ? "" : c1.min);
  sheet.getRange(ranges.LIMIT_C1_MAX).setValue(isNaN(c1.max) ? "" : c1.max);
  sheet.getRange(ranges.LIMIT_R1_MIN).setValue(isNaN(r1.min) ? "" : r1.min);
  sheet.getRange(ranges.LIMIT_R1_MAX).setValue(isNaN(r1.max) ? "" : r1.max);

  sheet.getRange(ranges.LIMIT_C2_MIN).setValue(isNaN(c2.min) ? "" : c2.min);
  sheet.getRange(ranges.LIMIT_C2_MAX).setValue(isNaN(c2.max) ? "" : c2.max);
  sheet.getRange(ranges.LIMIT_R2_MIN).setValue(isNaN(r2.min) ? "" : r2.min);
  sheet.getRange(ranges.LIMIT_R2_MAX).setValue(isNaN(r2.max) ? "" : r2.max);
}

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

  // Read active spec limits directly off the sheet display (B22:E23) and ensure min <= max
  var rawC1 = getAbsPair(sheet.getRange(ranges.LIMIT_C1_MIN).getValue(), sheet.getRange(ranges.LIMIT_C1_MAX).getValue());
  var rawR1 = getAbsPair(sheet.getRange(ranges.LIMIT_R1_MIN).getValue(), sheet.getRange(ranges.LIMIT_R1_MAX).getValue());
  var rawC2 = getAbsPair(sheet.getRange(ranges.LIMIT_C2_MIN).getValue(), sheet.getRange(ranges.LIMIT_C2_MAX).getValue());
  var rawR2 = getAbsPair(sheet.getRange(ranges.LIMIT_R2_MIN).getValue(), sheet.getRange(ranges.LIMIT_R2_MAX).getValue());

  var limits = {
    c1Min: rawC1.min, c1Max: rawC1.max,
    r1Min: rawR1.min, r1Max: rawR1.max,
    c2Min: rawC2.min, c2Max: rawC2.max,
    r2Min: rawR2.min, r2Max: rawR2.max
  };

  var cleanBarcodeStr = cleanKey(searchBarcode);
  var cleanPartStr = cleanKey(partNumber);

  var latestLogBySerial = {};

  for (var r = 1; r < logData.length; r++) {
    var row = logData[r];
    var trueSerial = String(row[logCols.TRUE_SERIAL - 1] || "").trim(); // Col C (3)
    var baseModel = cleanKey(row[logCols.BASE_MODEL - 1]);             // Col D (4)
    var cleanSerial = cleanKey(trueSerial);

    var isMatch = false;
    if (cleanBarcodeStr !== "" && cleanSerial.includes(cleanBarcodeStr)) {
      isMatch = true;
    } else if (cleanPartStr !== "" && baseModel === cleanPartStr && (cleanBarcodeStr === "" || cleanBarcodeStr === "undefined")) {
      isMatch = true;
    }

    if (isMatch && trueSerial !== "") {
      latestLogBySerial[cleanSerial] = { rowIdx: r + 1, data: row, trueSerial: trueSerial };
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

    var t1Status = String(row[logCols.TEST_1_STATUS - 1] || "").trim(); // Col V (22)
    var t2Status = String(row[logCols.TEST_2_STATUS - 1] || "").trim(); // Col W (23)

    if (t1Status.toUpperCase().includes("FAIL")) test1FailCount++;
    if (t2Status.toUpperCase().includes("FAIL")) test2FailCount++;

    var overallStat  = String(row[logCols.OVERALL_STATUS - 1] || "").trim();   // Col X (24) -> UI Col I
    var evalAction   = String(row[logCols.EVALUATION_ACTION - 1] || "").trim(); // Col Z (26) -> UI Col J
    var diagnostics  = String(row[logCols.DIAGNOSTICS - 1] || "").trim();        // Col Y (25) -> UI Col K
    var engComments  = String(row[logCols.ENG_COMMENTS - 1] || "").trim();       // Col AA (27) -> UI Col L

    var mappedRow = [
      serialHyperlinkFormula,                          // Col A (1): True Serial
      safeAbsNum(row[logCols.ROD_FORCE - 1]),          // Col B (2): Rod Force (Col F/6)
      safeAbsNum(row[logCols.COMP_1 - 1]),             // Col C (3): Comp 1 (Col H/8)
      safeAbsNum(row[logCols.REB_1 - 1]),              // Col D (4): Reb 1 (Col I/9)
      safeAbsNum(row[logCols.COMP_2 - 1]),             // Col E (5): Comp 2 (Col M/13)
      safeAbsNum(row[logCols.REB_2 - 1]),              // Col F (6): Reb 2 (Col N/14)
      t1Status,                                        // Col G (7): Test 1 Status (Col V/22)
      t2Status,                                        // Col H (8): Test 2 Status (Col W/23)
      overallStat,                                     // Col I (9): Overall Status (Col X/24)
      evalAction,                                      // Col J (10): Evaluation Action (Col Z/26)
      diagnostics,                                     // Col K (11): Diagnostics & Troubleshooting (Col Y/25)
      engComments                                      // Col L (12): Diagnostic Notes (Col AA/27)
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

  sheet.getRange(startRow, 2, numRows, 5).setNumberFormat("0.0");  
  sheet.getRange(startRow, 11, numRows, 2).setNumberFormat("0.0"); 

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
        if (cIdx === 2 && ((!isNaN(limits.c1Min) && val < limits.c1Min) || (!isNaN(limits.c1Max) && val > limits.c1Max))) isOut = true; // Col C
        if (cIdx === 3 && ((!isNaN(limits.r1Min) && val < limits.r1Min) || (!isNaN(limits.r1Max) && val > limits.r1Max))) isOut = true; // Col D
        if (cIdx === 4 && ((!isNaN(limits.c2Min) && val < limits.c2Min) || (!isNaN(limits.c2Max) && val > limits.c2Max))) isOut = true; // Col E
        if (cIdx === 5 && ((!isNaN(limits.r2Min) && val < limits.r2Min) || (!isNaN(limits.r2Max) && val > limits.r2Max))) isOut = true; // Col F
      }

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