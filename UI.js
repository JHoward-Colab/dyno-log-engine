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
    // 1. Clear Metadata Areas & Panels (C3:C6, A14:E18, B22:F23)
    sheet.getRange(ranges.CLEAR_METADATA_RANGE).clearContent(); 
    sheet.getRange("C6").clearContent(); 
    sheet.getRange("A14:E18").clearContent();  
    sheet.getRange("B22:F23").clearContent();
    sheet.getRange(ranges.PROGRAM_NAME_OUTPUT).clearContent();   
    
    // 2. Clear Results Table & Formatting (A27:L100)
    var tableRange = sheet.getRange(ranges.CLEAR_RESULTS_RANGE);   
    tableRange.clearContent();   
    tableRange.setBackground(null).setFontColor(null).setFontWeight("normal").setNumberFormat("@");   

    if (!barcode || barcode === "undefined" || barcode === "null") return;  
      
    var searchBarcode = barcode;  
    if (barcode.includes("-")) { searchBarcode = barcode.split("-")[0].trim(); }   
    else if (barcode.includes("_")) { searchBarcode = barcode.split("_")[0].trim(); }  
    if (!isNaN(searchBarcode) && searchBarcode.length === 4) { searchBarcode = "00" + searchBarcode; }  
      
    var searchCriteria = "title contains '" + searchBarcode + "' and (mimeType = 'application/vnd.google-apps.spreadsheet' or mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') and trashed = false";  
    var files = DriveApp.searchFiles(searchCriteria);  
    if (!files.hasNext()) { 
      sheet.getRange(ranges.FILE_LINK_OUTPUT).setValue("❌ Work Order File Not Found: " + searchBarcode); 
      sheet.getRange("C6").setValue("FAILED CROSS-CHECK");
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
      var woPartNumber = woSheet.getRange("D3").getValue(); 
      var woBomRevision = woSheet.getRange("D4").getValue();   
      
      sheet.getRange(ranges.BOM_REV_OUTPUT).setValue(woBomRevision); 
      sheet.getRange(ranges.PART_NO_OUTPUT).setValue(woPartNumber);   

      // ✅ RESTORED: Cell C6 Production Cross-Check Verification
      sheet.getRange("C6").setValue("MATCH: " + searchBarcode);
      
      var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);  
      var matchedProgramName = "";

      if (registrySheet && woPartNumber) {  
        var regValues = registrySheet.getDataRange().getValues();  
        var cleanWoPart = String(woPartNumber).trim().toLowerCase().replace(/[-_\s]/g, "");  
        
        var regProgIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.PROGRAM_NAME - 1;
        var regBaseModelIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.BASE_MODEL - 1;
        
        for (var rR = 1; rR < regValues.length; rR++) {  
          var regPartClean = String(regValues[rR][regBaseModelIdx] || "").trim().toLowerCase().replace(/[-_\s]/g, "");  
          var regProgName = String(regValues[rR][regProgIdx] || "").trim();  
          
          if (regPartClean === cleanWoPart && regProgName) {  
            matchedProgramName = regProgName;
            sheet.getRange(ranges.PROGRAM_NAME_OUTPUT).setValue(regProgName);  
            break;
          }  
        }  
      }  

      // 3. Populate A14:E18 Spec Summary Panel & Limit Cells (B22:F23)
      populateSpecSummaryAndLimits(ss, sheet, matchedProgramName || woPartNumber);

      // 4. Render re-mapped dyno records with X.X formatting & out-of-spec red highlighting
      renderOperatorTableWithFormatting(ss, sheet, woPartNumber);

    } catch(e) {
      Logger.log("WO Lookup Error: " + e.toString());
    }
  }  
}  

/**
 * Populates A14:E18 summary panel and B22:F23 limit cells from Part_Reference_Matrix.
 */
function populateSpecSummaryAndLimits(ss, sheet, programOrPart) {
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

  // Extract Speeds & Specs
  var c1Min = refRow[refCols.COMP_1_MIN - 1];
  var c1Max = refRow[refCols.COMP_1_MAX - 1];
  var r1Min = refRow[refCols.REB_1_MIN - 1];
  var r1Max = refRow[refCols.REB_1_MAX - 1];

  var c2Min = refRow[refCols.COMP_2_MIN - 1];
  var c2Max = refRow[refCols.COMP_2_MAX - 1];
  var r2Min = refRow[refCols.REB_2_MIN - 1];
  var r2Max = refRow[refCols.REB_2_MAX - 1];

  var slopeMin = refRow[refCols.SLOPE_1_MIN - 1];
  var loopAreaMin = refRow[refCols.LOOP_AREA_1_MIN - 1];

  // 1. Populate A14:E18 Summary Grid
  sheet.getRange("A14").setValue("Low Speed");
  sheet.getRange("C14").setValue(c1Min + " / " + c1Max);
  sheet.getRange("E14").setValue(r1Min + " / " + r1Max);

  sheet.getRange("A15").setValue("High Speed");
  sheet.getRange("C15").setValue(c2Min + " / " + c2Max);
  sheet.getRange("E15").setValue(r2Min + " / " + r2Max);

  sheet.getRange("C16").setValue(programOrPart);
  sheet.getRange("E16").setValue(programOrPart);

  sheet.getRange("A17").setValue("Slope / Loop");
  sheet.getRange("C17").setValue(slopeMin);
  sheet.getRange("E17").setValue(loopAreaMin);

  // 2. Populate B22:F23 Explicit Limit Cells
  var ranges = CONFIG.OPERATOR_STATION.RANGES;
  sheet.getRange(ranges.LIMIT_C1_MIN).setValue(c1Min);
  sheet.getRange(ranges.LIMIT_C1_MAX).setValue(c1Max);
  sheet.getRange(ranges.LIMIT_R1_MIN).setValue(r1Min);
  sheet.getRange(ranges.LIMIT_R1_MAX).setValue(r1Max);

  sheet.getRange(ranges.LIMIT_C2_MIN).setValue(c2Min);
  sheet.getRange(ranges.LIMIT_C2_MAX).setValue(c2Max);
  sheet.getRange(ranges.LIMIT_R2_MIN).setValue(r2Min);
  sheet.getRange(ranges.LIMIT_R2_MAX).setValue(r2Max);

  sheet.getRange(ranges.LIMIT_SLOPE).setValue(slopeMin);
}

/**
 * Renders dyno log rows for the selected part, mapping raw log columns 
 * to Operator Station table columns, formatting to X.X, and applying out-of-spec red styling.
 */
function renderOperatorTableWithFormatting(ss, sheet, partNumber) {
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);
  if (!logSheet || !refSheet) return;

  var logData = logSheet.getDataRange().getValues();
  var refData = refSheet.getDataRange().getValues();
  if (logData.length <= 1) return;

  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG;
  var refCols = CONFIG.COLUMNS.PART_REFERENCE_MATRIX;
  var ranges = CONFIG.OPERATOR_STATION.RANGES;

  // Extract part spec limits from Part_Reference_Matrix
  var limits = null;
  var cleanPart = String(partNumber).trim().toLowerCase();
  for (var i = 1; i < refData.length; i++) {
    var partNameInMatrix = String(refData[i][refCols.PROGRAM_NAME - 1] || "").trim().toLowerCase();
    if (partNameInMatrix === cleanPart) {
      limits = refData[i];
      break;
    }
  }

  // Filter & MAP raw log columns to Operator Station 12-column table
  var rowsToDisplay = [];
  var modelColIdx = logCols.BASE_MODEL - 1;

  for (var r = 1; r < logData.length; r++) {
    var row = logData[r];
    if (String(row[modelColIdx] || "").trim().toLowerCase() === cleanPart) {
      
      // ✅ COLUMN RE-MAPPING: Maps raw log indices to expected Operator Station table headers
      var mappedRow = [
        row[logCols.TRUE_SERIAL - 1],       // Col 1 (A): True Serial Number
        row[logCols.ROD_FORCE - 1],         // Col 2 (B): Rod Force
        row[logCols.COMP_1 - 1],            // Col 3 (C): Low Speed Comp
        row[logCols.REB_1 - 1],             // Col 4 (D): Low Speed Reb
        row[logCols.SLOPE_1 - 1],           // Col 5 (E): Slope
        row[logCols.LOOP_AREA_1 - 1],       // Col 6 (F): Loop Area
        row[logCols.COMP_2 - 1],            // Col 7 (G): High Speed Comp
        row[logCols.REB_2 - 1],             // Col 8 (H): High Speed Reb
        row[logCols.COMP_3 - 1],            // Col 9 (I): Speed 3 Comp
        row[logCols.REB_3 - 1],             // Col 10 (J): Speed 3 Reb
        row[logCols.OVERALL_STATUS - 1],    // Col 11 (K): Overall Status
        row[logCols.TIMESTAMP - 1]          // Col 12 (L): Date / Timestamp
      ];

      rowsToDisplay.push(mappedRow);
    }
  }

  if (rowsToDisplay.length === 0) return;

  var startRow = ranges.RESULTS_START_ROW; // Row 27
  var numRows = rowsToDisplay.length;
  var numCols = ranges.RESULTS_COL_COUNT;   // 12 columns
  var outputRange = sheet.getRange(startRow, ranges.RESULTS_START_COL, numRows, numCols);

  // 1. Write Mapped Data Values
  outputRange.setValues(rowsToDisplay);

  // 2. Set Decimal Formatting (X.X format on numeric columns B through J)
  var numericSubRange = sheet.getRange(startRow, 2, numRows, 9);
  numericSubRange.setNumberFormat("0.0");

  // 3. Build Formatting Arrays for Batch Highlighting
  var bgColors = [];
  var fontColors = [];

  for (var rIdx = 0; rIdx < numRows; rIdx++) {
    var rowBg = [];
    var rowFont = [];
    var rowData = rowsToDisplay[rIdx];

    for (var cIdx = 0; cIdx < numCols; cIdx++) {
      var val = parseFloat(rowData[cIdx]);
      var isOut = false;

      // Check min/max bounds on mapped force columns
      if (limits && !isNaN(val)) {
        if (cIdx === 2 && limits[refCols.COMP_1_MIN - 1] !== "" && (val < limits[refCols.COMP_1_MIN - 1] || val > limits[refCols.COMP_1_MAX - 1])) isOut = true; // Comp 1
        if (cIdx === 3 && limits[refCols.REB_1_MIN - 1] !== "" && (val < limits[refCols.REB_1_MIN - 1] || val > limits[refCols.REB_1_MAX - 1])) isOut = true;   // Reb 1
        if (cIdx === 6 && limits[refCols.COMP_2_MIN - 1] !== "" && (val < limits[refCols.COMP_2_MIN - 1] || val > limits[refCols.COMP_2_MAX - 1])) isOut = true; // Comp 2
        if (cIdx === 7 && limits[refCols.REB_2_MIN - 1] !== "" && (val < limits[refCols.REB_2_MIN - 1] || val > limits[refCols.REB_2_MAX - 1])) isOut = true;   // Reb 2
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

  // 4. Batch apply background and text formatting in a single API call
  outputRange.setBackgrounds(bgColors).setFontColors(fontColors);
}