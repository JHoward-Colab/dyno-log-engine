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
    // 1. Clear Header Metadata Area (Value cells only, preserving static column A/E labels)
    sheet.getRange(ranges.CLEAR_METADATA_RANGE).clearContent(); 
    sheet.getRange("C6").clearContent(); 
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

      // ✅ RESTORED: Cell C6 Cross-Check Verification
      sheet.getRange("C6").setValue(searchBarcode);
      
      var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);  
      var matchedProgramName = "";

      if (registrySheet && woPartNumber) {  
        var regValues = registrySheet.getDataRange().getValues();  
        var cleanWoPart = woPartNumber.toLowerCase().replace(/[-_\s]/g, "");  
        
        var regProgIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.PROGRAM_NAME - 1;
        var regBaseModelIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.BASE_MODEL - 1;
        
        for (var rR = 1; rR < regValues.length; rR++) {  
          var regPartClean = String(regValues[rR][regBaseModelIdx] || "").trim().toLowerCase().replace(/[-_\s]/g, "");  
          var regProgName = String(regValues[rR][regProgIdx] || "").trim();  
          
          if (regPartClean === cleanWoPart && regProgName) {  
            matchedProgramName = regProgName;
            sheet.getRange(ranges.PROGRAM_NAME_OUTPUT).setValue(regProgName);  
            sheet.getRange("E16").setValue(regProgName);
            break;
          }  
        }  
      }  

      // 3. Populate Limit Cells B22:F23 & Summary Values
      populateSpecLimits(ss, sheet, matchedProgramName || woPartNumber);

      // 4. Render dyno records with hyperlinks, X.X formatting & out-of-spec highlighting
      renderOperatorTableWithFormatting(ss, sheet, searchBarcode, woPartNumber);

    } catch(e) {
      Logger.log("WO Lookup Error: " + e.toString());
    }
  }  
}  

/**
 * Populates spec limit cells (B22:F23) from Part_Reference_Matrix.
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
  sheet.getRange(ranges.LIMIT_C1_MIN).setValue(refRow[refCols.COMP_1_MIN - 1]);
  sheet.getRange(ranges.LIMIT_C1_MAX).setValue(refRow[refCols.COMP_1_MAX - 1]);
  sheet.getRange(ranges.LIMIT_R1_MIN).setValue(refRow[refCols.REB_1_MIN - 1]);
  sheet.getRange(ranges.LIMIT_R1_MAX).setValue(refRow[refCols.REB_1_MAX - 1]);

  sheet.getRange(ranges.LIMIT_C2_MIN).setValue(refRow[refCols.COMP_2_MIN - 1]);
  sheet.getRange(ranges.LIMIT_C2_MAX).setValue(refRow[refCols.COMP_2_MAX - 1]);
  sheet.getRange(ranges.LIMIT_R2_MIN).setValue(refRow[refCols.REB_2_MIN - 1]);
  sheet.getRange(ranges.LIMIT_R2_MAX).setValue(refRow[refCols.REB_2_MAX - 1]);

  sheet.getRange(ranges.LIMIT_SLOPE).setValue(refRow[refCols.SLOPE_1_MIN - 1]);
}

/**
 * Renders dyno log rows filtered by Work Order / Serial Prefix, creating clickable 
 * hyperlinks back to Master_Dyno_Log, formatting to X.X, and applying out-of-spec red highlighting.
 */
function renderOperatorTableWithFormatting(ss, sheet, searchBarcode, partNumber) {
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  if (!logSheet) return;

  var logData = logSheet.getDataRange().getValues();
  if (logData.length <= 1) return;

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

  var rowsToDisplay = [];
  var rawRowIndexMap = []; // Tracks actual row index on Master_Dyno_Log

  for (var r = 1; r < logData.length; r++) {
    var row = logData[r];
    var trueSerial = String(row[logCols.TRUE_SERIAL - 1] || "").trim();
    var baseModel = String(row[logCols.BASE_MODEL - 1] || "").trim().toLowerCase();

    // Filter: Match serial against scanned WO barcode prefix
    var isMatch = false;
    if (cleanBarcode !== "" && trueSerial.toLowerCase().includes(cleanBarcode)) {
      isMatch = true;
    } else if (cleanPart !== "" && baseModel === cleanPart && (cleanBarcode === "" || cleanBarcode === "undefined")) {
      isMatch = true;
    }

    if (isMatch) {
      // ✅ HYPERLINK TO LINE ENTRY: Jump to exact row on Master_Dyno_Log
      var actualSheetRow = r + 1; // 1-based row index in sheet
      var rowLink = "#gid=" + logSheetId + "&range=A" + actualSheetRow;
      var serialHyperlinkFormula = '=HYPERLINK("' + rowLink + '", "' + trueSerial + '")';

      // Map raw log columns to Operator Station 12-column table layout
      var mappedRow = [
        serialHyperlinkFormula,             // Col 1 (A): True Serial (Hyperlinked)
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
      rawRowIndexMap.push(actualSheetRow);
    }
  }

  if (rowsToDisplay.length === 0) return;

  var startRow = ranges.RESULTS_START_ROW; // Row 27
  var numRows = rowsToDisplay.length;
  var numCols = ranges.RESULTS_COL_COUNT;   // 12 columns
  var outputRange = sheet.getRange(startRow, ranges.RESULTS_START_COL, numRows, numCols);

  // 1. Write Mapped Data Values & Formulas
  outputRange.setValues(rowsToDisplay);

  // 2. Set Decimal Formatting (X.X format on numeric force columns B through J)
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

      // Check force values against active limits
      if (!isNaN(val)) {
        if (cIdx === 2 && ((!isNaN(c1Min) && val < c1Min) || (!isNaN(c1Max) && val > c1Max))) isOut = true; // Comp 1
        if (cIdx === 3 && ((!isNaN(r1Min) && val < r1Min) || (!isNaN(r1Max) && val > r1Max))) isOut = true; // Reb 1
        if (cIdx === 6 && ((!isNaN(c2Min) && val < c2Min) || (!isNaN(c2Max) && val > c2Max))) isOut = true; // Comp 2
        if (cIdx === 7 && ((!isNaN(r2Min) && val < r2Min) || (!isNaN(r2Max) && val > r2Max))) isOut = true; // Reb 2
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

  // 4. Apply background and text formatting in a single API call
  outputRange.setBackgrounds(bgColors).setFontColors(fontColors);
}