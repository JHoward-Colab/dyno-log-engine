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
    // 1. Clear Header Metadata Area (C3:C5, C6 cross-check, and status panels)
    sheet.getRange(ranges.CLEAR_METADATA_RANGE).clearContent(); 
    sheet.getRange("C6").clearContent(); 
    sheet.getRange("C14:C18").clearContent();  
    sheet.getRange("E14:E18").clearContent();  
    sheet.getRange(ranges.PROGRAM_NAME_OUTPUT).clearContent();   
    
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

      // ✅ RESTORED: Cell C6 Cross-Check Verification Assignment
      sheet.getRange("C6").setValue(searchBarcode);
      
      var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);  
      if (registrySheet && woPartNumber) {  
        var regValues = registrySheet.getDataRange().getValues();  
        var cleanWoPart = String(woPartNumber).trim().toLowerCase().replace(/[-_\s]/g, "");  
        
        var regProgIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.PROGRAM_NAME - 1;
        var regBaseModelIdx = CONFIG.COLUMNS.PROGRAM_REGISTRY.BASE_MODEL - 1;
        
        for (var rR = 1; rR < regValues.length; rR++) {  
          var regPartClean = String(regValues[rR][regBaseModelIdx] || "").trim().toLowerCase().replace(/[-_\s]/g, "");  
          var regProgName = String(regValues[rR][regProgIdx] || "").trim();  
          
          if (regPartClean === cleanWoPart && regProgName) {  
            sheet.getRange(ranges.PROGRAM_NAME_OUTPUT).setValue(regProgName);  
            sheet.getRange("E16").setValue(regProgName);  
            break;
          }  
        }  
      }  

      // 3. Render matching records and apply production conditional formatting
      renderOperatorTableWithFormatting(ss, sheet, woPartNumber);

    } catch(e) {
      Logger.log("WO Lookup Error: " + e.toString());
    }
  }  
}  

/**
 * Renders dyno log rows for the selected part and applies batch 
 * conditional formatting (Red highlighting for out-of-spec forces).
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

  // Filter log data matching base model / part
  var rowsToDisplay = [];
  var modelColIdx = logCols.BASE_MODEL - 1;
  for (var r = 1; r < logData.length; r++) {
    var row = logData[r];
    if (String(row[modelColIdx] || "").trim().toLowerCase() === cleanPart) {
      rowsToDisplay.push(row);
    }
  }

  if (rowsToDisplay.length === 0) return;

  var startRow = ranges.RESULTS_START_ROW;
  var numRows = rowsToDisplay.length;
  var numCols = Math.min(rowsToDisplay[0].length, ranges.RESULTS_COL_COUNT);
  var outputRange = sheet.getRange(startRow, ranges.RESULTS_START_COL, numRows, numCols);

  // Truncate row array length to match expected table column width
  var trimmedRows = rowsToDisplay.map(function(row) {
    return row.slice(0, numCols);
  });

  // 1. Write Data Values
  outputRange.setValues(trimmedRows);

  // 2. Build Formatting Arrays
  var bgColors = [];
  var fontColors = [];

  var c1Idx = logCols.COMP_1 - 1;
  var r1Idx = logCols.REB_1 - 1;
  var c2Idx = logCols.COMP_2 - 1;
  var r2Idx = logCols.REB_2 - 1;

  for (var rIdx = 0; rIdx < numRows; rIdx++) {
    var rowBg = [];
    var rowFont = [];
    var rowData = trimmedRows[rIdx];

    for (var cIdx = 0; cIdx < numCols; cIdx++) {
      var val = parseFloat(rowData[cIdx]);
      var isOut = false;

      // Validate forces against matrix min/max bounds
      if (limits && !isNaN(val)) {
        if (cIdx === c1Idx && limits[refCols.COMP_1_MIN - 1] !== "" && (val < limits[refCols.COMP_1_MIN - 1] || val > limits[refCols.COMP_1_MAX - 1])) isOut = true;
        if (cIdx === r1Idx && limits[refCols.REB_1_MIN - 1] !== "" && (val < limits[refCols.REB_1_MIN - 1] || val > limits[refCols.REB_1_MAX - 1])) isOut = true;
        if (cIdx === c2Idx && limits[refCols.COMP_2_MIN - 1] !== "" && (val < limits[refCols.COMP_2_MIN - 1] || val > limits[refCols.COMP_2_MAX - 1])) isOut = true;
        if (cIdx === r2Idx && limits[refCols.REB_2_MIN - 1] !== "" && (val < limits[refCols.REB_2_MIN - 1] || val > limits[refCols.REB_2_MAX - 1])) isOut = true;
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

  // 3. Batch apply formats in a single API call
  outputRange.setBackgrounds(bgColors).setFontColors(fontColors);
}