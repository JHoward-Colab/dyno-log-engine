// =========================================================================
// 🖥️ USER INTERFACE & CONTROLLERS (UI.gs)
// Workspace Rendering, Button Actions & Triggers
// =========================================================================

function clickMasterSyncButton() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try { processDynoFiles(); } catch(e) { Logger.log("Watch folder alert: " + e.toString()); }
  try { retroactiveLogRecalculate(); } catch(e) { Logger.log("Reference Matrix Recalculation Alert: " + e.toString()); }
  try {
    var sheet = ss.getSheetByName("Operator_Station");
    if (sheet) {
      manageOperatorStation({ source: ss, range: sheet.getRange("C2") });
    }
  } catch(e) { Logger.log("Console screen alert: " + e.toString()); }
}

function manageOperatorStation(e) {
  var ss = e ? e.source : SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName("Operator_Station"); if (!sheet) return;
  var range = e ? e.range : sheet.getRange("C2"); var barcode = String(sheet.getRange("C2").getValue()).trim();
  var verifiedFileIdStr = "";
  
  if (range.getA1Notation() === "C2") {
    sheet.getRange("C3:C6").clearContent(); 
    sheet.getRange("C14:C18").clearContent();
    sheet.getRange("E14:E18").clearContent();
    sheet.getRange("C16").clearContent(); 
    sheet.getRange("A27:L100").clearContent(); 
    sheet.getRange("A27:L100").setBackground("").setFontColor("").setFontWeight("normal"); 
    if (barcode === "" || barcode === "undefined" || barcode === null) return;
    
    var searchBarcode = barcode;
    if (barcode.includes("-")) { searchBarcode = barcode.split("-")[0].trim(); } 
    else if (barcode.includes("_")) { searchBarcode = barcode.split("_")[0].trim(); }
    if (!isNaN(searchBarcode) && searchBarcode.length === 4) { searchBarcode = "00" + searchBarcode; }
    
    var searchCriteria = "title contains '" + searchBarcode + "' and (mimeType = 'application/vnd.google-apps.spreadsheet' or mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') and trashed = false";
    var files = DriveApp.searchFiles(searchCriteria);
    if (!files.hasNext()) { sheet.getRange("C3").setValue("❌ Work Order File Not Found: " + searchBarcode); return; }
    var file = files.next(); var fileUrl = file.getUrl(); verifiedFileIdStr = file.getId(); var realFileName = file.getName(); 
    
    sheet.getRange("C3").setFormula('=HYPERLINK("' + fileUrl + '", "🔗 Open ' + realFileName + '")'); 
    sheet.getRange("Z1").setValue(verifiedFileIdStr);
    
    try {
      var woSpreadsheet = SpreadsheetApp.openById(verifiedFileIdStr); var woSheet = woSpreadsheet.getSheets()[0]; 
      var woPartNumber = woSheet.getRange("D3").getValue(); var woBomRevision = woSheet.getRange("D4").getValue();  
      sheet.getRange("C4").setValue(woBomRevision); sheet.getRange("C5").setValue(woPartNumber); 
      
      var registrySheet = ss.getSheetByName("Program_Registry");
      if (registrySheet && woPartNumber) {
        var regValues = registrySheet.getDataRange().getValues();
        var cleanWoPart = String(woPartNumber).trim().toLowerCase().replace(/[-_\s]/g, "");
        for (var rR = 1; rR < regValues.length; rR++) {
          var regPartClean = String(regValues[rR][2] || "").trim().toLowerCase().replace(/[-_\s]/g, "");
          var regProgName = String(regValues[rR][0] || "").trim();
          var custAccount = String(regValues[rR][3] || "").trim();
          var vehicleSpec = String(regValues[rR][4] || "").trim();
          var valvingSpec = String(regValues[rR][6] || "").trim();
          var clickerTargets = String(regValues[rR][7] || "").trim();

          if (regPartClean === cleanWoPart && regProgName) {
            sheet.getRange("C16").setValue(regProgName);
            sheet.getRange("E16").setValue(regProgName);

            sheet.getRange("C14").setValue(custAccount); sheet.getRange("E14").setValue(custAccount);
            sheet.getRange("C15").setValue(vehicleSpec); sheet.getRange("E15").setValue(vehicleSpec);
            sheet.getRange("C17").setValue(valvingSpec); sheet.getRange("E17").setValue(valvingSpec);
            sheet.getRange("C18").setValue(clickerTargets); sheet.getRange("E18").setValue(clickerTargets);
            break;
          }
        }
      }
      SpreadsheetApp.flush(); 
    } catch(err) { sheet.getRange("C3").setValue("⚠️ Extraction Error: " + err.toString()); return; }
  } else { verifiedFileIdStr = String(sheet.getRange("Z1").getValue()).trim(); }
  
  try {
    var activeProgName = String(sheet.getRange("C16").getValue() || "").trim();
    if (activeProgName.length > 20) {
      activeProgName = activeProgName.substring(0, 14) + "_V" + activeProgName.slice(-1);
    }
    
    var masterLogSheet = ss.getSheetByName("Master_Dyno_Log"); var logData = masterLogSheet.getDataRange().getValues();
    var hMap = buildHeaderMap(logData[0]);
    var refSheet = ss.getSheetByName("Part_Reference_Matrix"); if (!refSheet) return;
    var refData = refSheet.getDataRange().getValues();
    var mMap = buildMatrixHeaderMap(refData[0]);
    
    var databaseSerialMap = {};
    for (var r = 1; r < logData.length; r++) {
      var sKey = String(logData[r][hMap.trueSerial] || "").replace(/[-_\s]/g, "").toLowerCase();
      if (sKey) databaseSerialMap[sKey] = { data: logData[r], rowIdx: r + 1 };
    }
    
    var matrixRowIdx = -1;
    var cleanActiveProg = activeProgName.replace(/[-_\s]/g, "").toLowerCase();
    for (var mx = 0; mx < refData.length; mx++) { 
      var cleanMatrixKey = String(refData[mx][0] || "").trim().replace(/[-_\s]/g, "").toLowerCase();
      if (cleanMatrixKey === cleanActiveProg || (cleanMatrixKey && cleanActiveProg.indexOf(cleanMatrixKey) !== -1)) { 
        matrixRowIdx = mx; break; 
      } 
    }
    
    var c1Min = 0, c1Max = 0, r1Min = 0, r1Max = 0;
    var c2Min = 0, c2Max = 0, r2Min = 0, r2Max = 0;

    if (matrixRowIdx !== -1) {
      var row = refData[matrixRowIdx];
      
      c1Min = parseFloat(row[mMap.c1Min]) || 0;
      c1Max = parseFloat(row[mMap.c1Max]) || 0;
      c2Min = parseFloat(row[mMap.c2Min]) || 0;
      c2Max = parseFloat(row[mMap.c2Max]) || 0;

      var r1RawA = Math.abs(parseFloat(row[mMap.r1Min]) || 0);
      var r1RawB = Math.abs(parseFloat(row[mMap.r1Max]) || 0);
      r1Min = Math.min(r1RawA, r1RawB);
      r1Max = Math.max(r1RawA, r1RawB);

      var r2RawA = Math.abs(parseFloat(row[mMap.r2Min]) || 0);
      var r2RawB = Math.abs(parseFloat(row[mMap.r2Max]) || 0);
      r2Min = Math.min(r2RawA, r2RawB);
      r2Max = Math.max(r2RawA, r2RawB);

      sheet.getRange("B22").setValue(c1Min); sheet.getRange("C22").setValue(c1Max);
      sheet.getRange("D22").setValue(r1Min); sheet.getRange("E22").setValue(r1Max);
      sheet.getRange("B23").setValue(c2Min); sheet.getRange("C23").setValue(c2Max);
      sheet.getRange("D23").setValue(r2Min); sheet.getRange("E23").setValue(r2Max);

      var rawSlopeVal = parseFloat(row[mMap.slope1Min]);
      sheet.getRange("F22").setValue(!isNaN(rawSlopeVal) ? rawSlopeVal.toFixed(1) : "No Limit");
      sheet.getRange("B22:E23").setNumberFormat("0.0");
    }
    
    if (verifiedFileIdStr !== "") {
      var activeWoSpreadsheet = SpreadsheetApp.openById(verifiedFileIdStr); var activeWoSheet = activeWoSpreadsheet.getSheets()[0]; var totalWoRows = activeWoSheet.getLastRow();
      if (totalWoRows >= 12) {
        var serials = activeWoSheet.getRange(12, 1, totalWoRows - 11, 1).getValues(); var outputGrid = [];
        var backgroundColorsGrid = []; var fontColorsGrid = []; var fontWeightsGrid = [];
        
        var resolvedRows = [];
        for (var i = 0; i < serials.length; i++) {
          var snRaw = String(serials[i][0] || "").trim();
          if (snRaw === "") { resolvedRows.push(null); continue; }
          
          var cleanSnRaw = snRaw.replace(/[-_\s]/g, "").toLowerCase();
          var cleanSnStripped = cleanSnRaw.split("-")[0].split("_")[0].replace(/^0+/, "");
          var matchedRow = databaseSerialMap[cleanSnRaw];
          
          if (!matchedRow) {
            for (var sKey in databaseSerialMap) {
              var sKeyStripped = sKey.split("-")[0].split("_")[0].replace(/^0+/, "");
              if (sKeyStripped.endsWith(cleanSnStripped) || cleanSnStripped.endsWith(sKeyStripped)) {
                var rawSuffix = cleanSnRaw.split("-")[1] || cleanSnRaw.slice(-3);
                var lookupSuffix = sKey.split("-")[1] || sKey.slice(-3);
                if (rawSuffix === lookupSuffix) {
                  matchedRow = databaseSerialMap[sKey]; break;
                }
              }
            }
          }
          resolvedRows.push(matchedRow);
        }

        for (var i = 0; i < serials.length; i++) {
          var snRaw = String(serials[i][0] || "").trim();
          if (snRaw !== "") {
            var matchedRow = resolvedRows[i];
            
            var rowBackgrounds = ["#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF"];
            var rowFonts = ["#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000"];
            var rowWeights = ["normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal", "normal"];
            
            var pure9Digits = snRaw;
            var patternMatch = snRaw.match(/(\d{6})[-_](\d{3})$/) || snRaw.match(/(\d{6})_(\d{3})$/) || snRaw.match(/(\d{6})(\d{3})$/);
            if (patternMatch) pure9Digits = patternMatch[1] + "-" + patternMatch[2];
            
            if (matchedRow) {
              var sRow = matchedRow.data;
              var logSheetRowIndexNumber = matchedRow.rowIdx;
              
              var individualRF = parseFloat(sRow[hMap.rodForce]) || 0;
              var lowC = parseFloat(sRow[hMap.comp1]) || 0;
              var lowR = Math.abs(parseFloat(sRow[hMap.reb1])) || 0;
              var midC = parseFloat(sRow[hMap.comp2]) || 0;
              var midR = Math.abs(parseFloat(sRow[hMap.reb2])) || 0; 
              
              var test1Result = String(sRow[21] || "INITIALIZING"); 
              var test2ResultString = String(sRow[22] || "INITIALIZING"); 
              var totalStatus = String(sRow[23] || "INITIALIZING"); 
              var diagnosticNotes = String(sRow[24] || "");        
              var logSheetTeardownStatus = String(sRow[25] || ""); 
              var rawEngCommentText = String(sRow[26] || "");       

              var extNotesLower = logSheetTeardownStatus.toLowerCase();
              var skipHighlight = extNotesLower.includes("approved") || extNotesLower.includes("management") || extNotesLower.includes("no issue found");
              
              var applyFaultHighlight = function(colIndex) {
                rowFonts[colIndex] = "#C5221F";      
                rowWeights[colIndex] = "bold";        
                rowBackgrounds[colIndex] = "#FCE8E6"; 
              };

              // 🔒 STRICT BLUEPRINT TARGET CELL HIGHLIGHTING (TEST 1 HARD LIMITS ONLY)
              if (!skipHighlight) {
                if (c1Min > 0 && c1Max > 0 && (lowC < c1Min || lowC > c1Max)) applyFaultHighlight(2); // Low Speed Comp (Col C)
                if (r1Min > 0 && r1Max > 0 && (lowR < r1Min || lowR > r1Max)) applyFaultHighlight(3); // Low Speed Reb (Col D)
                if (c2Min > 0 && c2Max > 0 && (midC < c2Min || midC > c2Max)) applyFaultHighlight(4); // Med Speed Comp (Col E)
                if (r2Min > 0 && r2Max > 0 && (midR < r2Min || midR > r2Max)) applyFaultHighlight(5); // Med Speed Reb (Col F)
              }
              
              if (test1Result.includes("FAIL")) {
                rowBackgrounds[6] = "#FCE8E6"; rowFonts[6] = "#C5221F"; rowWeights[6] = "bold";
              } else if (test1Result.includes("PASS")) {
                rowBackgrounds[6] = "#E6F4EA"; rowFonts[6] = "#137333";
              }

              if (test2ResultString.includes("FAIL")) {
                rowBackgrounds[7] = "#FCE8E6"; rowFonts[7] = "#C5221F"; rowWeights[7] = "bold";
              } else if (test2ResultString.includes("PASS")) {
                rowBackgrounds[7] = "#E6F4EA"; rowFonts[7] = "#137333";
              }
              
              if (totalStatus.includes("FAIL")) { 
                rowBackgrounds[8] = "#C5221F"; rowFonts[8] = "#FFFFFF"; rowWeights[8] = "bold"; 
              } else if (totalStatus.includes("OVERRIDE")) { 
                rowBackgrounds[8] = "#E6F4EA"; rowFonts[8] = "#137333"; rowWeights[8] = "bold"; 
              } else if (totalStatus.includes("PASS")) {
                rowBackgrounds[8] = "#E6F4EA"; rowFonts[8] = "#137333";
              }
              
              var sheetId = masterLogSheet.getSheetId();
              var linkFormula = '=HYPERLINK("' + ss.getUrl() + '#gid=' + sheetId + '&range=X' + logSheetRowIndexNumber + '", "' + pure9Digits + '")';
              
              outputGrid.push([ 
                linkFormula, 
                individualRF.toFixed(1), 
                lowC.toFixed(1), 
                lowR.toFixed(1), 
                midC.toFixed(1), 
                midR.toFixed(1), 
                test1Result, 
                test2ResultString, 
                totalStatus, 
                logSheetTeardownStatus, 
                diagnosticNotes, 
                rawEngCommentText 
              ]);
            } else { 
              outputGrid.push([pure9Digits, "", "", "", "", "", "NOT RUN", "NOT RUN", "NOT TESTED YET", "", "", ""]); 
            }
            backgroundColorsGrid.push(rowBackgrounds); fontColorsGrid.push(rowFonts); fontWeightsGrid.push(rowWeights);
          }
        }
        
        if (outputGrid.length > 0) {
          var targetRange = sheet.getRange(27, 1, outputGrid.length, 12);
          targetRange.setValues(outputGrid); 
          targetRange.setBackgrounds(backgroundColorsGrid); 
          targetRange.setFontColors(fontColorsGrid);
          targetRange.setFontWeights(fontWeightsGrid);
          
          sheet.getRange(27, 2, outputGrid.length, 5).setNumberFormat("0.0");
          
          var dynamicLastRow = 27 + outputGrid.length - 1;
          var statusBarFormula = '=IF(OR(COUNTIF(I27:I' + dynamicLastRow + ', "*FAIL*") > 0, COUNTIF(G27:H' + dynamicLastRow + ', "*FAIL*") > 0), "🔴 WORK ORDER BLOCKED: FAILED UNITS REQUIRING EVALUATION", "🟢 WORK ORDER COMPLETED & PASSING")';
          sheet.getRange("A8").setFormula(statusBarFormula);
          sheet.getRange("A11").setFormula(statusBarFormula);
        }
      }
    }
  } catch(err) { sheet.getRange("C3").setValue("⚠️ Sync Error: " + err.toString()); }
}

function installableConsoleTrigger(e) {
  if (e && e.range.getSheet().getName() === "Operator_Station" && e.range.getA1Notation() === "C2") manageOperatorStation(e);
}

function installableOnEdit(e) {
  if (e && e.range.getSheet().getName() === "Operator_Station" && e.range.getA1Notation() === "C2") manageOperatorStation({ source: e.source, range: e.range });
}