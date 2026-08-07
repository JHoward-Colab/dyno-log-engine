// =========================================================================
// INTERFACE CONTROLLERS (UNIFIED MASTER MACRO PIPELINE)
// =========================================================================
function clickMasterSyncButton() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1. Run data ingestion first
  try { processDynoFiles(); } catch(e) { Logger.log("Watch folder alert: " + e.toString()); }
  // 2. Compute Test 1 (Global Specs), Test 2 (Batch Gate), and update Part Reference baselines
  try { retroactiveLogRecalculate(); } catch(e) { Logger.log("Reference Matrix Recalculation Alert: " + e.toString()); }
  // 3. Render matching rows to the operator workspace panel
  try {
    var sheet = ss.getSheetByName("Operator_Station");
    if (sheet) {
      manageOperatorStation({ source: ss, range: sheet.getRange("C2") });
    }
  } catch(e) { Logger.log("Console screen alert: " + e.toString()); }
}

// =========================================================================
// UTILITY ENGINES: GLOBAL MATH, NOMINAL BUCKETING & INDEX RESOLVERS
// =========================================================================
function calculateMean(arr) {
  var clean = arr.filter(function(x) { return !isNaN(x) && x !== null && x !== ""; });
  return clean.length === 0 ? 0 : clean.reduce(function(a, b) { return a + b; }, 0) / clean.length;
}

function calculateSD(arr, m) {
  var clean = arr.filter(function(x) { return !isNaN(x) && x !== null && x !== ""; });
  if (clean.length <= 1) return 0;
  var meanVal = (m !== undefined) ? m : calculateMean(clean);
  var variance = clean.map(function(x) { return Math.pow(x - meanVal, 2); }).reduce(function(a, b) { return a + b; }, 0) / (clean.length - 1);
  return Math.sqrt(variance);
}

function snapToNominalSpeed(val) {
  var v = parseFloat(val);
  if (isNaN(v) || v <= 0) return 0;
  if (v < 250) return 100;
  if (v < 700) return 400;
  if (v < 1800) return 1000;
  return 2500;
}

function buildHeaderMap(headerRow) {
  var map = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Dyno_Log");
  
  if (!headerRow && sheet) {
    headerRow = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 32)).getValues()[0];
  }
  if (!headerRow) return map;
  
  for (var i = 0; i < headerRow.length; i++) {
    var cleanHeader = String(headerRow[i] || "").trim().toLowerCase();
    if (cleanHeader.includes("timestamp")) map.timestamp = i;
    else if (cleanHeader.includes("program name")) map.programName = i;
    else if (cleanHeader.includes("true serial number") || cleanHeader.includes("serial number")) map.trueSerial = i;
    else if (cleanHeader.includes("base model")) map.baseModel = i;
    else if (cleanHeader.includes("valving version")) map.valvingVersion = i;
    else if (cleanHeader.includes("rod force")) map.rodForce = i;
    else if (cleanHeader.includes("comp peak 1")) map.comp1 = i;
    else if (cleanHeader.includes("reb peak 1")) map.reb1 = i;
    else if (cleanHeader.includes("comp peak 2")) map.comp2 = i;
    else if (cleanHeader.includes("reb peak 2")) map.reb2 = i;
    else if (cleanHeader.includes("comp peak 3")) map.comp3 = i;
    else if (cleanHeader.includes("reb peak 3")) map.reb3 = i;
    else if (cleanHeader.includes("crossover slope 1")) map.slope1 = i;
    else if (cleanHeader.includes("loop area 1")) map.loopArea1 = i;
    else if (cleanHeader.includes("loop area 2")) map.loopArea2 = i;
    else if (cleanHeader.includes("speed")) {
      if (cleanHeader.includes("1")) map.speed1 = i;
      else if (cleanHeader.includes("2")) map.speed2 = i;
      else if (cleanHeader.includes("3")) map.speed3 = i;
    }
  }
  
  // 🔒 HARD INDEX SAFEGUARDS
  if (map.timestamp === undefined) map.timestamp = 0;
  if (map.programName === undefined) map.programName = 1;
  if (map.trueSerial === undefined) map.trueSerial = 2;
  if (map.baseModel === undefined) map.baseModel = 3;
  if (map.valvingVersion === undefined) map.valvingVersion = 4;
  if (map.rodForce === undefined) map.rodForce = 5;
  if (map.speed1 === undefined) map.speed1 = 6;
  if (map.comp1 === undefined) map.comp1 = 7;
  if (map.reb1 === undefined) map.reb1 = 8;
  if (map.slope1 === undefined) map.slope1 = 9;
  if (map.loopArea1 === undefined) map.loopArea1 = 10;
  if (map.speed2 === undefined) map.speed2 = 11;
  if (map.comp2 === undefined) map.comp2 = 12;
  if (map.reb2 === undefined) map.reb2 = 13;
  if (map.speed3 === undefined) map.speed3 = 16;
  if (map.comp3 === undefined) map.comp3 = 17;
  if (map.reb3 === undefined) map.reb3 = 18;
  
  map.test1Status = 21;    // Column V
  map.test2Status = 22;    // Column W
  map.status = 23;         // Column X
  map.diagnostics = 24;    // Column Y
  map.teardown = 25;       // Column Z
  map.engComments = 26;    // Column AA
  
  return map;
}

function buildMatrixHeaderMap(headerRow) {
  var map = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Part_Reference_Matrix");
  if (!headerRow && sheet) {
    headerRow = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 30)).getValues()[0];
  }
  
  map.speed1 = 1;  // Col B
  map.speed2 = 2;  // Col C
  map.speed3 = 3;  // Col D
  map.c1Mean = 4;  // Col E
  map.c1SD = 5;    // Col F
  map.c1Min = 6;   // Col G
  map.c1Max = 7;   // Col H
  map.r1Mean = 8;  // Col I
  map.r1SD = 9;    // Col J
  map.r1Min = 10;  // Col K
  map.r1Max = 11;  // Col L
  map.slope1Min = 12; // Col M
  map.loopArea1Min = 13; // Col N
  map.c2Mean = 14; // Col O
  map.c2SD = 15;   // Col P
  map.c2Min = 16;  // Col Q
  map.c2Max = 17;  // Col R
  map.r2Mean = 18; // Col S
  map.r2SD = 19;   // Col T
  map.r2Min = 20;  // Col U
  map.r2Max = 21;  // Col V
  map.healthStamp = 25;  // Col Z
  map.controlMode = 26;  // Col AA
  map.sampleCount = 28;  // Col AC
  
  if (!headerRow) return map;
  
  for (var i = 0; i < headerRow.length; i++) {
    var cleanHeader = String(headerRow[i] || "").trim().toLowerCase();
    if (cleanHeader.includes("speed 1") || cleanHeader.includes("speed1")) map.speed1 = i;
    else if (cleanHeader.includes("speed 2") || cleanHeader.includes("speed2")) map.speed2 = i;
    else if (cleanHeader.includes("speed 3") || cleanHeader.includes("speed3")) map.speed3 = i;
    else if (cleanHeader.includes("process health stamp")) map.healthStamp = i;
    else if (cleanHeader.includes("system control mode")) map.controlMode = i;
    else if (cleanHeader.includes("sample count")) map.sampleCount = i;
    
    if (i < 28) {
      if (cleanHeader.includes("comp 1 min") || cleanHeader.includes("c1 min")) map.c1Min = i;
      else if (cleanHeader.includes("comp 1 max") || cleanHeader.includes("c1 max")) map.c1Max = i;
      else if (cleanHeader.includes("reb 1 min") || cleanHeader.includes("r1 min")) map.r1Min = i;
      else if (cleanHeader.includes("reb 1 max") || cleanHeader.includes("r1 max")) map.r1Max = i;
      else if (cleanHeader.includes("slope 1 min")) map.slope1Min = i;
      else if (cleanHeader.includes("comp 2 min") || cleanHeader.includes("c2 min")) map.c2Min = i;
      else if (cleanHeader.includes("comp 2 max") || cleanHeader.includes("c2 max")) map.c2Max = i;
      else if (cleanHeader.includes("reb 2 min") || cleanHeader.includes("r2 min")) map.r2Min = i;
      else if (cleanHeader.includes("reb 2 max") || cleanHeader.includes("r2 max")) map.r2Max = i;
    }
  }
  return map;
}

// =========================================================================
// ENGINE 1: ADAPTIVE DYNO PROCESSOR (CLEAN STACK APPEND)
// =========================================================================
function processDynoFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Dyno_Log");
  if (!sheet) return;
  
  var folderName = "01_Watch_Folder"; 
  var folders = DriveApp.getFoldersByName(folderName);
  var folder;
  if (folders.hasNext()) { folder = folders.next(); } else { return; }
  
  var files = folder.getFiles();
  var tempFileMap = {};
  var now = new Date();
  
  while (files.hasNext()) {
    var file = files.next(); var rawName = file.getName(); var lastUpdated = file.getLastUpdated().getTime();
    if (rawName.toLowerCase().endsWith('.csv')) {
      var fileName = rawName.replace(/\.csv$/i, "").replace(/\.txt$/i, "").trim(); var tempGroupKey = "";
      var serialMatch = rawName.match(/\d{6}-\d{3}/) || rawName.match(/\d+-\d+/);
      if (serialMatch) { tempGroupKey = serialMatch[0].trim(); } else {
        var clearParts = fileName.split(/[\s_]+/);
        tempGroupKey = clearParts.length >= 3 ? clearParts[clearParts.length - 2].trim() : clearParts[0].trim();
      }
      
      var detectedModel = "";
      if (fileName.toLowerCase().includes("pvp")) {
        detectedModel = fileName.split("pvp")[0].replace(/[^a-zA-Z0-9_]/g, " ").trim().split(" ")[0];
      } else {
        var spaceParts = fileName.split(" ");
        if (spaceParts[0].includes("_V")) {
          detectedModel = spaceParts[0].trim();
        } else if (rawName.includes(tempGroupKey)) {
          var preSerialText = rawName.split(tempGroupKey)[0].trim();
          if (preSerialText.endsWith('_') || preSerialText.endsWith(' ')) { preSerialText = preSerialText.slice(0, -1).trim(); }
          detectedModel = preSerialText;
        }
      }
      if (detectedModel === "" || detectedModel.startsWith('_PROCESSING_')) {
        var rawParts = fileName.split(/[\s_]+/); detectedModel = rawParts[0].replace('_PROCESSING_', '');
        if (rawParts[1] && rawParts[1].toUpperCase().startsWith('V')) { detectedModel += "_" + rawParts[1]; }
      }
      
      var speedSuffix = "";
      if (fileName.toLowerCase().includes("pvp")) {
        speedSuffix = "pvp";
      } else {
        var nameParts = fileName.split(/[\s_]+/); 
        speedSuffix = nameParts[nameParts.length - 1].toLowerCase().trim();
        if (speedSuffix === "interval" && nameParts.length >= 2) {
          speedSuffix = nameParts[nameParts.length - 2].toLowerCase().trim();
        }
      }
      
      if (speedSuffix === 'pvp' || !isNaN(parseInt(speedSuffix))) {
        if (!tempFileMap[tempGroupKey]) { tempFileMap[tempGroupKey] = { pvp: null, intervals: [], youngestFileTime: 0, parsedBaseModel: detectedModel }; }
        if (lastUpdated > tempFileMap[tempGroupKey].youngestFileTime) tempFileMap[tempGroupKey].youngestFileTime = lastUpdated;
        if (speedSuffix === 'pvp') { tempFileMap[tempGroupKey].pvp = file; } else { tempFileMap[tempGroupKey].intervals.push(file); }
      }
    }
  }
  
  var registrySheet = ss.getSheetByName("Program_Registry"); var registryData = registrySheet ? registrySheet.getDataRange().getValues() : [];
  var masterLogData = sheet.getDataRange().getValues();
  var hMap = buildHeaderMap(masterLogData[0]);
  
  for (var tempKey in tempFileMap) {
    var pack = tempFileMap[tempKey];
    if (pack.pvp && pack.intervals.length >= 3) {
      if ((now.getTime() - pack.youngestFileTime) / 1000 < 4) continue;
      try {
        var rawBlobStr = pack.pvp.getBlob().getDataAsString("UTF-8").trim();
        var pvpContent = Utilities.parseCsv(rawBlobStr); var trueSerial = tempKey;
        var trueDynoProgramName = pack.parsedBaseModel;
        
        if (trueDynoProgramName.length > 20) {
          trueDynoProgramName = trueDynoProgramName.substring(0, 14) + "_V" + trueDynoProgramName.slice(-1);
        }
        
        var cleanTrueProgName = trueDynoProgramName.replace(/[-_\s]/g, "").toLowerCase();
        for (var rR = 1; rR < registryData.length; rR++) {
          var regKey = String(registryData[rR][0] || "").trim().replace(/[-_\s]/g, "").toLowerCase();
          if (regKey === cleanTrueProgName) { 
            trueDynoProgramName = String(registryData[rR][0] || "").trim(); break; 
          }
        }
        
        var registryBaseModelText = trueDynoProgramName; var registryValvingVersion = "PRODUCTION_RUN";
        var cleanMatchName = trueDynoProgramName.replace(/[-_\s]/g, "").toLowerCase();
        for (var regR = 1; regR < registryData.length; regR++) {
          var cleanRegCell = String(registryData[regR][0] || "").trim().replace(/[-_\s]/g, "").toLowerCase();
          if (cleanRegCell === cleanMatchName) {
            registryBaseModelText = String(registryData[regR][2] || "").trim();
            registryValvingVersion = String(registryData[regR][6] || "").trim(); 
            break;
          }
        }
        
        var rawRodForce = 0; var intervalMetrics = [];
        for (var i = 0; i < pack.intervals.length; i++) {
          var file = pack.intervals[i]; var rawIntStr = file.getBlob().getDataAsString("UTF-8").trim();
          if (rawIntStr === "") continue; var rows = Utilities.parseCsv(rawIntStr); var speedTarget = 0;
          for (var j = 0; j < Math.min(rows.length, 15); j++) { if (rows[j][0] && rows[j][0].trim().toLowerCase() == "velocity amplitude") { speedTarget = parseInt(rows[j][1]); break; } }
          
          if (!speedTarget || isNaN(speedTarget)) {
            for (var r = 14; r < rows.length; r++) {
              if (rows[r][0] && !isNaN(parseFloat(rows[r][0]))) {
                var testV = Math.abs(parseFloat(rows[r][0]));
                if (testV > speedTarget) speedTarget = testV;
              }
            }
          }
          speedTarget = snapToNominalSpeed(speedTarget);
          
          for (var j = 0; j < rows.length; j++) { if (rows[j][0] && rows[j][0].trim().toLowerCase() == "rod force") { rawRodForce = parseFloat(rows[j][1]); break; } }
          var maxComp = -9999, maxReb = 9999, totalArea = 0, nearZeroPoints = [];
          for (var r = 14; r < rows.length; r++) {
            if (rows[r].length < 2 || isNaN(parseFloat(rows[r][0]))) continue;
            var vel = parseFloat(rows[r][0]); var force = parseFloat(rows[r][1]);
            if (force > maxComp) maxComp = force; if (force < maxReb) maxReb = force;
            totalArea += Math.abs(force); if (Math.abs(vel) < ((speedTarget || 500) * 0.15)) nearZeroPoints.push({ x: vel, y: force });
          }
          var sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
          for (var p = 0; p < nearZeroPoints.length; p++) { sumX += nearZeroPoints[p].x; sumY += nearZeroPoints[p].y; sumXY += (nearZeroPoints[p].x * nearZeroPoints[p].y); var varXX = (nearZeroPoints[p].x * nearZeroPoints[p].x); sumXX += varXX; }
          var denom = (nearZeroPoints.length * sumXX) - (sumX * sumX);
          var slope = denom == 0 ? 0 : ((nearZeroPoints.length * sumXY) - (sumX * sumY)) / denom;
          intervalMetrics.push({ speed: speedTarget, maxComp: maxComp, maxReb: maxReb, slope: slope, area: totalArea });
        }
        intervalMetrics.sort(function(a, b) { return a.speed - b.speed; });
        
        if (intervalMetrics.length >= 3) {
          var lowSp = intervalMetrics[0].speed;
          var highSp = intervalMetrics[2].speed;
          if (lowSp <= 100 || highSp <= 1000) {
            intervalMetrics[0].speed = 100;
            intervalMetrics[1].speed = 400;
            intervalMetrics[2].speed = 1000;
          } else {
            intervalMetrics[0].speed = 500;
            intervalMetrics[1].speed = 1000;
            intervalMetrics[2].speed = 2500;
          }
        }
        
        var pvpSlots = [{ comp: 0, reb: 0 }, { comp: 0, reb: 0 }, { comp: 0, reb: 0 }];
        if (pvpContent.length >= 11) {
          pvpSlots[0].comp = parseFloat(pvpContent[8][3]) || 0; pvpSlots[0].reb = parseFloat(pvpContent[8][5]) || 0;
          pvpSlots[1].comp = parseFloat(pvpContent[9][3]) || 0; pvpSlots[1].reb = parseFloat(pvpContent[9][5]) || 0;
          pvpSlots[2].comp = parseFloat(pvpContent[10][3]) || 0; pvpSlots[2].reb = parseFloat(pvpContent[10][5]) || 0;
        }
        
        var outputRowArray = [];
        for (var c = 0; c < masterLogData[0].length; c++) { outputRowArray.push(""); }
        
        if (hMap.timestamp !== undefined) outputRowArray[hMap.timestamp] = new Date();
        if (hMap.programName !== undefined) outputRowArray[hMap.programName] = trueDynoProgramName;
        if (hMap.trueSerial !== undefined) outputRowArray[hMap.trueSerial] = trueSerial;
        if (hMap.baseModel !== undefined) outputRowArray[hMap.baseModel] = registryBaseModelText; 
        if (hMap.valvingVersion !== undefined) outputRowArray[hMap.valvingVersion] = registryValvingVersion; 
        if (hMap.rodForce !== undefined) outputRowArray[hMap.rodForce] = rawRodForce;
        
        if (intervalMetrics[0]) { 
          if (hMap.speed1 !== undefined) outputRowArray[hMap.speed1] = intervalMetrics[0].speed;
          if (hMap.comp1 !== undefined) outputRowArray[hMap.comp1] = pvpSlots[0].comp; 
          if (hMap.reb1 !== undefined) outputRowArray[hMap.reb1] = pvpSlots[0].reb; 
          if (hMap.slope1 !== undefined) outputRowArray[hMap.slope1] = intervalMetrics[0].slope; 
          if (hMap.loopArea1 !== undefined) outputRowArray[hMap.loopArea1] = intervalMetrics[0].area; 
        }
        if (intervalMetrics[1]) { 
          if (hMap.speed2 !== undefined) outputRowArray[hMap.speed2] = intervalMetrics[1].speed;
          if (hMap.comp2 !== undefined) outputRowArray[hMap.comp2] = pvpSlots[1].comp; 
          if (hMap.reb2 !== undefined) outputRowArray[hMap.reb2] = pvpSlots[1].reb; 
          if (hMap.loopArea2 !== undefined) outputRowArray[hMap.loopArea2] = intervalMetrics[1].area; 
        }
        if (intervalMetrics[2]) { 
          if (hMap.speed3 !== undefined) outputRowArray[hMap.speed3] = intervalMetrics[2].speed;
          if (hMap.comp3 !== undefined) outputRowArray[hMap.comp3] = pvpSlots[2].comp; 
          if (hMap.reb3 !== undefined) outputRowArray[hMap.reb3] = pvpSlots[2].reb; 
        }
        if (hMap.status !== undefined) outputRowArray[hMap.status] = "PASS";
        
        var dynamicNextRowIndex = sheet.getLastRow() + 1;
        sheet.getRange(dynamicNextRowIndex, 1, 1, outputRowArray.length).setValues([outputRowArray]);
        
        var archiveFolder; try { archiveFolder = folder.getFoldersByName("02_Archive").next(); } catch(e) { archiveFolder = folder.createFolder("02_Archive"); }
        pack.pvp.moveTo(archiveFolder); for (var f = 0; f < pack.intervals.length; f++) { pack.intervals[f].moveTo(archiveFolder); }
      } catch(err) { Logger.log("Watch folder processing failed: " + err.toString()); }
    }
  }
}

// =========================================================================
// HISTORICAL RE-RUN ENGINE (STRICT SELF-HEALING UNIFIED SPC BALANCER)
// =========================================================================
function retroactiveLogRecalculate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Master_Dyno_Log");
  var refSheet = ss.getSheetByName("Part_Reference_Matrix");
  var regSheet = ss.getSheetByName("Program_Registry");
  if (!logSheet || !refSheet) return;
  
  var logRange = logSheet.getDataRange();
  var logData = logRange.getValues(); if (logData.length < 2) return;
  var hMap = buildHeaderMap(logData[0]);
  var refData = refSheet.getDataRange().getValues();
  var mMap = buildMatrixHeaderMap(refData[0]);
  
  var mean = function(arr) { var clean = arr.filter(function(x){return !isNaN(x);}); return clean.length === 0 ? 0 : clean.reduce(function(a,b){return a+b;},0)/clean.length; };
  var sd = function(arr, m) { var clean = arr.filter(function(x){return !isNaN(x);}); return clean.length <= 1 ? 0 : Math.sqrt(clean.map(function(x){return Math.pow(x-m,2);}).reduce(function(a,b){return a+b;},0)/(clean.length - 1)); };

  var batchGroups = {};
  var historicalGroups = {};

  // Build PROGRAM_NAME / BASE_MODEL -> DYNAMIC_KEY mapping from Program_Registry
  var modelToDynamicKey = {};
  if (regSheet) {
    var regValues = regSheet.getDataRange().getValues();
    for (var k = 1; k < regValues.length; k++) {
      var regProgName = String(regValues[k][0] || "").trim().toLowerCase().replace(/[-_\s]/g, ""); // Col A
      var regDynKey   = String(regValues[k][1] || "").trim();                                     // Col B
      var regBaseModel= String(regValues[k][2] || "").trim().toLowerCase().replace(/[-_\s]/g, ""); // Col C
      
      var targetKey = regDynKey || regValues[k][0] || regValues[k][2];
      if (regProgName) modelToDynamicKey[regProgName] = targetKey;
      if (regBaseModel) modelToDynamicKey[regBaseModel] = targetKey;
    }
  }
  
  // Pass 1: Cohort Batch & Historical Program Grouping
  for (var r = 1; r < logData.length; r++) {
    var prog = String(logData[r][hMap.programName] || "").trim();
    var serial = String(logData[r][hMap.trueSerial] || "").trim();
    var baseModel = String(logData[r][hMap.baseModel] || "").trim().toLowerCase().replace(/[-_\s]/g, "");
    var cleanProg = prog.toLowerCase().replace(/[-_\s]/g, "");
    if (logData[r][0] === "") continue;
    
    if (serial !== "") {
      var batchId = serial.split("-")[0].trim();
      if (!batchGroups[batchId]) {
        batchGroups[batchId] = { c1: [], r1: [], c2: [], r2: [], rf: [], rowReferences: [] };
      }
      batchGroups[batchId].c1.push(parseFloat(logData[r][hMap.comp1]) || 0);
      batchGroups[batchId].r1.push(Math.abs(parseFloat(logData[r][hMap.reb1])) || 0);
      batchGroups[batchId].c2.push(parseFloat(logData[r][hMap.comp2]) || 0);
      batchGroups[batchId].r2.push(Math.abs(parseFloat(logData[r][hMap.reb2])) || 0);
      batchGroups[batchId].rf.push(parseFloat(logData[r][hMap.rodForce]) || 0);
      batchGroups[batchId].rowReferences.push(r);
    }
    
    // Group historical logs using mapped Dynamic Key fallback
    var resolvedGroupKey = modelToDynamicKey[cleanProg] || modelToDynamicKey[baseModel] || prog;
    if (resolvedGroupKey !== "") {
      if (!historicalGroups[resolvedGroupKey]) historicalGroups[resolvedGroupKey] = [];
      logData[r]._rowIdx = r + 1; 
      historicalGroups[resolvedGroupKey].push(logData[r]);
    }
  }
  
  // Pass 2: Matrix Baseline & Speed Bucketing
  for (var pName in historicalGroups) {
    var pool = historicalGroups[pName];
    var countN = pool.length;
    var cleanPName = pName.replace(/[-_\s]/g, "").toLowerCase();
    
    var refRowIdx = -1;
    for (var mx = 1; mx < refData.length; mx++) {
      var matrixKeyRaw = String(refData[mx][0] || "").trim();
      if (!matrixKeyRaw) continue;
      var cleanMatrixKey = matrixKeyRaw.replace(/[-_\s]/g, "").toLowerCase();
      
      if (cleanMatrixKey === cleanPName || 
         (cleanMatrixKey.length >= 6 && cleanPName.indexOf(cleanMatrixKey) !== -1) || 
         (cleanPName.length >= 6 && cleanMatrixKey.indexOf(cleanPName) !== -1)) {
        refRowIdx = mx + 1;
        break;
      }
    }
    
    if (refRowIdx !== -1) {
      var row = refData[refRowIdx - 1];
      var c1MinRaw = row[mMap.c1Min];
      var isSeeded = c1MinRaw !== "" && c1MinRaw !== null && !isNaN(parseFloat(c1MinRaw));
      
      var c1Vals = [], r1Vals = [], c2Vals = [], r2Vals = [];
      var s1Vals = [], s2Vals = [], s3Vals = [];
      
      for (var s = 0; s < pool.length; s++) {
        var d = pool[s];
        c1Vals.push(parseFloat(d[hMap.comp1]) || 0);
        r1Vals.push(Math.abs(parseFloat(d[hMap.reb1])) || 0);
        c2Vals.push(parseFloat(d[hMap.comp2]) || 0);
        r2Vals.push(Math.abs(parseFloat(d[hMap.reb2])) || 0);
        
        var sp1 = snapToNominalSpeed(d[hMap.speed1]);
        var sp2 = snapToNominalSpeed(d[hMap.speed2]);
        var sp3 = snapToNominalSpeed(d[hMap.speed3]);
        
        if (sp1 > 0) s1Vals.push(sp1);
        if (sp2 > 0) s2Vals.push(sp2);
        if (sp3 > 0) s3Vals.push(sp3);
      }
      
      var rawSp1 = s1Vals.length > 0 ? mean(s1Vals) : 100;
      var rawSp2 = s2Vals.length > 0 ? mean(s2Vals) : 400;
      var rawSp3 = s3Vals.length > 0 ? mean(s3Vals) : 1000;
      
      var snappedSp1 = snapToNominalSpeed(rawSp1);
      var snappedSp2 = snapToNominalSpeed(rawSp2);
      var snappedSp3 = snapToNominalSpeed(rawSp3);
      
      if (snappedSp1 <= 100 || snappedSp3 <= 1000) {
        snappedSp1 = 100; snappedSp2 = 400; snappedSp3 = 1000;
      } else {
        snappedSp1 = 500; snappedSp2 = 1000; snappedSp3 = 2500;
      }
      
      if (mMap.speed1 !== undefined) refSheet.getRange(refRowIdx, mMap.speed1 + 1).setValue(snappedSp1);
      if (mMap.speed2 !== undefined) refSheet.getRange(refRowIdx, mMap.speed2 + 1).setValue(snappedSp2);
      if (mMap.speed3 !== undefined) refSheet.getRange(refRowIdx, mMap.speed3 + 1).setValue(snappedSp3);
      
      for (var s = 0; s < pool.length; s++) {
        var d = pool[s];
        var rIdx = d._rowIdx;
        if (rIdx) {
          if (d[hMap.speed1] !== snappedSp1) logSheet.getRange(rIdx, hMap.speed1 + 1).setValue(snappedSp1);
          if (d[hMap.speed2] !== snappedSp2) logSheet.getRange(rIdx, hMap.speed2 + 1).setValue(snappedSp2);
          if (d[hMap.speed3] !== snappedSp3) logSheet.getRange(rIdx, hMap.speed3 + 1).setValue(snappedSp3);
        }
      }
      
      if (!isSeeded) {
        var c1M = mean(c1Vals), c1S = sd(c1Vals, c1M);
        var r1M = mean(r1Vals), r1S = sd(r1Vals, r1M);
        var c2M = mean(c2Vals), c2S = sd(c2Vals, c2M);
        var r2M = mean(r2Vals), r2S = sd(r2Vals, r2M);
        
        if (mMap.c1Mean !== undefined) refSheet.getRange(refRowIdx, mMap.c1Mean + 1).setValue(parseFloat(c1M.toFixed(1)));
        if (mMap.c1SD !== undefined) refSheet.getRange(refRowIdx, mMap.c1SD + 1).setValue(parseFloat(c1S.toFixed(2)));
        if (mMap.r1Mean !== undefined) refSheet.getRange(refRowIdx, mMap.r1Mean + 1).setValue(parseFloat(r1M.toFixed(1)));
        if (mMap.r1SD !== undefined) refSheet.getRange(refRowIdx, mMap.r1SD + 1).setValue(parseFloat(r1S.toFixed(2)));
        if (mMap.c2Mean !== undefined) refSheet.getRange(refRowIdx, mMap.c2Mean + 1).setValue(parseFloat(c2M.toFixed(1)));
        if (mMap.c2SD !== undefined) refSheet.getRange(refRowIdx, mMap.c2SD + 1).setValue(parseFloat(c2S.toFixed(2)));
        if (mMap.r2Mean !== undefined) refSheet.getRange(refRowIdx, mMap.r2Mean + 1).setValue(parseFloat(r2M.toFixed(1)));
        if (mMap.r2SD !== undefined) refSheet.getRange(refRowIdx, mMap.r2SD + 1).setValue(parseFloat(r2S.toFixed(2)));
        
        if (countN > 2) {
          if (mMap.c1Min !== undefined) refSheet.getRange(refRowIdx, mMap.c1Min + 1).setValue(parseFloat((c1M - 3*c1S).toFixed(1)));
          if (mMap.c1Max !== undefined) refSheet.getRange(refRowIdx, mMap.c1Max + 1).setValue(parseFloat((c1M + 3*c1S).toFixed(1)));
          if (mMap.r1Min !== undefined) refSheet.getRange(refRowIdx, mMap.r1Min + 1).setValue(parseFloat((r1M - 3*r1S).toFixed(1)));
          if (mMap.r1Max !== undefined) refSheet.getRange(refRowIdx, mMap.r1Max + 1).setValue(parseFloat((r1M + 3*r1S).toFixed(1)));
          if (mMap.c2Min !== undefined) refSheet.getRange(refRowIdx, mMap.c2Min + 1).setValue(parseFloat((c2M - 3*c2S).toFixed(1)));
          if (mMap.c2Max !== undefined) refSheet.getRange(refRowIdx, mMap.c2Max + 1).setValue(parseFloat((c2M + 3*c2S).toFixed(1)));
          if (mMap.r2Min !== undefined) refSheet.getRange(refRowIdx, mMap.r2Min + 1).setValue(parseFloat((r2M - 3*r2S).toFixed(1)));
          if (mMap.r2Max !== undefined) refSheet.getRange(refRowIdx, mMap.r2Max + 1).setValue(parseFloat((r2M + 3*r2S).toFixed(1)));
        }
      }
      
      var procHealth = isSeeded ? "🟢 Stage 0: SEEDED BLUEPRINT ACTIVE (3σ)" : (countN >= 100 ? "🟢 Stage 4: MATURE SPC LOCKED (3σ)" : "Establishing Baseline");
      if (mMap.healthStamp !== undefined) refSheet.getRange(refRowIdx, mMap.healthStamp + 1).setValue(procHealth);
      if (mMap.controlMode !== undefined) refSheet.getRange(refRowIdx, mMap.controlMode + 1).setValue(isSeeded ? "MANUAL GRACE LIMITS LOADED" : "AUTOMATED STATISTICAL SPC LAYER ACTIVE");
      if (mMap.sampleCount !== undefined) refSheet.getRange(refRowIdx, mMap.sampleCount + 1).setValue(countN);
    }
  }

  // Pass 3: Rolling Batch Math Limits
  var batchStats = {};
  for (var bId in batchGroups) {
    var b = batchGroups[bId];
    var mC1 = mean(b.c1), sC1 = sd(b.c1, mC1);
    var mR1 = mean(b.r1), sR1 = sd(b.r1, mR1);
    var mC2 = mean(b.c2), sC2 = sd(b.c2, mC2);
    var mR2 = mean(b.r2), sR2 = sd(b.r2, mR2);
    var mRF = mean(b.rf), sRF = sd(b.rf, mRF);
    
    batchStats[bId] = {
      c1Min: mC1 - 2*sC1, c1Max: mC1 + 2*sC1,
      r1Min: mR1 - 2*sR1, r1Max: mR1 + 2*sR1,
      c2Min: mC2 - 2*sC2, c2Max: mC2 + 2*sC2,
      r2Min: mR2 - 2*sR2, r2Max: mR2 + 2*sR2,
      rfMin: mRF - 2*sRF, count: b.c1.length
    };
  }

  // Pass 4: In-Memory Multi-Gate Diagnostic Tag Fingerprinting
  var qualityOutputSubMatrix = [];
  for (var r = 1; r < logData.length; r++) {
    var pName = String(logData[r][hMap.programName] || "").trim();
    var serial = String(logData[r][hMap.trueSerial] || "").trim();
    var baseModel = String(logData[r][hMap.baseModel] || "").trim().toLowerCase().replace(/[-_\s]/g, "");
    var cleanPName = pName.toLowerCase().replace(/[-_\s]/g, "");
    var batchId = serial.split("-")[0].trim();
    
    var test1Result = "INITIALIZING"; var test2Result = "INITIALIZING"; var finalStatus = "PASS";
    var diagnosticNotes = "✅ SHOCK IS WITHIN TOLERANCE."; var extNotes = String(logData[r][hMap.teardown] || "").trim();
    var failTags = [];
    
    if (pName && logData[r][0] !== "") {
      // Resolve Dynamic Key lookup from Program_Registry
      var resolvedDynamicKey = modelToDynamicKey[cleanPName] || modelToDynamicKey[baseModel] || pName;
      var cleanResolvedKey  = resolvedDynamicKey.toLowerCase().replace(/[-_\s]/g, "");
      
      var refRow = null;
      for (var mx = 1; mx < refData.length; mx++) {
        var mKey = String(refData[mx][0] || "").trim().replace(/[-_\s]/g, "").toLowerCase();
        if (mKey && (mKey === cleanResolvedKey || cleanResolvedKey.indexOf(mKey) !== -1 || mKey.indexOf(cleanResolvedKey) !== -1)) {
          refRow = refData[mx]; break;
        }
      }
      
      var lowC = parseFloat(logData[r][hMap.comp1]) || 0; var lowR = Math.abs(parseFloat(logData[r][hMap.reb1])) || 0;
      var midC = parseFloat(logData[r][hMap.comp2]) || 0; var midR = Math.abs(parseFloat(logData[r][hMap.reb2])) || 0;
      var sl1  = parseFloat(logData[r][hMap.slope1]) || 0; var individualRF = parseFloat(logData[r][hMap.rodForce]) || 0;
      
      // Global Limit Check Tagging (Test 1)
      if (refRow) {
        var t1Pass = true;
        var valC1Min = parseFloat(refRow[mMap.c1Min]); var valC1Max = parseFloat(refRow[mMap.c1Max]);
        var valR1Min = parseFloat(refRow[mMap.r1Min]); var valR1Max = parseFloat(refRow[mMap.r1Max]);
        var valC2Min = parseFloat(refRow[mMap.c2Min]); var valC2Max = parseFloat(refRow[mMap.c2Max]);
        var valR2Min = parseFloat(refRow[mMap.r2Min]); var valR2Max = parseFloat(refRow[mMap.r2Max]);
        var slope1Min = parseFloat(refRow[mMap.slope1Min]) || 0;
        
        if (!isNaN(valC1Min)) {
          if (lowC < valC1Min || lowC > valC1Max) { t1Pass = false; failTags.push("[C1_FAIL]"); }
          if (lowR < valR1Min || lowR > valR1Max) { t1Pass = false; failTags.push("[R1_FAIL]"); }
          if (midC < valC2Min || midC > valC2Max) { t1Pass = false; failTags.push("[C2_FAIL]"); }
          if (midR < valR2Min || midR > valR2Max) { t1Pass = false; failTags.push("[R2_FAIL]"); }
          if (sl1 < slope1Min) { t1Pass = false; failTags.push("[SLOPE_FAIL]"); }
        }
        test1Result = t1Pass ? "PASS" : "FAIL (BLUEPRINT)";
      }
      
      // Cohort Outlier Check Tagging (Test 2)
      var cStat = batchStats[batchId];
      var defectAnalysis = "";
      if (cStat && cStat.count > 2) {
        var t2Pass = true;
        var lowGasPressure = (individualRF < cStat.rfMin);
        var lowC1 = (lowC < cStat.c1Min || lowC > cStat.c1Max);
        var lowR1 = (lowR < cStat.r1Min || lowR > cStat.r1Max);
        var lowC2 = (midC < cStat.c2Min || midC > cStat.c2Max);
        var lowR2 = (midR < cStat.r2Min || midR > cStat.r2Max);
        
        if (lowGasPressure) { t2Pass = false; failTags.push("[RF_FAIL]"); }
        if (lowC1) { t2Pass = false; failTags.push("[C1_FAIL]"); }
        if (lowR1) { t2Pass = false; failTags.push("[R1_FAIL]"); }
        if (lowC2) { t2Pass = false; failTags.push("[C2_FAIL]"); }
        if (lowR2) { t2Pass = false; failTags.push("[R2_FAIL]"); }
        
        if (!t2Pass) {
          if (lowGasPressure) defectAnalysis = "Gas Pressure Deficient.";
          else if (lowC1 && lowR1) defectAnalysis = "Symmetric drop. Potential bypass.";
          else if (lowC1) defectAnalysis = "Compression outlier variation.";
          else if (lowR1) defectAnalysis = "Rebound outlier variation.";
          else defectAnalysis = "Outlier variance detected.";
        }
        test2Result = t2Pass ? "PASS" : "FAIL (OUTLIER)";
      }
      
      var uniqueFailTags = [];
      for (var f = 0; f < failTags.length; f++) {
        if (uniqueFailTags.indexOf(failTags[f]) === -1) uniqueFailTags.push(failTags[f]);
      }
      
      var cleanExt = extNotes.toLowerCase();
      var globalPass = (test1Result === "INITIALIZING" || !test1Result.includes("FAIL")) && (test2Result === "INITIALIZING" || !test2Result.includes("FAIL"));
      diagnosticNotes = globalPass ? "✅ SHOCK IS WITHIN TOLERANCE." : "❌ ERROR: " + uniqueFailTags.join(" ") + " | " + defectAnalysis;
      
      if (cleanExt.includes("approved") || cleanExt.includes("management")) {
        test1Result = "PASS (OVERRIDE)"; test2Result = "PASS (OVERRIDE)"; finalStatus = "PASS (OVERRIDE)";
        diagnosticNotes = "👔 DISCRETIONARY CLEAR: Released via Management Sign-off.";
      } else if (cleanExt.includes("no issue found")) {
        test1Result = "PASS"; test2Result = "PASS"; finalStatus = "PASS";
        diagnosticNotes = "🛠️ TEARDOWN VALIDATED: Assembly clear.";
      } else if (!globalPass) {
        finalStatus = "FAIL";
      }
    }
    qualityOutputSubMatrix.push([test1Result, test2Result, finalStatus, diagnosticNotes]);
  }
  
  if (qualityOutputSubMatrix.length > 0) {
    logSheet.getRange(2, 22, qualityOutputSubMatrix.length, 4).setValues(qualityOutputSubMatrix);
  }
  SpreadsheetApp.flush();
}

// =========================================================================
// ENGINE 2: OPERATOR CONSOLE INTERFACE CONTROLLER
// =========================================================================
function manageOperatorStation(e) {
  var ss = e ? e.source : SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName("Operator_Station"); if (!sheet) return;
  var range = e ? e.range : sheet.getRange("C2"); var barcode = String(sheet.getRange("C2").getValue()).trim();
  var verifiedFileIdStr = "";
  
  if (range.getA1Notation() === "C2") {
    sheet.getRange("C3:C5").clearContent(); 
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
    sheet.getRange("C3").setFormula('=HYPERLINK("' + fileUrl + '", "🔗 Open ' + realFileName + '")'); sheet.getRange("Z1").setValue(verifiedFileIdStr);
    
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
          if (regPartClean === cleanWoPart && regProgName) {
            sheet.getRange("C16").setValue(regProgName); break;
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
    
    if (matrixRowIdx !== -1) {
      var row = refData[matrixRowIdx];
      sheet.getRange("B22").setValue(parseFloat(row[mMap.c1Min]) || 0); sheet.getRange("C22").setValue(parseFloat(row[mMap.c1Max]) || 0);
      sheet.getRange("D22").setValue(parseFloat(row[mMap.r1Min]) || 0); sheet.getRange("E22").setValue(parseFloat(row[mMap.r1Max]) || 0);
      sheet.getRange("B23").setValue(parseFloat(row[mMap.c2Min]) || 0); sheet.getRange("C23").setValue(parseFloat(row[mMap.c2Max]) || 0);
      sheet.getRange("D23").setValue(parseFloat(row[mMap.r2Min]) || 0); sheet.getRange("E23").setValue(parseFloat(row[mMap.r2Max]) || 0);
      var rawSlopeVal = parseFloat(row[mMap.slope1Min]);
      sheet.getRange("F22").setValue(!isNaN(rawSlopeVal) ? rawSlopeVal.toFixed(1) : "No Limit");
    }
    
    if (verifiedFileIdStr !== "") {
      var activeWoSpreadsheet = SpreadsheetApp.openById(verifiedFileIdStr); var activeWoSheet = activeWoSpreadsheet.getSheets()[0]; var totalWoRows = activeWoSheet.getLastRow();
      if (totalWoRows >= 12) {
        var serials = activeWoSheet.getRange(12, 1, totalWoRows - 11, 1).getValues(); var outputGrid = [];
        var backgroundColorsGrid = []; var fontColorsGrid = []; var fontWeightsGrid = [];
        
        // PASS 1: Aggregation loop using zero-insensitive token-tail boundaries
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

        // PASS 2: Table Generation & Diagnostic Tag Failure Highlighter Pass
        for (var i = 0; i < serials.length; i++) {
          var snRaw = String(serials[i][0] || "").trim();
          if (snRaw !== "") {
            var matchedRow = resolvedRows[i];
            
            var rowBackgrounds = ["", "", "", "", "", "", "", "", "", "", "", ""];
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
              
              // 🚀 RELIABLE FINGERPRINT PARSER ENGINE FOR OUTLIER CELL HIGHLIGHTS
              var applyFaultHighlight = function(colIndex) {
                rowFonts[colIndex] = "#FF0000";       
                rowWeights[colIndex] = "bold";        
                rowBackgrounds[colIndex] = "#FADBD8"; 
              };

              if (!skipHighlight && !diagnosticNotes.includes("✅")) {
                if (diagnosticNotes.indexOf("[RF_FAIL]") !== -1)    applyFaultHighlight(1); // Column B
                if (diagnosticNotes.indexOf("[C1_FAIL]") !== -1)    applyFaultHighlight(2); // Column C
                if (diagnosticNotes.indexOf("[R1_FAIL]") !== -1)    applyFaultHighlight(3); // Column D
                if (diagnosticNotes.indexOf("[C2_FAIL]") !== -1)    applyFaultHighlight(4); // Column E
                if (diagnosticNotes.indexOf("[R2_FAIL]") !== -1)    applyFaultHighlight(5); // Column F
                if (diagnosticNotes.indexOf("[SLOPE_FAIL]") !== -1) { applyFaultHighlight(2); applyFaultHighlight(3); }
              }
              
              if (test1Result.includes("FAIL")) {
                rowBackgrounds[6] = "#FADBD8"; rowFonts[6] = "#C0392B"; rowWeights[6] = "bold";
              } else if (test1Result.includes("PASS")) {
                rowBackgrounds[6] = "#D4EFDF"; rowFonts[6] = "#196F3D";
              }

              if (test2ResultString.includes("FAIL")) {
                rowBackgrounds[7] = "#FADBD8"; rowFonts[7] = "#C0392B"; rowWeights[7] = "bold";
              } else if (test2ResultString.includes("PASS")) {
                rowBackgrounds[7] = "#D4EFDF"; rowFonts[7] = "#196F3D";
              }
              
              // 🚀 CRITICAL OVERALL STATUS BLOCK (White text over solid crimson background block)
              if (totalStatus.includes("FAIL")) { 
                rowBackgrounds[8] = "#C0392B"; rowFonts[8] = "#FFFFFF"; rowWeights[8] = "bold"; 
              } else if (totalStatus.includes("OVERRIDE")) { 
                rowBackgrounds[8] = "#D4EFDF"; rowFonts[8] = "#196F3D"; rowWeights[8] = "bold"; 
              } else if (totalStatus.includes("PASS")) {
                rowBackgrounds[8] = "#D4EFDF"; rowFonts[8] = "#196F3D";
              }
              
              var sheetId = masterLogSheet.getSheetId();
              var linkFormula = '=HYPERLINK("' + ss.getUrl() + '#gid=' + sheetId + '&range=X' + logSheetRowIndexNumber + '", "' + pure9Digits + '")';
              
              outputGrid.push([ linkFormula, individualRF.toFixed(1), lowC.toFixed(1), lowR.toFixed(1), midC.toFixed(1), midR.toFixed(1), test1Result, test2ResultString, totalStatus, logSheetTeardownStatus, diagnosticNotes, rawEngCommentText ]);
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
          
          var dynamicLastRow = 27 + outputGrid.length - 1;
          var statusBarFormula = '=IF(OR(COUNTIF(I27:I' + dynamicLastRow + ', "*FAIL*") > 0, COUNTIF(G27:H' + dynamicLastRow + ', "*FAIL*") > 0), "🔴 WORK ORDER BLOCKED: FAILED UNITS REQUIRING EVALUATION", "🟢 WORK ORDER COMPLETED & PASSING")';
          sheet.getRange("A8").setFormula(statusBarFormula);
          sheet.getRange("A11").setFormula(statusBarFormula);
        }
      }
    }
  } catch(err) { sheet.getRange("C3").setValue("⚠️ Sync Error: " + err.toString()); }
}

// =========================================================================
// UNIFIED AUTHORIZED BACKGROUND EDIT LISTENERS
// =========================================================================
function installableConsoleTrigger(e) {
  if (e && e.range.getSheet().getName() === "Operator_Station" && e.range.getA1Notation() === "C2") manageOperatorStation(e);
}

function installableOnEdit(e) {
  if (e && e.range.getSheet().getName() === "Operator_Station" && e.range.getA1Notation() === "C2") manageOperatorStation({ source: e.source, range: e.range });
}