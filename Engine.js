// =========================================================================
// 🏎️ UNIFIED DYNO ENGINE (Engine.gs)
// Adaptive Multi-File Ingestion, Hysteresis Regression & SPC Recalculator
// =========================================================================

/**
 * Calculates arithmetic mean of an array, filtering invalid values.
 */
function calculateMean(arr) {
  var clean = arr.filter(function(x) { return !isNaN(x) && x !== null && x !== ""; });
  return clean.length === 0 ? 0 : clean.reduce(function(a, b) { return a + b; }, 0) / clean.length;
}

/**
 * Calculates sample standard deviation of an array.
 */
function calculateSD(arr, m) {
  var clean = arr.filter(function(x) { return !isNaN(x) && x !== null && x !== ""; });
  if (clean.length <= 1) return 0;
  var meanVal = (m !== undefined) ? m : calculateMean(clean);
  var variance = clean.map(function(x) { return Math.pow(x - meanVal, 2); }).reduce(function(a, b) { return a + b; }, 0) / (clean.length - 1);
  return Math.sqrt(variance);
}

/**
 * Buckets speed values to nominal target speeds.
 */
function snapToNominalSpeed(val) {
  var v = parseFloat(val);
  if (isNaN(v) || v <= 0) return 0;
  var nominals = CONFIG.NOMINAL_SPEEDS || [100, 400, 1000, 2500];
  if (v < 250) return nominals[0];   // 100
  if (v < 700) return nominals[1];   // 400
  if (v < 1800) return nominals[2];  // 1000
  return nominals[3];                // 2500
}

/**
 * Builds header map dictionary from Master_Dyno_Log header row.
 */
function buildHeaderMap(headerRow) {
  var map = {};
  if (!headerRow) return map;
  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG;
  
  map.timestamp = logCols.TIMESTAMP - 1;
  map.programName = logCols.PROGRAM_NAME - 1;
  map.trueSerial = logCols.TRUE_SERIAL - 1;
  map.baseModel = logCols.BASE_MODEL - 1;
  map.valvingVersion = logCols.VALVING_VERSION - 1;
  map.rodForce = logCols.ROD_FORCE - 1;
  map.speed1 = logCols.SPEED_1 - 1;
  map.comp1 = logCols.COMP_1 - 1;
  map.reb1 = logCols.REB_1 - 1;
  map.slope1 = logCols.SLOPE_1 - 1;
  map.loopArea1 = logCols.LOOP_AREA_1 - 1;
  map.speed2 = logCols.SPEED_2 - 1;
  map.comp2 = logCols.COMP_2 - 1;
  map.reb2 = logCols.REB_2 - 1;
  map.loopArea2 = logCols.LOOP_AREA_2 - 1;
  map.speed3 = logCols.SPEED_3 - 1;
  map.comp3 = logCols.COMP_3 - 1;
  map.reb3 = logCols.REB_3 - 1;
  map.test1Status = logCols.TEST_1_STATUS - 1;
  map.test2Status = logCols.TEST_2_STATUS - 1;
  map.overallStatus = logCols.OVERALL_STATUS - 1;
  map.diagnostics = logCols.DIAGNOSTICS - 1;
  map.evaluationAction = logCols.EVALUATION_ACTION - 1;
  map.engComments = logCols.ENG_COMMENTS - 1;
  
  return map;
}

/**
 * Builds matrix header map dictionary from Part_Reference_Matrix header row.
 */
function buildMatrixHeaderMap(headerRow) {
  var refCols = CONFIG.COLUMNS.PART_REFERENCE_MATRIX;
  return {
    dynamicKey: refCols.DYNAMIC_KEY - 1,
    speed1: refCols.SPEED_1 - 1,
    speed2: refCols.SPEED_2 - 1,
    speed3: refCols.SPEED_3 - 1,
    c1Mean: refCols.COMP_1_MEAN - 1,
    c1SD: refCols.COMP_1_SD - 1,
    c1Min: refCols.COMP_1_MIN - 1,
    c1Max: refCols.COMP_1_MAX - 1,
    r1Mean: refCols.REB_1_MEAN - 1,
    r1SD: refCols.REB_1_SD - 1,
    r1Min: refCols.REB_1_MIN - 1,
    r1Max: refCols.REB_1_MAX - 1,
    slope1Min: refCols.SLOPE_1_MIN - 1,
    c2Mean: refCols.COMP_2_MEAN - 1,
    c2SD: refCols.COMP_2_SD - 1,
    c2Min: refCols.COMP_2_MIN - 1,
    c2Max: refCols.COMP_2_MAX - 1,
    r2Mean: refCols.REB_2_MEAN - 1,
    r2SD: refCols.REB_2_SD - 1,
    r2Min: refCols.REB_2_MIN - 1,
    r2Max: refCols.REB_2_MAX - 1,
    c3Mean: refCols.COMP_3_MEAN - 1,
    c3SD: refCols.COMP_3_SD - 1,
    c3Min: refCols.COMP_3_MIN - 1,
    c3Max: refCols.COMP_3_MAX - 1,
    r3Mean: refCols.REB_3_MEAN - 1,
    r3SD: refCols.REB_3_SD - 1,
    r3Min: refCols.REB_3_MIN - 1,
    r3Max: refCols.REB_3_MAX - 1,
    sampleCount: refCols.SAMPLE_COUNT - 1,
    healthStamp: refCols.HEALTH_STAMP - 1,
    controlMode: refCols.CONTROL_MODE - 1
  };
}

// =========================================================================
// ENGINE 1: ADAPTIVE DYNO PROCESSOR (DEV-SAFE STACK APPEND)
// =========================================================================
function processDynoFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  if (!sheet) return;
  
  var watchFolderName = CONFIG.FOLDERS.WATCH_FOLDER || "01_Watch_Folder_DEV"; 
  var archiveFolderName = CONFIG.FOLDERS.ARCHIVE_FOLDER || "02_Archive_DEV";
  
  var folders = DriveApp.getFoldersByName(watchFolderName);
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
  
  var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY); var registryData = registrySheet ? registrySheet.getDataRange().getValues() : [];
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
        for (var c = 0; c < 24; c++) { outputRowArray.push(""); }
        
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
        if (hMap.test1Status !== undefined) outputRowArray[hMap.test1Status] = "PASS";
        if (hMap.test2Status !== undefined) outputRowArray[hMap.test2Status] = "PASS";
        if (hMap.overallStatus !== undefined) outputRowArray[hMap.overallStatus] = "PASS";
        
        var dynamicNextRowIndex = sheet.getLastRow() + 1;
        sheet.getRange(dynamicNextRowIndex, 1, 1, 24).setValues([outputRowArray]);
        
        var archiveFolder;
        var archiveFolders = DriveApp.getFoldersByName(archiveFolderName);
        if (archiveFolders.hasNext()) {
          archiveFolder = archiveFolders.next();
        } else {
          archiveFolder = folder.createFolder(archiveFolderName);
        }

        pack.pvp.moveTo(archiveFolder); 
        for (var f = 0; f < pack.intervals.length; f++) { 
          pack.intervals[f].moveTo(archiveFolder); 
        }
      } catch(err) { Logger.log("Watch folder processing failed: " + err.toString()); }
    }
  }
}

// =========================================================================
// HISTORICAL RE-RUN ENGINE (STRICT SELF-HEALING UNIFIED SPC BALANCER)
// =========================================================================
function retroactiveLogRecalculate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);
  var regSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);
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
  
  // Build BASE_MODEL & PROGRAM_NAME -> DYNAMIC_KEY map from Program_Registry
  var modelToDynamicKey = {};
  if (regSheet) {
    var regValues = regSheet.getDataRange().getValues();
    var regCols = CONFIG.COLUMNS.PROGRAM_REGISTRY;
    for (var k = 1; k < regValues.length; k++) {
      var bm = String(regValues[k][regCols.BASE_MODEL - 1] || "").trim().toLowerCase().replace(/[-_\s]/g, "");
      var pn = String(regValues[k][regCols.PROGRAM_NAME - 1] || "").trim().toLowerCase().replace(/[-_\s]/g, "");
      var dk = String(regValues[k][regCols.DYNAMIC_KEY - 1] || "").trim();
      
      var targetKey = dk || regValues[k][regCols.PROGRAM_NAME - 1] || regValues[k][regCols.BASE_MODEL - 1];
      if (bm) modelToDynamicKey[bm] = targetKey;
      if (pn) modelToDynamicKey[pn] = targetKey;
    }
  }

  // Pass 1: Filter & Deduplicate for SPC Baseline Pool
  for (var r = 1; r < logData.length; r++) {
    var prog = String(logData[r][hMap.programName] || "").trim();
    var serial = String(logData[r][hMap.trueSerial] || "").trim();
    var baseModel = String(logData[r][hMap.baseModel] || "").trim().toLowerCase().replace(/[-_\s]/g, "");
    var cleanProgName = prog.toLowerCase().replace(/[-_\s]/g, "");
    var overallStatus = String(logData[r][hMap.overallStatus] || "").toUpperCase().trim();
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
    
    // Resolve DYNAMIC_KEY via Base Model or Program Name lookup
    var resolvedGroupKey = modelToDynamicKey[baseModel] || modelToDynamicKey[cleanProgName] || prog;
    var isPassingRun = overallStatus.indexOf("PASS") !== -1;

    if (resolvedGroupKey !== "" && isPassingRun) {
      if (!historicalGroups[resolvedGroupKey]) {
        historicalGroups[resolvedGroupKey] = {};
      }
      logData[r]._rowIdx = r + 1;
      
      var serialKey = serial !== "" ? serial.toLowerCase() : ("row_" + (r + 1));
      historicalGroups[resolvedGroupKey][serialKey] = logData[r];
    }
  }
  
  // Pass 2: Calculate Baselines across speeds
  for (var pName in historicalGroups) {
    var serialMap = historicalGroups[pName];
    var pool = Object.keys(serialMap).map(function(k) { return serialMap[k]; });
    var countN = pool.length;
    var cleanPName = pName.replace(/[-_\s]/g, "").toLowerCase();
    
    var refRowIdx = -1;
    for (var mx = 1; mx < refData.length; mx++) {
      var matrixKeyRaw = String(refData[mx][mMap.dynamicKey] || "").trim();
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
    var diagnosticNotes = "✅ SHOCK IS WITHIN TOLERANCE."; var extNotes = String(logData[r][hMap.diagnostics] || "").trim();
    var failTags = [];
    
    if (pName && logData[r][0] !== "") {
      // Resolve DYNAMIC_KEY via Program Registry lookup or fallback to raw name
      var resolvedDynamicKey = modelToDynamicKey[baseModel] || modelToDynamicKey[cleanPName] || pName;
      var cleanResolvedKey = resolvedDynamicKey.toLowerCase().replace(/[-_\s]/g, "");
      
      var refRow = null;
      for (var mx = 1; mx < refData.length; mx++) {
        var mKey = String(refData[mx][mMap.dynamicKey] || "").trim().toLowerCase().replace(/[-_\s]/g, "");
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
    var startColIdx = (CONFIG.COLUMNS.MASTER_DYNO_LOG.TEST_1_STATUS) || 19;
    logSheet.getRange(2, startColIdx, qualityOutputSubMatrix.length, 4).setValues(qualityOutputSubMatrix);
  }
  SpreadsheetApp.flush();
}