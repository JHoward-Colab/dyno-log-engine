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
  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG;[cite: 1]
  
  map.timestamp = logCols.TIMESTAMP - 1;[cite: 1]
  map.programName = logCols.PROGRAM_NAME - 1;[cite: 1]
  map.trueSerial = logCols.TRUE_SERIAL - 1;[cite: 1]
  map.baseModel = logCols.BASE_MODEL - 1;[cite: 1]
  map.valvingVersion = logCols.VALVING_VERSION - 1;[cite: 1]
  map.rodForce = logCols.ROD_FORCE - 1;[cite: 1]
  map.speed1 = logCols.SPEED_1 - 1;[cite: 1]
  map.comp1 = logCols.COMP_1 - 1;[cite: 1]
  map.reb1 = logCols.REB_1 - 1;[cite: 1]
  map.slope1 = logCols.SLOPE_1 - 1;[cite: 1]
  map.loopArea1 = logCols.LOOP_AREA_1 - 1;[cite: 1]
  map.speed2 = logCols.SPEED_2 - 1;[cite: 1]
  map.comp2 = logCols.COMP_2 - 1;[cite: 1]
  map.reb2 = logCols.REB_2 - 1;[cite: 1]
  map.loopArea2 = logCols.LOOP_AREA_2 - 1;[cite: 1]
  map.speed3 = logCols.SPEED_3 - 1;[cite: 1]
  map.comp3 = logCols.COMP_3 - 1;[cite: 1]
  map.reb3 = logCols.REB_3 - 1;[cite: 1]
  map.test1Status = logCols.TEST_1_STATUS - 1;[cite: 1]
  map.test2Status = logCols.TEST_2_STATUS - 1;[cite: 1]
  map.overallStatus = logCols.OVERALL_STATUS - 1;[cite: 1]
  map.diagnostics = logCols.DIAGNOSTICS - 1;[cite: 1]
  map.evaluationAction = logCols.EVALUATION_ACTION - 1;[cite: 1]
  map.engComments = logCols.ENG_COMMENTS - 1;[cite: 1]

  // Dynamic header scanner fallback (e.g. if Column Z holds Evaluation Action)
  for (var i = 0; i < headerRow.length; i++) {
    var cleanH = String(headerRow[i] || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    if (cleanH.indexOf("evaluation") !== -1 || cleanH.indexOf("action") !== -1) map.evaluationAction = i;
    if (cleanH.indexOf("comment") !== -1 || cleanH.indexOf("notes") !== -1) map.engComments = i;
  }
  
  return map;
}

/**
 * Builds matrix header map dictionary from Part_Reference_Matrix header row.
 */
function buildMatrixHeaderMap(headerRow) {
  var refCols = CONFIG.COLUMNS.PART_REFERENCE_MATRIX;[cite: 1]
  return {
    dynamicKey: refCols.DYNAMIC_KEY - 1,[cite: 1]
    speed1: refCols.SPEED_1 - 1,[cite: 1]
    speed2: refCols.SPEED_2 - 1,[cite: 1]
    speed3: refCols.SPEED_3 - 1,[cite: 1]
    c1Mean: refCols.COMP_1_MEAN - 1,[cite: 1]
    c1SD: refCols.COMP_1_SD - 1,[cite: 1]
    c1Min: refCols.COMP_1_MIN - 1,[cite: 1]
    c1Max: refCols.COMP_1_MAX - 1,[cite: 1]
    r1Mean: refCols.REB_1_MEAN - 1,[cite: 1]
    r1SD: refCols.REB_1_SD - 1,[cite: 1]
    r1Min: refCols.REB_1_MIN - 1,[cite: 1]
    r1Max: refCols.REB_1_MAX - 1,[cite: 1]
    slope1Min: refCols.SLOPE_1_MIN - 1,[cite: 1]
    loopArea1Min: refCols.LOOP_AREA_1_MIN - 1,[cite: 1]
    c2Mean: refCols.COMP_2_MEAN - 1,[cite: 1]
    c2SD: refCols.COMP_2_SD - 1,[cite: 1]
    c2Min: refCols.COMP_2_MIN - 1,[cite: 1]
    c2Max: refCols.COMP_2_MAX - 1,[cite: 1]
    r2Mean: refCols.REB_2_MEAN - 1,[cite: 1]
    r2SD: refCols.REB_2_SD - 1,[cite: 1]
    r2Min: refCols.REB_2_MIN - 1,[cite: 1]
    r2Max: refCols.REB_2_MAX - 1,[cite: 1]
    slope2Min: refCols.SLOPE_2_MIN - 1,[cite: 1]
    c3Mean: refCols.COMP_3_MEAN - 1,[cite: 1]
    c3SD: refCols.COMP_3_SD - 1,[cite: 1]
    c3Min: refCols.COMP_3_MIN - 1,[cite: 1]
    c3Max: refCols.COMP_3_MAX - 1,[cite: 1]
    r3Mean: refCols.REB_3_MEAN - 1,[cite: 1]
    r3SD: refCols.REB_3_SD - 1,[cite: 1]
    r3Min: refCols.REB_3_MIN - 1,[cite: 1]
    r3Max: refCols.REB_3_MAX - 1,[cite: 1]
    sampleCount: refCols.SAMPLE_COUNT - 1,[cite: 1]
    // Engineering SPC Metrics (Columns AG to AN)
    rawCompSD: refCols.RAW_COMP_SD - 1,[cite: 1]
    rawRebSD: refCols.RAW_REB_SD - 1,[cite: 1]
    origCompBase: refCols.ORIGINAL_COMP_BASE - 1,[cite: 1]
    origRebBase: refCols.ORIGINAL_REB_BASE - 1,[cite: 1]
    compRangeWidth: refCols.COMP_RANGE_WIDTH - 1,[cite: 1]
    rebRangeWidth: refCols.REB_RANGE_WIDTH - 1,[cite: 1]
    compDriftPct: refCols.COMP_DRIFT_PCT - 1,[cite: 1]
    rebDriftPct: refCols.REB_DRIFT_PCT - 1,[cite: 1]
    healthStamp: refCols.HEALTH_STAMP - 1,[cite: 1]
    controlMode: refCols.CONTROL_MODE - 1[cite: 1]
  };
}

// =========================================================================
// ENGINE 1: ADAPTIVE DYNO PROCESSOR (DEV-SAFE STACK APPEND)
// =========================================================================
function processDynoFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();[cite: 1]
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);[cite: 1]
  if (!sheet) return;
  
  var watchFolderName = CONFIG.FOLDERS.WATCH_FOLDER || "01_Watch_Folder_DEV";[cite: 1]
  var archiveFolderName = CONFIG.FOLDERS.ARCHIVE_FOLDER || "02_Archive_DEV";[cite: 1]
  
  var folders = DriveApp.getFoldersByName(watchFolderName);[cite: 1]
  var folder;
  if (folders.hasNext()) { folder = folders.next(); } else { return; }
  
  var files = folder.getFiles();[cite: 1]
  var tempFileMap = {};
  var now = new Date();
  
  while (files.hasNext()) {
    var file = files.next(); var rawName = file.getName(); var lastUpdated = file.getLastUpdated().getTime();[cite: 1]
    if (rawName.toLowerCase().endsWith('.csv')) {[cite: 1]
      var fileName = rawName.replace(/\.csv$/i, "").replace(/\.txt$/i, "").trim(); var tempGroupKey = "";[cite: 1]
      var serialMatch = rawName.match(/\d{6}-\d{3}/) || rawName.match(/\d+-\d+/);[cite: 1]
      if (serialMatch) { tempGroupKey = serialMatch[0].trim(); } else {[cite: 1]
        var clearParts = fileName.split(/[\s_]+/);[cite: 1]
        tempGroupKey = clearParts.length >= 3 ? clearParts[clearParts.length - 2].trim() : clearParts[0].trim();[cite: 1]
      }
      
      var detectedModel = "";[cite: 1]
      if (fileName.toLowerCase().includes("pvp")) {[cite: 1]
        detectedModel = fileName.split("pvp")[0].replace(/[^a-zA-Z0-9_]/g, " ").trim().split(" ")[0];[cite: 1]
      } else {
        var spaceParts = fileName.split(" ");[cite: 1]
        if (spaceParts[0].includes("_V")) {[cite: 1]
          detectedModel = spaceParts[0].trim();[cite: 1]
        } else if (rawName.includes(tempGroupKey)) {[cite: 1]
          var preSerialText = rawName.split(tempGroupKey)[0].trim();[cite: 1]
          if (preSerialText.endsWith('_') || preSerialText.endsWith(' ')) { preSerialText = preSerialText.slice(0, -1).trim(); }[cite: 1]
          detectedModel = preSerialText;[cite: 1]
        }
      }
      if (detectedModel === "" || detectedModel.startsWith('_PROCESSING_')) {[cite: 1]
        var rawParts = fileName.split(/[\s_]+/); detectedModel = rawParts[0].replace('_PROCESSING_', '');[cite: 1]
        if (rawParts[1] && rawParts[1].toUpperCase().startsWith('V')) { detectedModel += "_" + rawParts[1]; }[cite: 1]
      }
      
      var speedSuffix = "";[cite: 1]
      if (fileName.toLowerCase().includes("pvp")) {[cite: 1]
        speedSuffix = "pvp";[cite: 1]
      } else {
        var nameParts = fileName.split(/[\s_]+/); [cite: 1]
        speedSuffix = nameParts[nameParts.length - 1].toLowerCase().trim();[cite: 1]
        if (speedSuffix === "interval" && nameParts.length >= 2) {[cite: 1]
          speedSuffix = nameParts[nameParts.length - 2].toLowerCase().trim();[cite: 1]
        }
      }
      
      if (speedSuffix === 'pvp' || !isNaN(parseInt(speedSuffix))) {[cite: 1]
        if (!tempFileMap[tempGroupKey]) { tempFileMap[tempGroupKey] = { pvp: null, intervals: [], youngestFileTime: 0, parsedBaseModel: detectedModel }; }[cite: 1]
        if (lastUpdated > tempFileMap[tempGroupKey].youngestFileTime) tempFileMap[tempGroupKey].youngestFileTime = lastUpdated;[cite: 1]
        if (speedSuffix === 'pvp') { tempFileMap[tempGroupKey].pvp = file; } else { tempFileMap[tempGroupKey].intervals.push(file); }[cite: 1]
      }
    }
  }
  
  var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY); var registryData = registrySheet ? registrySheet.getDataRange().getValues() : [];[cite: 1]
  var masterLogData = sheet.getDataRange().getValues();[cite: 1]
  var hMap = buildHeaderMap(masterLogData[0]);[cite: 1]
  
  for (var tempKey in tempFileMap) {[cite: 1]
    var pack = tempFileMap[tempKey];[cite: 1]
    if (pack.pvp && pack.intervals.length >= 3) {[cite: 1]
      if ((now.getTime() - pack.youngestFileTime) / 1000 < 4) continue;[cite: 1]
      try {
        var rawBlobStr = pack.pvp.getBlob().getDataAsString("UTF-8").trim();[cite: 1]
        var pvpContent = Utilities.parseCsv(rawBlobStr); var trueSerial = tempKey;[cite: 1]
        var trueDynoProgramName = pack.parsedBaseModel;[cite: 1]
        
        if (trueDynoProgramName.length > 20) {[cite: 1]
          trueDynoProgramName = trueDynoProgramName.substring(0, 14) + "_V" + trueDynoProgramName.slice(-1);[cite: 1]
        }
        
        var cleanTrueProgName = trueDynoProgramName.replace(/[-_\s]/g, "").toLowerCase();[cite: 1]
        for (var rR = 1; rR < registryData.length; rR++) {[cite: 1]
          var regKey = String(registryData[rR][0] || "").trim().replace(/[-_\s]/g, "").toLowerCase();[cite: 1]
          if (regKey === cleanTrueProgName) { [cite: 1]
            trueDynoProgramName = String(registryData[rR][0] || "").trim(); break;  [cite: 1]
          }
        }
        
        var registryBaseModelText = trueDynoProgramName; var registryValvingVersion = "PRODUCTION_RUN";[cite: 1]
        var cleanMatchName = trueDynoProgramName.replace(/[-_\s]/g, "").toLowerCase();[cite: 1]
        for (var regR = 1; regR < registryData.length; regR++) {[cite: 1]
          var cleanRegCell = String(registryData[regR][0] || "").trim().replace(/[-_\s]/g, "").toLowerCase();[cite: 1]
          if (cleanRegCell === cleanMatchName) {[cite: 1]
            registryBaseModelText = String(registryData[regR][2] || "").trim();[cite: 1]
            registryValvingVersion = String(registryData[regR][6] || "").trim(); [cite: 1]
            break;[cite: 1]
          }
        }
        
        var rawRodForce = 0; var intervalMetrics = [];[cite: 1]
        for (var i = 0; i < pack.intervals.length; i++) {[cite: 1]
          var file = pack.intervals[i]; var rawIntStr = file.getBlob().getDataAsString("UTF-8").trim();[cite: 1]
          if (rawIntStr === "") continue; var rows = Utilities.parseCsv(rawIntStr); var speedTarget = 0;[cite: 1]
          for (var j = 0; j < Math.min(rows.length, 15); j++) { if (rows[j][0] && rows[j][0].trim().toLowerCase() == "velocity amplitude") { speedTarget = parseInt(rows[j][1]); break; } }[cite: 1]
          
          if (!speedTarget || isNaN(speedTarget)) {[cite: 1]
            for (var r = 14; r < rows.length; r++) {[cite: 1]
              if (rows[r][0] && !isNaN(parseFloat(rows[r][0]))) {[cite: 1]
                var testV = Math.abs(parseFloat(rows[r][0]));[cite: 1]
                if (testV > speedTarget) speedTarget = testV;[cite: 1]
              }
            }
          }
          speedTarget = snapToNominalSpeed(speedTarget);[cite: 1]
          
          for (var j = 0; j < rows.length; j++) { if (rows[j][0] && rows[j][0].trim().toLowerCase() == "rod force") { rawRodForce = parseFloat(rows[j][1]); break; } }[cite: 1]
          var maxComp = -9999, maxReb = 9999, totalArea = 0, nearZeroPoints = [];[cite: 1]
          for (var r = 14; r < rows.length; r++) {[cite: 1]
            if (rows[r].length < 2 || isNaN(parseFloat(rows[r][0]))) continue;[cite: 1]
            var vel = parseFloat(rows[r][0]); var force = parseFloat(rows[r][1]);[cite: 1]
            if (force > maxComp) maxComp = force; if (force < maxReb) maxReb = force;[cite: 1]
            totalArea += Math.abs(force); if (Math.abs(vel) < ((speedTarget || 500) * 0.15)) nearZeroPoints.push({ x: vel, y: force });[cite: 1]
          }
          var sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;[cite: 1]
          for (var p = 0; p < nearZeroPoints.length; p++) { sumX += nearZeroPoints[p].x; sumY += nearZeroPoints[p].y; sumXY += (nearZeroPoints[p].x * nearZeroPoints[p].y); var varXX = (nearZeroPoints[p].x * nearZeroPoints[p].x); sumXX += varXX; }[cite: 1]
          var denom = (nearZeroPoints.length * sumXX) - (sumX * sumX);[cite: 1]
          var slope = denom == 0 ? 0 : ((nearZeroPoints.length * sumXY) - (sumX * sumY)) / denom;[cite: 1]
          intervalMetrics.push({ speed: speedTarget, maxComp: maxComp, maxReb: maxReb, slope: slope, area: totalArea });[cite: 1]
        }
        intervalMetrics.sort(function(a, b) { return a.speed - b.speed; });[cite: 1]
        
        if (intervalMetrics.length >= 3) {[cite: 1]
          var lowSp = intervalMetrics[0].speed;[cite: 1]
          var highSp = intervalMetrics[2].speed;[cite: 1]
          if (lowSp <= 100 || highSp <= 1000) {[cite: 1]
            intervalMetrics[0].speed = 100;[cite: 1]
            intervalMetrics[1].speed = 400;[cite: 1]
            intervalMetrics[2].speed = 1000;[cite: 1]
          } else {
            intervalMetrics[0].speed = 500;[cite: 1]
            intervalMetrics[1].speed = 1000;[cite: 1]
            intervalMetrics[2].speed = 2500;[cite: 1]
          }
        }
        
        var pvpSlots = [{ comp: 0, reb: 0 }, { comp: 0, reb: 0 }, { comp: 0, reb: 0 }];[cite: 1]
        if (pvpContent.length >= 11) {[cite: 1]
          pvpSlots[0].comp = parseFloat(pvpContent[8][3]) || 0; pvpSlots[0].reb = parseFloat(pvpContent[8][5]) || 0;[cite: 1]
          pvpSlots[1].comp = parseFloat(pvpContent[9][3]) || 0; pvpSlots[1].reb = parseFloat(pvpContent[9][5]) || 0;[cite: 1]
          pvpSlots[2].comp = parseFloat(pvpContent[10][3]) || 0; pvpSlots[2].reb = parseFloat(pvpContent[10][5]) || 0;[cite: 1]
        }
        
        var outputRowArray = [];[cite: 1]
        for (var c = 0; c < 24; c++) { outputRowArray.push(""); }[cite: 1]
        
        if (hMap.timestamp !== undefined) outputRowArray[hMap.timestamp] = new Date();[cite: 1]
        if (hMap.programName !== undefined) outputRowArray[hMap.programName] = trueDynoProgramName;[cite: 1]
        if (hMap.trueSerial !== undefined) outputRowArray[hMap.trueSerial] = trueSerial;[cite: 1]
        if (hMap.baseModel !== undefined) outputRowArray[hMap.baseModel] = registryBaseModelText; [cite: 1]
        if (hMap.valvingVersion !== undefined) outputRowArray[hMap.valvingVersion] = registryValvingVersion; [cite: 1]
        if (hMap.rodForce !== undefined) outputRowArray[hMap.rodForce] = rawRodForce;[cite: 1]
        
        if (intervalMetrics[0]) { [cite: 1]
          if (hMap.speed1 !== undefined) outputRowArray[hMap.speed1] = intervalMetrics[0].speed;[cite: 1]
          if (hMap.comp1 !== undefined) outputRowArray[hMap.comp1] = pvpSlots[0].comp; [cite: 1]
          if (hMap.reb1 !== undefined) outputRowArray[hMap.reb1] = pvpSlots[0].reb; [cite: 1]
          if (hMap.slope1 !== undefined) outputRowArray[hMap.slope1] = intervalMetrics[0].slope; [cite: 1]
          if (hMap.loopArea1 !== undefined) outputRowArray[hMap.loopArea1] = intervalMetrics[0].area; [cite: 1]
        }
        if (intervalMetrics[1]) { [cite: 1]
          if (hMap.speed2 !== undefined) outputRowArray[hMap.speed2] = intervalMetrics[1].speed;[cite: 1]
          if (hMap.comp2 !== undefined) outputRowArray[hMap.comp2] = pvpSlots[1].comp; [cite: 1]
          if (hMap.reb2 !== undefined) outputRowArray[hMap.reb2] = pvpSlots[1].reb; [cite: 1]
          if (hMap.loopArea2 !== undefined) outputRowArray[hMap.loopArea2] = intervalMetrics[1].area; [cite: 1]
        }
        if (intervalMetrics[2]) { [cite: 1]
          if (hMap.speed3 !== undefined) outputRowArray[hMap.speed3] = intervalMetrics[2].speed;[cite: 1]
          if (hMap.comp3 !== undefined) outputRowArray[hMap.comp3] = pvpSlots[2].comp; [cite: 1]
          if (hMap.reb3 !== undefined) outputRowArray[hMap.reb3] = pvpSlots[2].reb; [cite: 1]
        }
        if (hMap.test1Status !== undefined) outputRowArray[hMap.test1Status] = "PASS";[cite: 1]
        if (hMap.test2Status !== undefined) outputRowArray[hMap.test2Status] = "PASS";[cite: 1]
        if (hMap.overallStatus !== undefined) outputRowArray[hMap.overallStatus] = "PASS";[cite: 1]
        
        var dynamicNextRowIndex = sheet.getLastRow() + 1;[cite: 1]
        sheet.getRange(dynamicNextRowIndex, 1, 1, 24).setValues([outputRowArray]);[cite: 1]
        
        var archiveFolder;[cite: 1]
        var archiveFolders = DriveApp.getFoldersByName(archiveFolderName);[cite: 1]
        if (archiveFolders.hasNext()) {[cite: 1]
          archiveFolder = archiveFolders.next();[cite: 1]
        } else {
          archiveFolder = folder.createFolder(archiveFolderName);[cite: 1]
        }

        pack.pvp.moveTo(archiveFolder); [cite: 1]
        for (var f = 0; f < pack.intervals.length; f++) { [cite: 1]
          pack.intervals[f].moveTo(archiveFolder); [cite: 1]
        }
      } catch(err) { Logger.log("Watch folder processing failed: " + err.toString()); }[cite: 1]
    }
  }
}

// =========================================================================
// HISTORICAL RE-RUN ENGINE (STRICT SELF-HEALING UNIFIED SPC BALANCER)
// =========================================================================
function retroactiveLogRecalculate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();[cite: 1]
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);[cite: 1]
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);[cite: 1]
  var regSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);[cite: 1]
  if (!logSheet || !refSheet) return;[cite: 1]
  
  var logRange = logSheet.getDataRange();[cite: 1]
  var logData = logRange.getValues(); if (logData.length < 2) return;[cite: 1]
  var hMap = buildHeaderMap(logData[0]);[cite: 1]
  var refData = refSheet.getDataRange().getValues();[cite: 1]
  var mMap = buildMatrixHeaderMap(refData[0]);[cite: 1]
  
  var mean = function(arr) { var clean = arr.filter(function(x){return !isNaN(x);}); return clean.length === 0 ? 0 : clean.reduce(function(a,b){return a+b;},0)/clean.length; };[cite: 1]
  var sd = function(arr, m) { var clean = arr.filter(function(x){return !isNaN(x);}); return clean.length <= 1 ? 0 : Math.sqrt(clean.map(function(x){return Math.pow(x-m,2);}).reduce(function(a,b){return a+b;},0)/(clean.length - 1)); };[cite: 1]

  var batchGroups = {};[cite: 1]
  var historicalGroups = {};[cite: 1]
  
  // Build BASE_MODEL & PROGRAM_NAME -> DYNAMIC_KEY map from Program_Registry
  var modelToDynamicKey = {};[cite: 1]
  if (regSheet) {[cite: 1]
    var regValues = regSheet.getDataRange().getValues();[cite: 1]
    var regCols = CONFIG.COLUMNS.PROGRAM_REGISTRY;[cite: 1]
    for (var k = 1; k < regValues.length; k++) {[cite: 1]
      var bm = String(regValues[k][regCols.BASE_MODEL - 1] || "").trim().toLowerCase().replace(/[-_\s]/g, "");[cite: 1]
      var pn = String(regValues[k][regCols.PROGRAM_NAME - 1] || "").trim().toLowerCase().replace(/[-_\s]/g, "");[cite: 1]
      var dk = String(regValues[k][regCols.DYNAMIC_KEY - 1] || "").trim();[cite: 1]
      
      var targetKey = dk || regValues[k][regCols.PROGRAM_NAME - 1] || regValues[k][regCols.BASE_MODEL - 1];[cite: 1]
      if (bm) modelToDynamicKey[bm] = targetKey;[cite: 1]
      if (pn) modelToDynamicKey[pn] = targetKey;[cite: 1]
    }
  }

  // Pass 1: Filter & Deduplicate for SPC Baseline Pool
  for (var r = 1; r < logData.length; r++) {[cite: 1]
    var prog = String(logData[r][hMap.programName] || "").trim();[cite: 1]
    var serial = String(logData[r][hMap.trueSerial] || "").trim();[cite: 1]
    var baseModel = String(logData[r][hMap.baseModel] || "").trim().toLowerCase().replace(/[-_\s]/g, "");[cite: 1]
    var cleanProgName = prog.toLowerCase().replace(/[-_\s]/g, "");[cite: 1]
    var overallStatus = String(logData[r][hMap.overallStatus] || "").toUpperCase().trim();[cite: 1]
    if (logData[r][0] === "") continue;[cite: 1]
    
    if (serial !== "") {[cite: 1]
      var batchId = serial.split("-")[0].trim();[cite: 1]
      if (!batchGroups[batchId]) {[cite: 1]
        batchGroups[batchId] = { c1: [], r1: [], c2: [], r2: [], rf: [], rowReferences: [] };[cite: 1]
      }
      batchGroups[batchId].c1.push(parseFloat(logData[r][hMap.comp1]) || 0);[cite: 1]
      batchGroups[batchId].r1.push(Math.abs(parseFloat(logData[r][hMap.reb1])) || 0);[cite: 1]
      batchGroups[batchId].c2.push(parseFloat(logData[r][hMap.comp2]) || 0);[cite: 1]
      batchGroups[batchId].r2.push(Math.abs(parseFloat(logData[r][hMap.reb2])) || 0);[cite: 1]
      batchGroups[batchId].rf.push(parseFloat(logData[r][hMap.rodForce]) || 0);[cite: 1]
      batchGroups[batchId].rowReferences.push(r);[cite: 1]
    }
    
    // Resolve DYNAMIC_KEY via Base Model or Program Name lookup
    var resolvedGroupKey = modelToDynamicKey[baseModel] || modelToDynamicKey[cleanProgName] || prog;[cite: 1]
    var isPassingRun = overallStatus.indexOf("PASS") !== -1;[cite: 1]

    if (resolvedGroupKey !== "" && isPassingRun) {[cite: 1]
      if (!historicalGroups[resolvedGroupKey]) {[cite: 1]
        historicalGroups[resolvedGroupKey] = {};[cite: 1]
      }
      logData[r]._rowIdx = r + 1;[cite: 1]
      
      var serialKey = serial !== "" ? serial.toLowerCase() : ("row_" + (r + 1));[cite: 1]
      historicalGroups[resolvedGroupKey][serialKey] = logData[r];[cite: 1]
    }
  }
  
  // Pass 2: Calculate Baselines across speeds
  for (var pName in historicalGroups) {[cite: 1]
    var serialMap = historicalGroups[pName];[cite: 1]
    var pool = Object.keys(serialMap).map(function(k) { return serialMap[k]; });[cite: 1]
    var countN = pool.length;[cite: 1]
    var cleanPName = pName.replace(/[-_\s]/g, "").toLowerCase();[cite: 1]
    
    var refRowIdx = -1;[cite: 1]
    for (var mx = 1; mx < refData.length; mx++) {[cite: 1]
      var matrixKeyRaw = String(refData[mx][mMap.dynamicKey] || "").trim();[cite: 1]
      if (!matrixKeyRaw) continue;[cite: 1]
      var cleanMatrixKey = matrixKeyRaw.replace(/[-_\s]/g, "").toLowerCase();[cite: 1]
      
      if (cleanMatrixKey === cleanPName || [cite: 1]
         (cleanMatrixKey.length >= 6 && cleanPName.indexOf(cleanMatrixKey) !== -1) || [cite: 1]
         (cleanPName.length >= 6 && cleanMatrixKey.indexOf(cleanPName) !== -1)) {[cite: 1]
        refRowIdx = mx + 1;[cite: 1]
        break;[cite: 1]
      }
    }
    
    if (refRowIdx !== -1) {[cite: 1]
      var row = refData[refRowIdx - 1];[cite: 1]
      var c1MinRaw = row[mMap.c1Min];[cite: 1]
      var isSeeded = c1MinRaw !== "" && c1MinRaw !== null && !isNaN(parseFloat(c1MinRaw));[cite: 1]
      
      var c1Vals = [], r1Vals = [], c2Vals = [], r2Vals = [];[cite: 1]
      var s1Vals = [], s2Vals = [], s3Vals = [];[cite: 1]
      var sl1Vals = [], la1Vals = [], sl2Vals = [];[cite: 1]
      
      for (var s = 0; s < pool.length; s++) {[cite: 1]
        var d = pool[s];[cite: 1]
        c1Vals.push(parseFloat(d[hMap.comp1]) || 0);[cite: 1]
        r1Vals.push(Math.abs(parseFloat(d[hMap.reb1])) || 0);[cite: 1]
        c2Vals.push(parseFloat(d[hMap.comp2]) || 0);[cite: 1]
        r2Vals.push(Math.abs(parseFloat(d[hMap.reb2])) || 0);[cite: 1]
        
        if (!isNaN(parseFloat(d[hMap.slope1]))) sl1Vals.push(parseFloat(d[hMap.slope1]));[cite: 1]
        if (!isNaN(parseFloat(d[hMap.loopArea1]))) la1Vals.push(parseFloat(d[hMap.loopArea1]));[cite: 1]
        if (!isNaN(parseFloat(d[hMap.slope2]))) sl2Vals.push(parseFloat(d[hMap.slope2]));[cite: 1]

        var sp1 = snapToNominalSpeed(d[hMap.speed1]);[cite: 1]
        var sp2 = snapToNominalSpeed(d[hMap.speed2]);[cite: 1]
        var sp3 = snapToNominalSpeed(d[hMap.speed3]);[cite: 1]
        
        if (sp1 > 0) s1Vals.push(sp1);[cite: 1]
        if (sp2 > 0) s2Vals.push(sp2);[cite: 1]
        if (sp3 > 0) s3Vals.push(sp3);[cite: 1]
      }
      
      var rawSp1 = s1Vals.length > 0 ? mean(s1Vals) : 100;[cite: 1]
      var rawSp2 = s2Vals.length > 0 ? mean(s2Vals) : 400;[cite: 1]
      var rawSp3 = s3Vals.length > 0 ? mean(s3Vals) : 1000;[cite: 1]
      
      var snappedSp1 = snapToNominalSpeed(rawSp1);[cite: 1]
      var snappedSp2 = snapToNominalSpeed(rawSp2);[cite: 1]
      var snappedSp3 = snapToNominalSpeed(rawSp3);[cite: 1]
      
      if (snappedSp1 <= 100 || snappedSp3 <= 1000) {[cite: 1]
        snappedSp1 = 100; snappedSp2 = 400; snappedSp3 = 1000;[cite: 1]
      } else {
        snappedSp1 = 500; snappedSp2 = 1000; snappedSp3 = 2500;[cite: 1]
      }
      
      if (mMap.speed1 !== undefined) refSheet.getRange(refRowIdx, mMap.speed1 + 1).setValue(snappedSp1);[cite: 1]
      if (mMap.speed2 !== undefined) refSheet.getRange(refRowIdx, mMap.speed2 + 1).setValue(snappedSp2);[cite: 1]
      if (mMap.speed3 !== undefined) refSheet.getRange(refRowIdx, mMap.speed3 + 1).setValue(snappedSp3);[cite: 1]
      
      for (var s = 0; s < pool.length; s++) {[cite: 1]
        var d = pool[s];[cite: 1]
        var rIdx = d._rowIdx;[cite: 1]
        if (rIdx) {[cite: 1]
          if (d[hMap.speed1] !== snappedSp1) logSheet.getRange(rIdx, hMap.speed1 + 1).setValue(snappedSp1);[cite: 1]
          if (d[hMap.speed2] !== snappedSp2) logSheet.getRange(rIdx, hMap.speed2 + 1).setValue(snappedSp2);[cite: 1]
          if (d[hMap.speed3] !== snappedSp3) logSheet.getRange(rIdx, hMap.speed3 + 1).setValue(snappedSp3);[cite: 1]
        }
      }
      
      var c1M = mean(c1Vals), c1S = sd(c1Vals, c1M);[cite: 1]
      var r1M = mean(r1Vals), r1S = sd(r1Vals, r1M);[cite: 1]
      var c2M = mean(c2Vals), c2S = sd(c2Vals, c2M);[cite: 1]
      var r2M = mean(r2Vals), r2S = sd(r2Vals, r2M);[cite: 1]

      if (!isSeeded) {[cite: 1]
        if (mMap.c1Mean !== undefined) refSheet.getRange(refRowIdx, mMap.c1Mean + 1).setValue(parseFloat(c1M.toFixed(1)));[cite: 1]
        if (mMap.c1SD !== undefined) refSheet.getRange(refRowIdx, mMap.c1SD + 1).setValue(parseFloat(c1S.toFixed(2)));[cite: 1]
        if (mMap.r1Mean !== undefined) refSheet.getRange(refRowIdx, mMap.r1Mean + 1).setValue(parseFloat(r1M.toFixed(1)));[cite: 1]
        if (mMap.r1SD !== undefined) refSheet.getRange(refRowIdx, mMap.r1SD + 1).setValue(parseFloat(r1S.toFixed(2)));[cite: 1]
        if (mMap.c2Mean !== undefined) refSheet.getRange(refRowIdx, mMap.c2Mean + 1).setValue(parseFloat(c2M.toFixed(1)));[cite: 1]
        if (mMap.c2SD !== undefined) refSheet.getRange(refRowIdx, mMap.c2SD + 1).setValue(parseFloat(c2S.toFixed(2)));[cite: 1]
        if (mMap.r2Mean !== undefined) refSheet.getRange(refRowIdx, mMap.r2Mean + 1).setValue(parseFloat(r2M.toFixed(1)));[cite: 1]
        if (mMap.r2SD !== undefined) refSheet.getRange(refRowIdx, mMap.r2SD + 1).setValue(parseFloat(r2S.toFixed(2)));[cite: 1]
        
        if (countN > 2) {[cite: 1]
          if (mMap.c1Min !== undefined) refSheet.getRange(refRowIdx, mMap.c1Min + 1).setValue(Math.max(0, parseFloat((c1M - 3*c1S).toFixed(1))));[cite: 1]
          if (mMap.c1Max !== undefined) refSheet.getRange(refRowIdx, mMap.c1Max + 1).setValue(parseFloat((c1M + 3*c1S).toFixed(1)));[cite: 1]
          if (mMap.r1Min !== undefined) refSheet.getRange(refRowIdx, mMap.r1Min + 1).setValue(Math.max(0, parseFloat((r1M - 3*r1S).toFixed(1))));[cite: 1]
          if (mMap.r1Max !== undefined) refSheet.getRange(refRowIdx, mMap.r1Max + 1).setValue(parseFloat((r1M + 3*r1S).toFixed(1)));[cite: 1]
          if (mMap.c2Min !== undefined) refSheet.getRange(refRowIdx, mMap.c2Min + 1).setValue(Math.max(0, parseFloat((c2M - 3*c2S).toFixed(1))));[cite: 1]
          if (mMap.c2Max !== undefined) refSheet.getRange(refRowIdx, mMap.c2Max + 1).setValue(parseFloat((c2M + 3*c2S).toFixed(1)));[cite: 1]
          if (mMap.r2Min !== undefined) refSheet.getRange(refRowIdx, mMap.r2Min + 1).setValue(Math.max(0, parseFloat((r2M - 3*r2S).toFixed(1))));[cite: 1]
          if (mMap.r2Max !== undefined) refSheet.getRange(refRowIdx, mMap.r2Max + 1).setValue(parseFloat((r2M + 3*r2S).toFixed(1)));[cite: 1]
        }
      }

      // Populate Minimum Slopes & Loop Area Metrics (Columns M, N, W)
      if (sl1Vals.length > 0 && mMap.slope1Min !== undefined) refSheet.getRange(refRowIdx, mMap.slope1Min + 1).setValue(parseFloat(Math.min.apply(null, sl1Vals).toFixed(1)));[cite: 1]
      if (la1Vals.length > 0 && mMap.loopArea1Min !== undefined) refSheet.getRange(refRowIdx, mMap.loopArea1Min + 1).setValue(parseFloat(Math.min.apply(null, la1Vals).toFixed(1)));[cite: 1]
      if (sl2Vals.length > 0 && mMap.slope2Min !== undefined) refSheet.getRange(refRowIdx, mMap.slope2Min + 1).setValue(parseFloat(Math.min.apply(null, sl2Vals).toFixed(1)));[cite: 1]

      // Populate Engineering SPC Metrics (Columns AG to AN)
      var origComp = parseFloat(row[mMap.origCompBase]);[cite: 1]
      if (isNaN(origComp) || origComp === 0) {[cite: 1]
        origComp = c1M;
        if (mMap.origCompBase !== undefined) refSheet.getRange(refRowIdx, mMap.origCompBase + 1).setValue(parseFloat(origComp.toFixed(1)));[cite: 1]
      }

      var origReb = parseFloat(row[mMap.origRebBase]);[cite: 1]
      if (isNaN(origReb) || origReb === 0) {[cite: 1]
        origReb = r1M;
        if (mMap.origRebBase !== undefined) refSheet.getRange(refRowIdx, mMap.origRebBase + 1).setValue(parseFloat(origReb.toFixed(1)));[cite: 1]
      }

      if (mMap.rawCompSD !== undefined) refSheet.getRange(refRowIdx, mMap.rawCompSD + 1).setValue(parseFloat(c1S.toFixed(2)));[cite: 1]
      if (mMap.rawRebSD !== undefined) refSheet.getRange(refRowIdx, mMap.rawRebSD + 1).setValue(parseFloat(r1S.toFixed(2)));[cite: 1]

      var cMin = parseFloat(row[mMap.c1Min]), cMax = parseFloat(row[mMap.c1Max]);[cite: 1]
      var rMin = parseFloat(row[mMap.r1Min]), rMax = parseFloat(row[mMap.r1Max]);[cite: 1]
      var cWidth = (!isNaN(cMin) && !isNaN(cMax)) ? (cMax - cMin) : (6 * c1S);[cite: 1]
      var rWidth = (!isNaN(rMin) && !isNaN(rMax)) ? (rMax - rMin) : (6 * r1S);[cite: 1]

      if (mMap.compRangeWidth !== undefined) refSheet.getRange(refRowIdx, mMap.compRangeWidth + 1).setValue(parseFloat(cWidth.toFixed(1)));[cite: 1]
      if (mMap.rebRangeWidth !== undefined) refSheet.getRange(refRowIdx, mMap.rebRangeWidth + 1).setValue(parseFloat(rWidth.toFixed(1)));[cite: 1]

      var compDrift = origComp !== 0 ? (((c1M - origComp) / origComp) * 100) : 0;[cite: 1]
      var rebDrift  = origReb !== 0 ? (((r1M - origReb) / origReb) * 100) : 0;[cite: 1]

      if (mMap.compDriftPct !== undefined) refSheet.getRange(refRowIdx, mMap.compDriftPct + 1).setValue(parseFloat(compDrift.toFixed(2)));[cite: 1]
      if (mMap.rebDriftPct !== undefined) refSheet.getRange(refRowIdx, mMap.rebDriftPct + 1).setValue(parseFloat(rebDrift.toFixed(2)));[cite: 1]
      
      // Update Column AO: Health Stamp
      var maxDrift = Math.max(Math.abs(compDrift), Math.abs(rebDrift));[cite: 1]
      var procHealth = "";[cite: 1]
      if (maxDrift > 10) {[cite: 1]
        procHealth = "🔴 WARNING: UNACCEPTABLE DRIFT (" + maxDrift.toFixed(1) + "%)";[cite: 1]
      } else if (isSeeded) {[cite: 1]
        procHealth = "🟢 Stage 0: SEEDED BLUEPRINT ACTIVE (3σ)";[cite: 1]
      } else if (countN >= 100) {[cite: 1]
        procHealth = "🟢 Stage 4: MATURE SPC LOCKED (3σ)";[cite: 1]
      } else {
        procHealth = "🟡 Stage 1: BUILDING BASELINE (" + countN + " samples)";[cite: 1]
      }
      if (mMap.healthStamp !== undefined) refSheet.getRange(refRowIdx, mMap.healthStamp + 1).setValue(procHealth);[cite: 1]

      // Update Column AP: Control Mode
      var controlModeText = "MANUAL GRACE LIMITS LOADED";[cite: 1]
      if (countN >= 10 && !isSeeded) {[cite: 1]
        controlModeText = "AUTOMATED STATISTICAL SPC LAYER ACTIVE";[cite: 1]
      } else if (countN >= 10 && isSeeded) {[cite: 1]
        controlModeText = "HYBRID: SEEDED BLUEPRINT WITH ACTIVE SPC";[cite: 1]
      }
      if (mMap.controlMode !== undefined) refSheet.getRange(refRowIdx, mMap.controlMode + 1).setValue(controlModeText);[cite: 1]

      if (mMap.sampleCount !== undefined) refSheet.getRange(refRowIdx, mMap.sampleCount + 1).setValue(countN);[cite: 1]
    }
  }

  // Pass 3: Rolling Batch Math Limits
  var batchStats = {};[cite: 1]
  for (var bId in batchGroups) {[cite: 1]
    var b = batchGroups[bId];[cite: 1]
    var mC1 = mean(b.c1), sC1 = sd(b.c1, mC1);[cite: 1]
    var mR1 = mean(b.r1), sR1 = sd(b.r1, mR1);[cite: 1]
    var mC2 = mean(b.c2), sC2 = sd(b.c2, mC2);[cite: 1]
    var mR2 = mean(b.r2), sR2 = sd(b.r2, mR2);[cite: 1]
    var mRF = mean(b.rf), sRF = sd(b.rf, mRF);[cite: 1]
    
    batchStats[bId] = {
      c1Min: mC1 - 2*sC1, c1Max: mC1 + 2*sC1,[cite: 1]
      r1Min: mR1 - 2*sR1, r1Max: mR1 + 2*sR1,[cite: 1]
      c2Min: mC2 - 2*sC2, c2Max: mC2 + 2*sC2,[cite: 1]
      r2Min: mR2 - 2*sR2, r2Max: mR2 + 2*sR2,[cite: 1]
      rfMin: mRF - 2*sRF, count: b.c1.length[cite: 1]
    };
  }

  // Pass 4: In-Memory Multi-Gate Diagnostic Tag Fingerprinting
  var qualityOutputSubMatrix = [];[cite: 1]
  for (var r = 1; r < logData.length; r++) {[cite: 1]
    var pName = String(logData[r][hMap.programName] || "").trim();[cite: 1]
    var serial = String(logData[r][hMap.trueSerial] || "").trim();[cite: 1]
    var baseModel = String(logData[r][hMap.baseModel] || "").trim().toLowerCase().replace(/[-_\s]/g, "");[cite: 1]
    var cleanPName = pName.toLowerCase().replace(/[-_\s]/g, "");[cite: 1]
    var batchId = serial.split("-")[0].trim();[cite: 1]
    
    var test1Result = "INITIALIZING"; var test2Result = "INITIALIZING"; var finalStatus = "PASS";[cite: 1]
    
    // Dynamically fetch evaluation action and comments regardless of Column W or Z positioning
    var evalAction = String(logData[r][hMap.evaluationAction] || "").trim();
    var engComm = String(logData[r][hMap.engComments] || "").trim();
    var failTags = [];[cite: 1]
    
    if (pName && logData[r][0] !== "") {[cite: 1]
      var resolvedDynamicKey = modelToDynamicKey[baseModel] || modelToDynamicKey[cleanPName] || pName;[cite: 1]
      var cleanResolvedKey = resolvedDynamicKey.toLowerCase().replace(/[-_\s]/g, "");[cite: 1]
      
      var refRow = null;[cite: 1]
      for (var mx = 1; mx < refData.length; mx++) {[cite: 1]
        var mKey = String(refData[mx][mMap.dynamicKey] || "").trim().toLowerCase().replace(/[-_\s]/g, "");[cite: 1]
        if (mKey && (mKey === cleanResolvedKey || cleanResolvedKey.indexOf(mKey) !== -1 || mKey.indexOf(cleanResolvedKey) !== -1)) {[cite: 1]
          refRow = refData[mx]; break;[cite: 1]
        }
      }
      
      var lowC = parseFloat(logData[r][hMap.comp1]) || 0; var lowR = Math.abs(parseFloat(logData[r][hMap.reb1])) || 0;[cite: 1]
      var midC = parseFloat(logData[r][hMap.comp2]) || 0; var midR = Math.abs(parseFloat(logData[r][hMap.reb2])) || 0;[cite: 1]
      var sl1  = parseFloat(logData[r][hMap.slope1]) || 0; var individualRF = parseFloat(logData[r][hMap.rodForce]) || 0;[cite: 1]
      
      // Global Limit Check Tagging (Test 1)
      if (refRow) {[cite: 1]
        var t1Pass = true;[cite: 1]
        var valC1Min = parseFloat(refRow[mMap.c1Min]); var valC1Max = parseFloat(refRow[mMap.c1Max]);[cite: 1]
        var valR1Min = parseFloat(refRow[mMap.r1Min]); var valR1Max = parseFloat(refRow[mMap.r1Max]);[cite: 1]
        var valC2Min = parseFloat(refRow[mMap.c2Min]); var valC2Max = parseFloat(refRow[mMap.c2Max]);[cite: 1]
        var valR2Min = parseFloat(refRow[mMap.r2Min]); var valR2Max = parseFloat(refRow[mMap.r2Max]);[cite: 1]
        var slope1Min = parseFloat(refRow[mMap.slope1Min]) || 0;[cite: 1]
        
        if (!isNaN(valC1Min)) {[cite: 1]
          if (lowC < valC1Min || lowC > valC1Max) { t1Pass = false; failTags.push("[C1_FAIL]"); }[cite: 1]
          if (lowR < valR1Min || lowR > valR1Max) { t1Pass = false; failTags.push("[R1_FAIL]"); }[cite: 1]
          if (midC < valC2Min || midC > valC2Max) { t1Pass = false; failTags.push("[C2_FAIL]"); }[cite: 1]
          if (midR < valR2Min || midR > valR2Max) { t1Pass = false; failTags.push("[R2_FAIL]"); }[cite: 1]
          if (sl1 < slope1Min) { t1Pass = false; failTags.push("[SLOPE_FAIL]"); }[cite: 1]
        }
        test1Result = t1Pass ? "PASS" : "FAIL (BLUEPRINT)";[cite: 1]
      }
      
      // Cohort Outlier Check Tagging (Test 2)
      var cStat = batchStats[batchId];[cite: 1]
      var defectAnalysis = "";[cite: 1]
      if (cStat && cStat.count > 2) {[cite: 1]
        var t2Pass = true;[cite: 1]
        var lowGasPressure = (individualRF < cStat.rfMin);[cite: 1]
        var lowC1 = (lowC < cStat.c1Min || lowC > cStat.c1Max);[cite: 1]
        var lowR1 = (lowR < cStat.r1Min || lowR > cStat.r1Max);[cite: 1]
        var lowC2 = (midC < cStat.c2Min || midC > cStat.c2Max);[cite: 1]
        var lowR2 = (midR < cStat.r2Min || midR > cStat.r2Max);[cite: 1]
        
        if (lowGasPressure) { t2Pass = false; failTags.push("[RF_FAIL]"); }[cite: 1]
        if (lowC1) { t2Pass = false; failTags.push("[C1_FAIL]"); }[cite: 1]
        if (lowR1) { t2Pass = false; failTags.push("[R1_FAIL]"); }[cite: 1]
        if (lowC2) { t2Pass = false; failTags.push("[C2_FAIL]"); }[cite: 1]
        if (lowR2) { t2Pass = false; failTags.push("[R2_FAIL]"); }[cite: 1]
        
        if (!t2Pass) {[cite: 1]
          if (lowGasPressure) defectAnalysis = "Gas Pressure Deficient.";[cite: 1]
          else if (lowC1 && lowR1) defectAnalysis = "Symmetric drop. Potential bypass.";[cite: 1]
          else if (lowC1) defectAnalysis = "Compression outlier variation.";[cite: 1]
          else if (lowR1) defectAnalysis = "Rebound outlier variation.";[cite: 1]
          else defectAnalysis = "Outlier variance detected.";[cite: 1]
        }
        test2Result = t2Pass ? "PASS" : "FAIL (OUTLIER)";[cite: 1]
      }
      
      var uniqueFailTags = [];[cite: 1]
      for (var f = 0; f < failTags.length; f++) {[cite: 1]
        if (uniqueFailTags.indexOf(failTags[f]) === -1) uniqueFailTags.push(failTags[f]);[cite: 1]
      }
      
      var cleanExt = (evalAction + " " + engComm).toLowerCase();
      var globalPass = (test1Result === "INITIALIZING" || !test1Result.includes("FAIL")) && (test2Result === "INITIALIZING" || !test2Result.includes("FAIL"));[cite: 1]
      var diagnosticNotes = globalPass ? "✅ SHOCK IS WITHIN TOLERANCE." : "❌ ERROR: " + uniqueFailTags.join(" ") + " | " + defectAnalysis;
      
      if (cleanExt.includes("approved") || cleanExt.includes("management") || cleanExt.includes("override")) {
        test1Result = "PASS (OVERRIDE)"; test2Result = "PASS (OVERRIDE)"; finalStatus = "PASS (OVERRIDE)";[cite: 1]
        diagnosticNotes = "👔 DISCRETIONARY CLEAR: Released via Management Sign-off.";[cite: 1]
      } else if (cleanExt.includes("no issue found") || cleanExt.includes("re-tested pass") || cleanExt.includes("validated")) {
        test1Result = "PASS"; test2Result = "PASS"; finalStatus = "PASS";[cite: 1]
        diagnosticNotes = "🛠️ TEARDOWN VALIDATED: Assembly clear.";[cite: 1]
      } else if (!globalPass) {
        finalStatus = "FAIL";[cite: 1]
      }
    }
    qualityOutputSubMatrix.push([test1Result, test2Result, finalStatus, diagnosticNotes]);[cite: 1]
  }
  
  if (qualityOutputSubMatrix.length > 0) {[cite: 1]
    var startColIdx = (CONFIG.COLUMNS.MASTER_DYNO_LOG.TEST_1_STATUS) || 19;[cite: 1]
    logSheet.getRange(2, startColIdx, qualityOutputSubMatrix.length, 4).setValues(qualityOutputSubMatrix);[cite: 1]
  }
  SpreadsheetApp.flush();[cite: 1]
}