// =========================================================================
// 🏎️ UNIFIED DYNO ENGINE (Engine.gs)
// Adaptive Multi-File Ingestion, Hysteresis Regression & Staged SPC Engine
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

  for (var i = 0; i < headerRow.length; i++) {
    var cleanH = String(headerRow[i] || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    if (cleanH.indexOf("diagnostic") !== -1) continue;

    if (cleanH.indexOf("evaluation") !== -1 || cleanH === "action" || cleanH.indexOf("evalaction") !== -1) {
      map.evaluationAction = i;
    } else if (cleanH.indexOf("engcomment") !== -1 || cleanH.indexOf("engineeringcomment") !== -1 || (cleanH.indexOf("comment") !== -1 && cleanH.indexOf("diagnostic") === -1)) {
      map.engComments = i;
    }
  }
  
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
    loopArea1Min: refCols.LOOP_AREA_1_MIN - 1,
    c2Mean: refCols.COMP_2_MEAN - 1,
    c2SD: refCols.COMP_2_SD - 1,
    c2Min: refCols.COMP_2_MIN - 1,
    c2Max: refCols.COMP_2_MAX - 1,
    r2Mean: refCols.REB_2_MEAN - 1,
    r2SD: refCols.REB_2_SD - 1,
    r2Min: refCols.REB_2_MIN - 1,
    r2Max: refCols.REB_2_MAX - 1,
    slope2Min: refCols.SLOPE_2_MIN - 1,
    c3Mean: refCols.COMP_3_MEAN - 1,
    c3SD: refCols.COMP_3_SD - 1,
    c3Min: refCols.COMP_3_MIN - 1,
    c3Max: refCols.COMP_3_MAX - 1,
    r3Mean: refCols.REB_3_MEAN - 1,
    r3SD: refCols.REB_3_SD - 1,
    r3Min: refCols.REB_3_MIN - 1,
    r3Max: refCols.REB_3_MAX - 1,
    sampleCount: refCols.SAMPLE_COUNT - 1,
    rawCompSD: refCols.RAW_COMP_SD - 1,
    rawRebSD: refCols.RAW_REB_SD - 1,
    origCompBase: refCols.ORIGINAL_COMP_BASE - 1,
    origRebBase: refCols.ORIGINAL_REB_BASE - 1,
    compRangeWidth: refCols.COMP_RANGE_WIDTH - 1,
    rebRangeWidth: refCols.REB_RANGE_WIDTH - 1,
    compDriftPct: refCols.COMP_DRIFT_PCT - 1,
    rebDriftPct: refCols.REB_DRIFT_PCT - 1,
    healthStamp: refCols.HEALTH_STAMP - 1,
    controlMode: refCols.CONTROL_MODE - 1
  };
}

/**
 * Robust Delimiter-Smart CSV/TSV Parser Helper
 */
function parseCsvRobust(rawStr) {
  if (!rawStr) return [];
  var clean = rawStr.trim();
  if (!clean) return [];
  var lines = clean.split(/\r?\n/);
  var rows = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    var delim = line.indexOf('\t') !== -1 ? '\t' : (line.indexOf(';') !== -1 ? ';' : ',');
    var parts = line.split(delim).map(function(cell) {
      return cell.trim().replace(/^["']|["']$/g, '');
    });
    rows.push(parts);
  }
  return rows;
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
    var file = files.next(); 
    var rawName = file.getName(); 
    var lastUpdated = file.getLastUpdated().getTime();
    if (rawName.toLowerCase().endsWith('.csv') || rawName.toLowerCase().endsWith('.txt')) {
      var fileName = rawName.replace(/\.csv$/i, "").replace(/\.txt$/i, "").trim(); 
      var tempGroupKey = "";
      var serialMatch = rawName.match(/\d{6}-\d{3}/) || rawName.match(/\d+-\d+/);
      if (serialMatch) { 
        tempGroupKey = serialMatch[0].trim(); 
      } else {
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
        var rawParts = fileName.split(/[\s_]+/); 
        detectedModel = rawParts[0].replace('_PROCESSING_', '');
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
  
  var registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY); 
  var registryData = registrySheet ? registrySheet.getDataRange().getValues() : [];
  var masterLogData = sheet.getDataRange().getValues();
  var hMap = buildHeaderMap(masterLogData[0]);
  
  for (var tempKey in tempFileMap) {
    var pack = tempFileMap[tempKey];
    if (pack.pvp && pack.intervals.length >= 3) {
      if ((now.getTime() - pack.youngestFileTime) / 1000 < 4) continue;
      try {
        var rawBlobStr = pack.pvp.getBlob().getDataAsString("UTF-8").trim();
        var pvpContent = parseCsvRobust(rawBlobStr); 
        var trueSerial = tempKey;
        var trueDynoProgramName = pack.parsedBaseModel;
        
        if (trueDynoProgramName.length > 20) {
          trueDynoProgramName = trueDynoProgramName.substring(0, 14) + "_V" + trueDynoProgramName.slice(-1);
        }
        
        var cleanTrueProgName = trueDynoProgramName.replace(/[-_\s]/g, "").toLowerCase();
        for (var rR = 1; rR < registryData.length; rR++) {
          var regKey = String(registryData[rR][0] || "").trim().replace(/[-_\s]/g, "").toLowerCase();
          if (regKey === cleanTrueProgName) {
            trueDynoProgramName = String(registryData[rR][0] || "").trim(); 
            break;
          }
        }
        
        var registryBaseModelText = trueDynoProgramName; 
        var registryValvingVersion = "PRODUCTION_RUN";
        var cleanMatchName = trueDynoProgramName.replace(/[-_\s]/g, "").toLowerCase();
        for (var regR = 1; regR < registryData.length; regR++) {
          var cleanRegCell = String(registryData[regR][0] || "").trim().replace(/[-_\s]/g, "").toLowerCase();
          if (cleanRegCell === cleanMatchName) {
            registryBaseModelText = String(registryData[regR][2] || "").trim();
            registryValvingVersion = String(registryData[regR][6] || "").trim();
            break;
          }
        }
        
        var rawRodForce = 0; 
        var intervalMetrics = [];
        for (var i = 0; i < pack.intervals.length; i++) {
          var file = pack.intervals[i]; 
          var rawIntStr = file.getBlob().getDataAsString("UTF-8").trim();
          if (rawIntStr === "") continue; 
          var rows = parseCsvRobust(rawIntStr); 
          var speedTarget = 0;
          for (var j = 0; j < Math.min(rows.length, 15); j++) { 
            if (rows[j][0] && rows[j][0].trim().toLowerCase() == "velocity amplitude") { 
              speedTarget = parseInt(rows[j][1]); 
              break; 
            } 
          }
          
          if (!speedTarget || isNaN(speedTarget)) {
            for (var r = 14; r < rows.length; r++) {
              if (rows[r][0] && !isNaN(parseFloat(rows[r][0]))) {
                var testV = Math.abs(parseFloat(rows[r][0]));
                if (testV > speedTarget) speedTarget = testV;
              }
            }
          }
          speedTarget = snapToNominalSpeed(speedTarget);
          
          for (var j = 0; j < rows.length; j++) { 
            if (rows[j][0] && rows[j][0].trim().toLowerCase() == "rod force") { 
              rawRodForce = parseFloat(rows[j][1]); 
              break; 
            } 
          }
          var maxComp = -9999, maxReb = 9999, totalArea = 0, nearZeroPoints = [];
          for (var r = 14; r < rows.length; r++) {
            if (rows[r].length < 2 || isNaN(parseFloat(rows[r][0]))) continue;
            var vel = parseFloat(rows[r][0]); 
            var force = parseFloat(rows[r][1]);
            if (force > maxComp) maxComp = force; 
            if (force < maxReb) maxReb = force;
            totalArea += Math.abs(force); 
            if (Math.abs(vel) < ((speedTarget || 500) * 0.15)) nearZeroPoints.push({ x: vel, y: force });
          }
          var sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
          for (var p = 0; p < nearZeroPoints.length; p++) { 
            sumX += nearZeroPoints[p].x; 
            sumY += nearZeroPoints[p].y; 
            sumXY += (nearZeroPoints[p].x * nearZeroPoints[p].y); 
            var varXX = (nearZeroPoints[p].x * nearZeroPoints[p].x); 
            sumXX += varXX; 
          }
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
        var numericRows = [];
        for (var pr = 0; pr < pvpContent.length; pr++) {
          if (pvpContent[pr].length >= 4) {
            var valComp = parseFloat(pvpContent[pr][3]);
            var valReb = parseFloat(pvpContent[pr][5] || pvpContent[pr][4]);
            if (!isNaN(valComp) && valComp !== 0) {
              numericRows.push({ comp: Math.abs(valComp), reb: Math.abs(valReb || 0) });
            }
          }
        }
        if (numericRows.length >= 3) {
          pvpSlots = numericRows.slice(0, 3);
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
          var c1 = pvpSlots[0].comp || Math.abs(intervalMetrics[0].maxComp) || 0;
          var r1 = pvpSlots[0].reb || Math.abs(intervalMetrics[0].maxReb) || 0;
          if (hMap.speed1 !== undefined) outputRowArray[hMap.speed1] = intervalMetrics[0].speed;
          if (hMap.comp1 !== undefined) outputRowArray[hMap.comp1] = c1;
          if (hMap.reb1 !== undefined) outputRowArray[hMap.reb1] = r1;
          if (hMap.slope1 !== undefined) outputRowArray[hMap.slope1] = intervalMetrics[0].slope;
          if (hMap.loopArea1 !== undefined) outputRowArray[hMap.loopArea1] = intervalMetrics[0].area;
        }
        if (intervalMetrics[1]) {
          var c2 = pvpSlots[1].comp || Math.abs(intervalMetrics[1].maxComp) || 0;
          var r2 = pvpSlots[1].reb || Math.abs(intervalMetrics[1].maxReb) || 0;
          if (hMap.speed2 !== undefined) outputRowArray[hMap.speed2] = intervalMetrics[1].speed;
          if (hMap.comp2 !== undefined) outputRowArray[hMap.comp2] = c2;
          if (hMap.reb2 !== undefined) outputRowArray[hMap.reb2] = r2;
          if (hMap.loopArea2 !== undefined) outputRowArray[hMap.loopArea2] = intervalMetrics[1].area;
        }
        if (intervalMetrics[2]) {
          var c3 = pvpSlots[2].comp || Math.abs(intervalMetrics[2].maxComp) || 0;
          var r3 = pvpSlots[2].reb || Math.abs(intervalMetrics[2].maxReb) || 0;
          if (hMap.speed3 !== undefined) outputRowArray[hMap.speed3] = intervalMetrics[2].speed;
          if (hMap.comp3 !== undefined) outputRowArray[hMap.comp3] = c3;
          if (hMap.reb3 !== undefined) outputRowArray[hMap.reb3] = r3;
        }
        if (hMap.test1Status !== undefined) outputRowArray[hMap.test1Status] = "INITIALIZING";
        if (hMap.test2Status !== undefined) outputRowArray[hMap.test2Status] = "INITIALIZING";
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
// HISTORICAL RE-RUN ENGINE (ADAPTIVE SPC, MEMORY BATCHING & RATCHET)
// =========================================================================
function retroactiveLogRecalculate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);
  var regSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRAM_REGISTRY);
  if (!logSheet || !refSheet) return;
  
  var logRange = logSheet.getDataRange();
  var logData = logRange.getValues(); 
  if (logData.length < 2) return;
  var hMap = buildHeaderMap(logData[0]);
  var refData = refSheet.getDataRange().getValues();
  var mMap = buildMatrixHeaderMap(refData[0]);
  
  var mean = function(arr) { 
    var clean = arr.filter(function(x){ return !isNaN(x); }); 
    return clean.length === 0 ? 0 : clean.reduce(function(a,b){ return a+b; },0)/clean.length; 
  };
  var sd = function(arr, m) { 
    var clean = arr.filter(function(x){ return !isNaN(x); }); 
    return clean.length <= 1 ? 0 : Math.sqrt(clean.map(function(x){ return Math.pow(x-m,2); }).reduce(function(a,b){ return a+b; },0)/(clean.length - 1)); 
  };
  var cleanStr = function(s) { 
    return String(s || "").trim().toLowerCase().replace(/[-_\s]/g, ""); 
  };

  // Helper to safely write to the in-memory array
  function setRefVal(mx, colKey, val) {
    if (mMap[colKey] !== undefined) refData[mx][mMap[colKey]] = val;
  }

  // STEP 0: Build Mapping Dictionary & Seed Matrix
  var modelToDynamicKey = {};
  var existingMatrixKeys = {};
  for (var m = 1; m < refData.length; m++) {
    var mk = cleanStr(refData[m][mMap.dynamicKey]);
    if (mk) existingMatrixKeys[mk] = m + 1;
  }
  
  var rowsToAppend = [];
  if (regSheet) {
    var regValues = regSheet.getDataRange().getValues();
    var regCols = CONFIG.COLUMNS.PROGRAM_REGISTRY;
    for (var k = 1; k < regValues.length; k++) {
      var bm = cleanStr(regValues[k][regCols.BASE_MODEL - 1]);
      var pn = String(regValues[k][regCols.PROGRAM_NAME - 1] || "").trim();
      var pnClean = cleanStr(pn);
      var dk = String(regValues[k][regCols.DYNAMIC_KEY - 1] || "").trim();
      
      var targetKey = dk || pn || regValues[k][regCols.BASE_MODEL - 1];
      var cleanTargetKey = cleanStr(targetKey);

      if (bm) modelToDynamicKey[bm] = targetKey;
      if (pnClean) modelToDynamicKey[pnClean] = targetKey;

      if (cleanTargetKey && !existingMatrixKeys[cleanTargetKey]) {
        var newRefRow = new Array(42).fill("");
        newRefRow[mMap.dynamicKey] = targetKey;
        rowsToAppend.push(newRefRow);
        existingMatrixKeys[cleanTargetKey] = refData.length + rowsToAppend.length;
      }
    }
  }

  if (rowsToAppend.length > 0) {
    refData = refData.concat(rowsToAppend);
  }

  // STEP 1: Progressive Staged Learning & Limit Refining (Memory Batched)
  for (var mx = 1; mx < refData.length; mx++) {
    var matrixKey = String(refData[mx][mMap.dynamicKey] || "").trim();
    if (!matrixKey) continue;
    var cMatrixKey = cleanStr(matrixKey);

    var pool = [];
    for (var r = 1; r < logData.length; r++) {
      var logProg = String(logData[r][hMap.programName] || "").trim();
      var logBaseModel = String(logData[r][hMap.baseModel] || "").trim();
      var cResolvedKey = cleanStr(modelToDynamicKey[cleanStr(logBaseModel)] || modelToDynamicKey[cleanStr(logProg)] || logProg);

      if (cResolvedKey === cMatrixKey || cleanStr(logProg) === cMatrixKey || cleanStr(logBaseModel) === cMatrixKey) {
        logData[r]._rowIdx = r + 1;
        pool.push(logData[r]);
      }
    }
    if (pool.length === 0) continue;

    var countN = pool.length;
    var row = refData[mx];

    // Filter training pool to valid passing runs (Exclude Prototypes and Failures)
    var passPool = [];
    for (var pIdx = 0; pIdx < pool.length; pIdx++) {
      var pRow = pool[pIdx];
      var pStat = String(pRow[hMap.overallStatus] || pRow[hMap.test1Status] || "").toUpperCase();
      var pAction = String(pRow[hMap.evaluationAction] || "").toLowerCase();
      var pSerial = String(pRow[hMap.trueSerial] || "").toLowerCase();

      var isProto = pAction.includes("proto") || pAction.includes("not eval") || pStat.includes("PROTO") || pSerial.includes("proto") || pSerial.includes("dev");

      if (!pStat.includes("FAIL") && !pAction.includes("teardown") && !pAction.includes("issue found") && !isProto) {
        passPool.push(pRow);
      }
    }
    var trainingPool = (passPool.length >= 3) ? passPool : pool;
    if (trainingPool.length > 100) trainingPool = trainingPool.slice(-100);

    var c1Vals = [], r1Vals = [], c2Vals = [], r2Vals = [];
    var s1Vals = [], s2Vals = [], s3Vals = [];
    var sl1Vals = [], la1Vals = [], sl2Vals = [];

    for (var s = 0; s < trainingPool.length; s++) {
      var d = trainingPool[s];
      c1Vals.push(parseFloat(d[hMap.comp1]) || 0); 
      r1Vals.push(Math.abs(parseFloat(d[hMap.reb1])) || 0);
      c2Vals.push(parseFloat(d[hMap.comp2]) || 0); 
      r2Vals.push(Math.abs(parseFloat(d[hMap.reb2])) || 0);

      if (!isNaN(parseFloat(d[hMap.slope1]))) sl1Vals.push(parseFloat(d[hMap.slope1]));
      if (!isNaN(parseFloat(d[hMap.loopArea1]))) la1Vals.push(parseFloat(d[hMap.loopArea1]));
      if (!isNaN(parseFloat(d[hMap.slope2]))) sl2Vals.push(parseFloat(d[hMap.slope2]));

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

    setRefVal(mx, 'speed1', snappedSp1); 
    setRefVal(mx, 'speed2', snappedSp2); 
    setRefVal(mx, 'speed3', snappedSp3);

    var c1M = mean(c1Vals), c1S = sd(c1Vals, c1M); 
    var r1M = mean(r1Vals), r1S = sd(r1Vals, r1M);
    var c2M = mean(c2Vals), c2S = sd(c2Vals, c2M); 
    var r2M = mean(r2Vals), r2S = sd(r2Vals, r2M);

    // Progressive SPC Stage Tightening
    var stageNum = 1; 
    var sigmaMult = 3.0; 
    var pctCap = 0.15;
    if (countN >= 100) { 
      stageNum = 4; sigmaMult = 2.0; pctCap = 0.08; 
    } else if (countN >= 40) { 
      stageNum = 3; sigmaMult = 2.2; pctCap = 0.10; 
    } else if (countN >= 15) { 
      stageNum = 2; sigmaMult = 2.5; pctCap = 0.12; 
    }

    function calcAdaptiveLimits(meanVal, sdVal) {
      var effSpread = (sdVal > 0) ? Math.min(sigmaMult * sdVal, meanVal * pctCap) : (meanVal * pctCap);
      if (effSpread < meanVal * 0.03) effSpread = meanVal * 0.03; // 3% safety buffer
      return { min: parseFloat((meanVal - effSpread).toFixed(1)), max: parseFloat((meanVal + effSpread).toFixed(1)) };
    }

    var ctrlModeRaw = String(row[mMap.controlMode] || "").toUpperCase();
    var isManuallyLocked = ctrlModeRaw.includes("LOCK") || ctrlModeRaw.includes("MANUAL");

    if (!isManuallyLocked && trainingPool.length >= 3) {
      setRefVal(mx, 'c1Mean', parseFloat(c1M.toFixed(1))); 
      setRefVal(mx, 'c1SD', parseFloat(c1S.toFixed(2)));
      setRefVal(mx, 'r1Mean', parseFloat(r1M.toFixed(1))); 
      setRefVal(mx, 'r1SD', parseFloat(r1S.toFixed(2)));
      setRefVal(mx, 'c2Mean', parseFloat(c2M.toFixed(1))); 
      setRefVal(mx, 'c2SD', parseFloat(c2S.toFixed(2)));
      setRefVal(mx, 'r2Mean', parseFloat(r2M.toFixed(1))); 
      setRefVal(mx, 'r2SD', parseFloat(r2S.toFixed(2)));

      var lC1 = calcAdaptiveLimits(c1M, c1S); 
      var lR1 = calcAdaptiveLimits(r1M, r1S);
      var lC2 = calcAdaptiveLimits(c2M, c2S); 
      var lR2 = calcAdaptiveLimits(r2M, r2S);

      setRefVal(mx, 'c1Min', lC1.min); setRefVal(mx, 'c1Max', lC1.max);
      setRefVal(mx, 'r1Min', lR1.min); setRefVal(mx, 'r1Max', lR1.max);
      setRefVal(mx, 'c2Min', lC2.min); setRefVal(mx, 'c2Max', lC2.max);
      setRefVal(mx, 'r2Min', lR2.min); setRefVal(mx, 'r2Max', lR2.max);
    }

    if (sl1Vals.length > 0) setRefVal(mx, 'slope1Min', parseFloat(Math.min.apply(null, sl1Vals).toFixed(1)));
    if (la1Vals.length > 0) setRefVal(mx, 'loopArea1Min', parseFloat(Math.min.apply(null, la1Vals).toFixed(1)));
    if (sl2Vals.length > 0) setRefVal(mx, 'slope2Min', parseFloat(Math.min.apply(null, sl2Vals).toFixed(1)));

    var origComp = parseFloat(row[mMap.origCompBase]);
    if (isNaN(origComp) || origComp === 0) { 
      origComp = c1M; 
      setRefVal(mx, 'origCompBase', parseFloat(origComp.toFixed(1))); 
    }
    var origReb = parseFloat(row[mMap.origRebBase]);
    if (isNaN(origReb) || origReb === 0) { 
      origReb = r1M; 
      setRefVal(mx, 'origRebBase', parseFloat(origReb.toFixed(1))); 
    }

    setRefVal(mx, 'rawCompSD', parseFloat(c1S.toFixed(2))); 
    setRefVal(mx, 'rawRebSD', parseFloat(r1S.toFixed(2)));

    var cWidth = (!isNaN(refData[mx][mMap.c1Max]) && !isNaN(refData[mx][mMap.c1Min])) ? (refData[mx][mMap.c1Max] - refData[mx][mMap.c1Min]) : (6 * c1S);
    var rWidth = (!isNaN(refData[mx][mMap.r1Max]) && !isNaN(refData[mx][mMap.r1Min])) ? (refData[mx][mMap.r1Max] - refData[mx][mMap.r1Min]) : (6 * r1S);
    setRefVal(mx, 'compRangeWidth', parseFloat(cWidth.toFixed(1))); 
    setRefVal(mx, 'rebRangeWidth', parseFloat(rWidth.toFixed(1)));

    var compDrift = (origComp !== 0 && !isNaN(origComp)) ? (((c1M - origComp) / origComp) * 100) : 0;
    var rebDrift  = (origReb !== 0 && !isNaN(origReb)) ? (((r1M - origReb) / origReb) * 100) : 0;
    setRefVal(mx, 'compDriftPct', parseFloat(compDrift.toFixed(2))); 
    setRefVal(mx, 'rebDriftPct', parseFloat(rebDrift.toFixed(2)));

    var maxDrift = Math.max(Math.abs(compDrift), Math.abs(rebDrift));
    var procHealth = "";
    if (maxDrift >= 8) {
      procHealth = "🔴 CRITICAL DRIFT (" + (Math.abs(rebDrift) >= Math.abs(compDrift) ? "Reb" : "Comp") + " vs Origin)";
    } else if (maxDrift >= 4) {
      procHealth = "🟡 MODERATE DRIFT (" + (Math.abs(rebDrift) >= Math.abs(compDrift) ? "Reb" : "Comp") + " vs Origin)";
    } else {
      procHealth = "🟢 STABLE (" + maxDrift.toFixed(1) + "% drift)";
    }
    
    setRefVal(mx, 'healthStamp', procHealth);
    setRefVal(mx, 'controlMode', isManuallyLocked ? "MANUAL LOCKED SPEC" : ("ADAPTIVE SPC (Stage " + stageNum + " | " + sigmaMult + "σ | N=" + countN + ")"));
    setRefVal(mx, 'sampleCount', countN);
  }

  // MASS BATCH WRITE
  refSheet.getRange(1, 1, refData.length, refData[0].length).setValues(refData);

  // STEP 2: Compute Batch Statistics
  var batchGroups = {};
  for (var r = 1; r < logData.length; r++) {
    var serial = String(logData[r][hMap.trueSerial] || "").trim();
    if (serial !== "") {
      var batchId = serial.split("-")[0].trim();
      if (!batchGroups[batchId]) batchGroups[batchId] = { c1: [], r1: [], c2: [], r2: [], rf: [] };
      batchGroups[batchId].c1.push(parseFloat(logData[r][hMap.comp1]) || 0);
      batchGroups[batchId].r1.push(Math.abs(parseFloat(logData[r][hMap.reb1])) || 0);
      batchGroups[batchId].c2.push(parseFloat(logData[r][hMap.comp2]) || 0);
      batchGroups[batchId].r2.push(Math.abs(parseFloat(logData[r][hMap.reb2])) || 0);
      batchGroups[batchId].rf.push(parseFloat(logData[r][hMap.rodForce]) || 0);
    }
  }

  var batchStats = {};
  for (var bId in batchGroups) {
    var b = batchGroups[bId];
    var mC1 = mean(b.c1), sC1 = sd(b.c1, mC1); 
    var mR1 = mean(b.r1), sR1 = sd(b.r1, mR1); 
    var mC2 = mean(b.c2), sC2 = sd(b.c2, mC2); 
    var mR2 = mean(b.r2), sR2 = sd(b.r2, mR2);
    var mRF = mean(b.rf), sRF = sd(b.rf, mRF);
    batchStats[bId] = {
      c1Mean: mC1, r1Mean: mR1, c2Mean: mC2, r2Mean: mR2,
      c1Min: mC1 - 2*sC1, c1Max: mC1 + 2*sC1, 
      r1Min: mR1 - 2*sR1, r1Max: mR1 + 2*sR1,
      c2Min: mC2 - 2*sC2, c2Max: mC2 + 2*sC2, 
      r2Min: mR2 - 2*sR2, r2Max: mR2 + 2*sR2,
      rfMin: mRF - 2*sRF, count: b.c1.length
    };
  }

  // STEP 3: Quality Diagnostics (with Prototype Router & Historical Ratchet)
  var qualityOutputSubMatrix = [];
  for (var r = 1; r < logData.length; r++) {
    var pName = String(logData[r][hMap.programName] || "").trim();
    var serial = String(logData[r][hMap.trueSerial] || "").trim();
    var baseModel = String(logData[r][hMap.baseModel] || "").trim();
    var batchId = serial.split("-")[0].trim();
    
    var test1Result = "INITIALIZING"; 
    var test2Result = "INITIALIZING"; 
    var finalStatus = "PASS";
    var failTags = [];
    var evalAction = String(logData[r][hMap.evaluationAction] || "").trim();
    var evalLower = evalAction.toLowerCase();

    var prevT1 = String(logData[r][hMap.test1Status] || "").trim().toUpperCase();
    var prevOverall = String(logData[r][hMap.overallStatus] || "").trim().toUpperCase();
    var wasCertifiedPass = prevOverall.includes("PASS") || prevT1.includes("PASS");
    var hasExplicitAction = evalLower.includes("teardown") || evalLower.includes("retest") || evalLower.includes("issue found");
    
    var isProtoRow = evalLower.includes("prototype") || evalLower.includes("not evaluated") || evalLower.includes("dev") || serial.toLowerCase().includes("proto") || serial.toLowerCase().includes("dev");

    if (pName) {
      var resolvedKey = cleanStr(modelToDynamicKey[cleanStr(baseModel)] || modelToDynamicKey[cleanStr(pName)] || pName);
      var refRow = null;
      for (var mx = 1; mx < refData.length; mx++) {
        var mKey = cleanStr(refData[mx][mMap.dynamicKey]);
        if (mKey && (mKey === resolvedKey || resolvedKey.indexOf(mKey) !== -1)) { refRow = refData[mx]; break; }
      }
      
      var lowC = parseFloat(logData[r][hMap.comp1]) || 0; 
      var lowR = Math.abs(parseFloat(logData[r][hMap.reb1])) || 0;
      var midC = parseFloat(logData[r][hMap.comp2]) || 0; 
      var midR = Math.abs(parseFloat(logData[r][hMap.reb2])) || 0;
      var sl1  = parseFloat(logData[r][hMap.slope1]) || 0; 
      var individualRF = parseFloat(logData[r][hMap.rodForce]) || 0;
      
      var legacyTightenedVariance = false;

      if (refRow) {
        var t1Pass = true;
        if (lowC < parseFloat(refRow[mMap.c1Min]) || lowC > parseFloat(refRow[mMap.c1Max])) { t1Pass = false; failTags.push("[C1_FAIL]"); }
        if (lowR < parseFloat(refRow[mMap.r1Min]) || lowR > parseFloat(refRow[mMap.r1Max])) { t1Pass = false; failTags.push("[R1_FAIL]"); }
        if (midC < parseFloat(refRow[mMap.c2Min]) || midC > parseFloat(refRow[mMap.c2Max])) { t1Pass = false; failTags.push("[C2_FAIL]"); }
        if (midR < parseFloat(refRow[mMap.r2Min]) || midR > parseFloat(refRow[mMap.r2Max])) { t1Pass = false; failTags.push("[R2_FAIL]"); }
        if (sl1 < (parseFloat(refRow[mMap.slope1Min]) || 0)) { t1Pass = false; failTags.push("[SLOPE_FAIL]"); }

        // HISTORICAL INTEGRITY RATCHET
        if (!t1Pass && wasCertifiedPass && !hasExplicitAction && !isProtoRow) {
          t1Pass = true;
          test1Result = prevT1.includes("OVERRIDE") ? "PASS (OVERRIDE)" : "PASS";
          failTags = []; 
          legacyTightenedVariance = true;
        } else if (!isProtoRow) {
          test1Result = t1Pass ? "PASS" : "FAIL (BLUEPRINT)";
        }
      }
      
      var cStat = batchStats[batchId];
      var defectAnalysis = "";
      if (cStat && cStat.count > 2 && !isProtoRow) {
        var t2Pass = true;
        var lowGasPressure = (individualRF < cStat.rfMin);
        var lowC1 = (lowC < cStat.c1Min || lowC > cStat.c1Max);
        var lowR1 = (lowR < cStat.r1Min || lowR > cStat.r1Max);
        
        if (lowGasPressure) { t2Pass = false; failTags.push("[RF_FAIL]"); }
        if (lowC1) { t2Pass = false; failTags.push("[C1_FAIL]"); }
        if (lowR1) { t2Pass = false; failTags.push("[R1_FAIL]"); }
        
        if (!t2Pass) {
          if (lowGasPressure) defectAnalysis = "Gas Pressure Deficient.";
          else if (lowC1 && lowR1) defectAnalysis = "Symmetric drop. Potential bypass.";
          else defectAnalysis = "Outlier variance detected.";
        }
        test2Result = t2Pass ? "PASS" : "FAIL (OUTLIER)";
      }
      
      var batchDriftTag = "";
      if (refRow && cStat && cStat.count >= 3 && !isProtoRow) {
        var masterR1Mean = parseFloat(refRow[mMap.r1Mean]);
        if (!isNaN(masterR1Mean) && masterR1Mean > 0) {
          var bRebDrift = ((cStat.r1Mean - masterR1Mean) / masterR1Mean) * 100;
          if (Math.abs(bRebDrift) >= 5) batchDriftTag = "[BATCH_DRIFT_REB " + (bRebDrift > 0 ? "+" : "") + bRebDrift.toFixed(1) + "%]";
        }
      }

      var uniqueFailTags = [];
      for (var f = 0; f < failTags.length; f++) {
        if (uniqueFailTags.indexOf(failTags[f]) === -1) uniqueFailTags.push(failTags[f]);
      }
      
      var globalPass = (test1Result === "INITIALIZING" || !test1Result.includes("FAIL")) && (test2Result === "INITIALIZING" || !test2Result.includes("FAIL"));
      var diagnosticNotes = "";

      if (globalPass) {
        if (legacyTightenedVariance) diagnosticNotes = "✅ PASS (Legacy Baseline) | [LEGACY_PASS: Outside Refined Spec]";
        else if (batchDriftTag) diagnosticNotes = "✅ SHOCK IS WITHIN TOLERANCE. ⚠️ Notice: " + batchDriftTag;
        else diagnosticNotes = "✅ SHOCK IS WITHIN TOLERANCE.";
      } else {
        diagnosticNotes = "❌ ERROR: " + uniqueFailTags.join(" ") + " | " + defectAnalysis;
      }

      if (isProtoRow) {
        test1Result = "NOT EVALUATED"; 
        test2Result = "NOT EVALUATED"; 
        finalStatus = "PROTOTYPE";
        diagnosticNotes = "🔬 PROTOTYPE / R&D: Logged for development; excluded from production SPC.";
      } else if (evalLower.indexOf("override test 1") !== -1) {
        test1Result = "PASS (OVERRIDE)"; 
        finalStatus = (test2Result.indexOf("PASS") !== -1) ? "PASS (OVERRIDE)" : "FAIL";
        diagnosticNotes = (finalStatus === "FAIL") ? "⚠️ TEST 1 OVERRIDDEN | Test 2 Failure: " + uniqueFailTags.join(" ") : "⚠️ MANUAL OVERRIDE";
      } else if (evalLower.indexOf("override test 2") !== -1) {
        test2Result = "PASS (OVERRIDE)"; 
        finalStatus = (test1Result.indexOf("PASS") !== -1) ? "PASS (OVERRIDE)" : "FAIL";
        diagnosticNotes = (finalStatus === "FAIL") ? "⚠️ TEST 2 OVERRIDDEN | Test 1 Failure: " + uniqueFailTags.join(" ") : "⚠️ MANUAL OVERRIDE";
      } else if (evalLower.indexOf("override all") !== -1) {
        test1Result = "PASS (OVERRIDE)"; 
        test2Result = "PASS (OVERRIDE)"; 
        finalStatus = "PASS (OVERRIDE)"; 
        diagnosticNotes = "⚠️ MANUAL OVERRIDE";
      } else if (evalLower.indexOf("retest") !== -1) {
        finalStatus = "HOLD (RETEST REQUIRED)"; 
        diagnosticNotes = "⏳ RETEST REQUIRED.";
      } else if (evalLower.indexOf("teardown") !== -1) {
        finalStatus = "HOLD (TEARDOWN REQUIRED)"; 
        diagnosticNotes = "🔧 TEARDOWN REQUIRED.";
      } else if (evalLower.indexOf("no issue") !== -1) {
        test1Result = "PASS"; 
        test2Result = "PASS"; 
        finalStatus = "PASS"; 
        diagnosticNotes = "🛠️ TEARDOWN VALIDATED.";
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