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