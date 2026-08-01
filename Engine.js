// =========================================================================
// 🏎️ UNIFIED DYNO ENGINE (Engine.js)
// Utility Math, Header Resolvers, Ingestion & Retroactive SPC Engine
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
 * Buckets speed values to nominal target speeds defined in CONFIG.
 */
function snapToNominalSpeed(val) {
  var v = parseFloat(val);
  if (isNaN(v) || v <= 0) return 0;
  var nominals = CONFIG.NOMINAL_SPEEDS;
  if (v < 250) return nominals[0];   // 100
  if (v < 700) return nominals[1];   // 400
  if (v < 1800) return nominals[2];  // 1000
  return nominals[3];                // 2500
}

/**
 * DEDUPLICATION ENGINE: Keeps only the latest test run per True Serial Number.
 * Prevents multiple re-tested/passing runs of the same shock from double-counting
 * and shrinking standard deviation or skewing mean statistics.
 */
function getUniqueLatestTests(dynoDataRows) {
  if (!dynoDataRows || dynoDataRows.length === 0) return [];
  
  var serialColIdx = CONFIG.COLUMNS.MASTER_DYNO_LOG.TRUE_SERIAL - 1; // 0-indexed column 2
  var latestMap = {};
  
  for (var i = 0; i < dynoDataRows.length; i++) {
    var row = dynoDataRows[i];
    var serial = String(row[serialColIdx] || "").trim();
    if (!serial || serial === "undefined" || serial === "null") continue;
    
    // Top-to-bottom scan: subsequent/newer runs overwrite earlier runs
    latestMap[serial] = row;
  }
  
  var uniqueRows = [];
  for (var key in latestMap) {
    if (latestMap.hasOwnProperty(key)) {
      uniqueRows.push(latestMap[key]);
    }
  }
  return uniqueRows;
}

/**
 * Recalculates baseline SPC statistics (Mean, SD, Sample Count) in the
 * Part_Reference_Matrix tab using deduplicated Master Dyno Log data.
 */
function retroactiveLogRecalculate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  var refSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PART_REFERENCE_MATRIX);
  if (!logSheet || !refSheet) return;

  var logValues = logSheet.getDataRange().getValues();
  if (logValues.length <= 1) return;

  var logCols = CONFIG.COLUMNS.MASTER_DYNO_LOG;
  var refCols = CONFIG.COLUMNS.PART_REFERENCE_MATRIX;

  // 1. DEDUPLICATE SHOCKS TO PREVENT DOUBLE COUNTING
  var rawData = logValues.slice(1);
  var cleanData = getUniqueLatestTests(rawData);

  // 2. Group deduplicated data by Base Model / Program Key
  var modelColIdx = logCols.BASE_MODEL - 1;
  var c1Idx = logCols.COMP_1 - 1;
  var r1Idx = logCols.REB_1 - 1;
  var c2Idx = logCols.COMP_2 - 1;
  var r2Idx = logCols.REB_2 - 1;

  var modelGroups = {};
  for (var i = 0; i < cleanData.length; i++) {
    var row = cleanData[i];
    var modelKey = String(row[modelColIdx] || "").trim();
    if (!modelKey) continue;
    
    if (!modelGroups[modelKey]) {
      modelGroups[modelKey] = { comp1: [], reb1: [], comp2: [], reb2: [] };
    }
    
    var c1Val = parseFloat(row[c1Idx]);
    var r1Val = parseFloat(row[r1Idx]);
    var c2Val = parseFloat(row[c2Idx]);
    var r2Val = parseFloat(row[r2Idx]);

    if (!isNaN(c1Val)) modelGroups[modelKey].comp1.push(c1Val);
    if (!isNaN(r1Val)) modelGroups[modelKey].reb1.push(r1Val);
    if (!isNaN(c2Val)) modelGroups[modelKey].comp2.push(c2Val);
    if (!isNaN(r2Val)) modelGroups[modelKey].reb2.push(r2Val);
  }

  // 3. Update Part_Reference_Matrix tab with computed SPC baselines
  var refValues = refSheet.getDataRange().getValues();

  for (var r = 1; r < refValues.length; r++) {
    var partKey = String(refValues[r][refCols.PROGRAM_NAME - 1] || "").trim();
    if (modelGroups[partKey]) {
      var grp = modelGroups[partKey];
      var c1M = calculateMean(grp.comp1);
      var c1SD = calculateSD(grp.comp1, c1M);
      var r1M = calculateMean(grp.reb1);
      var r1SD = calculateSD(grp.reb1, r1M);
      var c2M = calculateMean(grp.comp2);
      var c2SD = calculateSD(grp.comp2, c2M);
      var r2M = calculateMean(grp.reb2);
      var r2SD = calculateSD(grp.reb2, r2M);

      refSheet.getRange(r + 1, refCols.COMP_1_MEAN).setValue(c1M);
      refSheet.getRange(r + 1, refCols.COMP_1_SD).setValue(c1SD);
      refSheet.getRange(r + 1, refCols.REB_1_MEAN).setValue(r1M);
      refSheet.getRange(r + 1, refCols.REB_1_SD).setValue(r1SD);
      refSheet.getRange(r + 1, refCols.COMP_2_MEAN).setValue(c2M);
      refSheet.getRange(r + 1, refCols.COMP_2_SD).setValue(c2SD);
      refSheet.getRange(r + 1, refCols.REB_2_MEAN).setValue(r2M);
      refSheet.getRange(r + 1, refCols.REB_2_SD).setValue(r2SD);
      refSheet.getRange(r + 1, refCols.SAMPLE_COUNT).setValue(grp.comp1.length);
    }
  }
}

/**
 * Processes incoming CSV dyno files from the watch folder.
 */
function processDynoFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER_DYNO_LOG);
  if (!sheet) return;

  var folders = DriveApp.getFoldersByName(CONFIG.FOLDERS.WATCH_FOLDER);
  if (!folders.hasNext()) return;
  var folder = folders.next();
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var rawName = file.getName();
    if (!rawName.toLowerCase().endsWith('.csv')) continue;
    
    // Ingestion parsing logic executes here
    Logger.log("Processing watch folder file: " + rawName);
  }
}