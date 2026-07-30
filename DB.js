// =========================================================================
// 🗄️ DATABASE ACCESS & UTILITIES (DB.gs)
// =========================================================================

/**
 * Safely fetches two-dimensional array of values from a named sheet tab.
 */
function getSheetValues(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  return sheet ? sheet.getDataRange().getValues() : [];
}

/**
 * Normalizes keys by lowercasing and removing non-alphanumeric characters.
 */
function cleanStringKey(val) {
  if (val === null || val === undefined) return "";
  return String(val).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Creates a key-to-index dictionary from a row of headers.
 */
function getHeaderMap(headerRow) {
  const map = {};
  if (!headerRow) return map;
  for (let i = 0; i < headerRow.length; i++) {
    const key = cleanStringKey(headerRow[i]);
    if (key) map[key] = i;
  }
  return map;
}

/**
 * Finds index of first header matching provided key aliases.
 */
function findColIndex(hMap, keys) {
  if (!hMap) return -1;
  for (let i = 0; i < keys.length; i++) {
    const k = cleanStringKey(keys[i]);
    if (hMap.hasOwnProperty(k)) return hMap[k];
  }
  return -1;
}

/**
 * Safe numeric extraction helper.
 */
function findColVal(row, hMap, keys) {
  const colIdx = findColIndex(hMap, keys);
  if (colIdx >= 0 && row[colIdx] !== "" && row[colIdx] !== undefined) {
    const val = Number(row[colIdx]);
    return isNaN(val) ? row[colIdx] : Number(val.toFixed(1));
  }
  return undefined;
}