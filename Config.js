// =========================================================================
// ⚙️ GLOBAL CONFIGURATION & CONSTANTS (Config.js)
// Single Source of Truth for Sheet Names, Ranges, and Column Mappings
// =========================================================================

const CONFIG = Object.freeze({
  SHEET_NAMES: Object.freeze({
    OPERATOR_STATION: "Operator Station",
    MASTER_DYNO_LOG: "Master_Dyno_Log",
    PART_REFERENCE_MATRIX: "Part_Reference_Matrix",
    PROGRAM_REGISTRY: "Program_Registry"
  }),

  // 🔒 DEV-SAFE FOLDER TARGETS
  FOLDERS: Object.freeze({
    WATCH_FOLDER: "01_Watch_Folder_DEV",
    ARCHIVE_FOLDER: "02_Archive_DEV"
  }),

  OPERATOR_STATION: Object.freeze({
    RANGES: Object.freeze({
      BARCODE_INPUT: "C4",
      CACHED_FILE_ID: "Z1",
      FILE_LINK_OUTPUT: "C5",
      BOM_REV_OUTPUT: "E4",
      PART_NO_OUTPUT: "E5",

      CLEAR_METADATA_RANGE: "C4:E5",
      CLEAR_RESULTS_RANGE: "A27:L100",

      LIMIT_C1_MIN: "B22",
      LIMIT_C1_MAX: "C22",
      LIMIT_R1_MIN: "D22",
      LIMIT_R1_MAX: "E22",
      LIMIT_C2_MIN: "B23",
      LIMIT_C2_MAX: "C23",
      LIMIT_R2_MIN: "D23",
      LIMIT_R2_MAX: "E23",
      LIMIT_SLOPE: "F22",

      RESULTS_START_ROW: 27,
      RESULTS_START_COL: 1,
      RESULTS_COL_COUNT: 12
    })
  }),

  COLUMNS: Object.freeze({
    MASTER_DYNO_LOG: Object.freeze({
      TIMESTAMP: 1,         // Col A
      PROGRAM_NAME: 2,      // Col B
      TRUE_SERIAL: 3,       // Col C
      BASE_MODEL: 4,        // Col D
      VALVING_VERSION: 5,   // Col E
      ROD_FORCE: 6,         // Col F
      SPEED_1: 7,           // Col G
      COMP_1: 8,            // Col H
      REB_1: 9,             // Col I
      SLOPE_1: 10,          // Col J
      LOOP_AREA_1: 11,      // Col K
      SPEED_2: 12,          // Col L
      COMP_2: 13,           // Col M
      REB_2: 14,            // Col N
      SPEED_3: 17,          // Col Q
      COMP_3: 18,           // Col R
      REB_3: 19,            // Col S
      TEST_1_STATUS: 22,    // Col V
      TEST_2_STATUS: 23,    // Col W
      OVERALL_STATUS: 24,   // Col X
      DIAGNOSTICS: 25,      // Col Y (Diagnostics & Troubleshooting)
      EVALUATION_ACTION: 26,// Col Z (Evaluation Action)
      ENG_COMMENTS: 27      // Col AA (Diagnostic Notes)
    }),

    PART_REFERENCE_MATRIX: Object.freeze({
      PROGRAM_NAME: 1,
      COMP_1_MIN: 7,
      COMP_1_MAX: 8,
      REB_1_MIN: 11,
      REB_1_MAX: 12,
      COMP_2_MIN: 17,
      COMP_2_MAX: 18,
      REB_2_MIN: 21,
      REB_2_MAX: 22,
      SLOPE_1_MIN: 13
    }),

    PROGRAM_REGISTRY: Object.freeze({
      PROGRAM_NAME: 1,
      BASE_MODEL: 3
    })
  }),

  NOMINAL_SPEEDS: Object.freeze([100, 400, 1000, 2500]),

  STATUS: Object.freeze({
    PASS: "PASS", FAIL: "FAIL", FAIL_BLUEPRINT: "FAIL (BLUEPRINT)",
    FAIL_OUTLIER: "FAIL (OUTLIER)", OVERRIDE_PASS: "PASS (OVERRIDE)",
    INITIALIZING: "INITIALIZING", NOT_RUN: "NOT RUN", NOT_TESTED: "NOT TESTED YET"
  }),

  DIAGNOSTIC_TAGS: Object.freeze({
    ROD_FORCE_FAIL: "[RF_FAIL]", COMP_1_FAIL: "[C1_FAIL]", REB_1_FAIL: "[R1_FAIL]",
    COMP_2_FAIL: "[C2_FAIL]", REB_2_FAIL: "[R2_FAIL]", SLOPE_FAIL: "[SLOPE_FAIL]"
  })
});