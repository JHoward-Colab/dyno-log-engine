// =========================================================================
// ⚙️ GLOBAL CONFIGURATION & CONSTANTS (Config.js)
// Single Source of Truth for Sheet Names, Ranges, and Column Mappings
// =========================================================================

var CONFIG = {
  SHEET_NAMES: {
    OPERATOR_STATION: "Operator Station",
    MASTER_DYNO_LOG: "Master_Dyno_Log",
    PART_REFERENCE_MATRIX: "Part_Reference_Matrix",
    PROGRAM_REGISTRY: "Program_Registry"
  },

  FOLDERS: {
    WATCH_FOLDER: "01_Watch_Folder_DEV",
    ARCHIVE_FOLDER: "02_Archive_DEV"
  },

  OPERATOR_STATION: {
    RANGES: {
      BARCODE_INPUT: "C4",
      CACHED_FILE_ID: "C5",
      FILE_LINK_OUTPUT: "C5",
      BOM_REV_OUTPUT: "E4",
      PART_NO_OUTPUT: "E5",

      CLEAR_METADATA_RANGE: "C4:E5",
      CLEAR_RESULTS_RANGE: "A27:L100",

      LIMIT_C1_MIN: "B22",
      LIMIT_C1_MAX: "B23",
      LIMIT_R1_MIN: "C22",
      LIMIT_R1_MAX: "C23",
      LIMIT_C2_MIN: "D22",
      LIMIT_C2_MAX: "D23",
      LIMIT_R2_MIN: "E22",
      LIMIT_R2_MAX: "E23",
      LIMIT_SLOPE: "F22",

      RESULTS_START_ROW: 27,
      RESULTS_START_COL: 1,
      RESULTS_COL_COUNT: 12
    }
  },

  COLUMNS: {
    MASTER_DYNO_LOG: {
      TIMESTAMP: 1,         // Col A
      PROGRAM_NAME: 2,      // Col B
      TRUE_SERIAL: 3,       // Col C
      BASE_MODEL: 4,        // Col D
      VALVING_VERSION: 5,   // Col E
      ROD_FORCE: 8,         // Col H
      COMP_1: 9,            // Col I
      REB_1: 10,            // Col J
      COMP_2: 11,           // Col K
      REB_2: 12,            // Col L
      COMP_3: 13,           // Col M
      REB_3: 14,            // Col N
      TEST_1_STATUS: 22,    // Col V
      TEST_2_STATUS: 23,    // Col W
      OVERALL_STATUS: 24,   // Col X
      TEARDOWN_ACTION: 25,  // Col Y
      EVALUATION_ACTION: 26,// Col Z
      DIAGNOSTICS: 27,      // Col AA
      ENG_COMMENTS: 28      // Col AB
    },

    PART_REFERENCE_MATRIX: {
      PROGRAM_NAME: 1,
      BASE_MODEL: 3,
      COMP_1_MIN: 2,
      COMP_1_MAX: 3,
      REB_1_MIN: 4,
      REB_1_MAX: 5,
      COMP_2_MIN: 6,
      COMP_2_MAX: 7,
      REB_2_MIN: 8,
      REB_2_MAX: 9,
      SLOPE_1_MIN: 10
    },

    PROGRAM_REGISTRY: {
      PROGRAM_NAME: 1,
      BASE_MODEL: 2
    }
  },

  NOMINAL_SPEEDS: [100, 400, 1000, 2500],

  STATUS: {
    PASS: "PASS",
    FAIL: "FAIL",
    FAIL_BLUEPRINT: "FAIL (BLUEPRINT)",
    FAIL_OUTLIER: "FAIL (OUTLIER)",
    OVERRIDE_PASS: "PASS (OVERRIDE)",
    INITIALIZING: "INITIALIZING",
    NOT_RUN: "NOT RUN",
    NOT_TESTED: "NOT TESTED YET"
  }
};