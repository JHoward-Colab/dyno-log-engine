// =========================================================================
// ⚙️ CONFIGURATION (Config.gs)
// =========================================================================

const CONFIG = Object.freeze({
  SHEET_NAMES: Object.freeze({
    MASTER_DYNO_LOG: "Master_Dyno_Log",
    PART_REFERENCE_MATRIX: "Part_Reference_Matrix",
    OPERATOR_STATION: "Operator_Station",
    PROGRAM_REGISTRY: "Program_Registry"
  }),

  FOLDERS: Object.freeze({
    WATCH_FOLDER: "01_Watch_Folder_DEV",
    ARCHIVE_FOLDER: "02_Archive_DEV"
  }),

  COLUMNS: Object.freeze({
    MASTER_DYNO_LOG: Object.freeze({
      TIMESTAMP: 1,          // Col A
      PROGRAM_NAME: 2,       // Col B
      TRUE_SERIAL: 3,        // Col C
      BASE_MODEL: 4,         // Col D
      VALVING_VERSION: 5,    // Col E
      ROD_FORCE: 6,          // Col F
      SPEED_1: 7,            // Col G
      COMP_1: 8,             // Col H
      REB_1: 9,              // Col I
      SLOPE_1: 10,           // Col J
      LOOP_AREA_1: 11,       // Col K
      SPEED_2: 12,           // Col L
      COMP_2: 13,            // Col M
      REB_2: 14,             // Col N
      LOOP_AREA_2: 15,       // Col O
      SPEED_3: 16,           // Col P
      COMP_3: 17,            // Col Q
      REB_3: 18,             // Col R
      TEST_1_STATUS: 19,     // Col S
      TEST_2_STATUS: 20,     // Col T
      OVERALL_STATUS: 21,    // Col U
      DIAGNOSTICS: 22,       // Col V
      EVALUATION_ACTION: 23, // Col W
      ENG_COMMENTS: 24       // Col X
    }),

    PROGRAM_REGISTRY: Object.freeze({
      PROGRAM_NAME: 1,         // Col A
      DYNAMIC_KEY: 2,          // Col B
      BASE_MODEL: 3,           // Col C
      CUSTOMER_ACCOUNT: 4,     // Col D
      VEHICLE_SPEC: 5,         // Col E
      PRODUCT_TYPE: 6,         // Col F
      VALVING_VERSION: 7,      // Col G
      ADJUSTER_SETTINGS: 8,    // Col H
      BOM_REV: 9,              // Col I
      APPROVED_LEGACY_REVS: 10,// Col J
      ECN_NOTES: 11            // Col K
    }),

    PART_REFERENCE_MATRIX: Object.freeze({
      DYNAMIC_KEY: 1,        // Col A
      SPEED_1: 2,            // Col B
      SPEED_2: 3,            // Col C
      SPEED_3: 4,            // Col D
      COMP_1_MEAN: 5,        // Col E
      COMP_1_SD: 6,          // Col F
      COMP_1_MIN: 7,         // Col G
      COMP_1_MAX: 8,         // Col H
      REB_1_MEAN: 9,         // Col I
      REB_1_SD: 10,          // Col J
      REB_1_MIN: 11,         // Col K
      REB_1_MAX: 12,         // Col L
      SLOPE_1_MIN: 13,       // Col M
      LOOP_AREA_1_MIN: 14,   // Col N
      COMP_2_MEAN: 15,       // Col O
      COMP_2_SD: 16,         // Col P
      COMP_2_MIN: 17,        // Col Q
      COMP_2_MAX: 18,        // Col R
      REB_2_MEAN: 19,        // Col S
      REB_2_SD: 20,          // Col T
      REB_2_MIN: 21,         // Col U
      REB_2_MAX: 22,         // Col V
      SLOPE_2_MIN: 23,       // Col W
      COMP_3_MEAN: 24,       // Col X
      COMP_3_SD: 25,         // Col Y
      COMP_3_MIN: 26,        // Col Z
      COMP_3_MAX: 27,        // Col AA
      REB_3_MEAN: 28,        // Col AB
      REB_3_SD: 29,          // Col AC
      REB_3_MIN: 30,         // Col AD
      REB_3_MAX: 31,         // Col AE
      SAMPLE_COUNT: 32,      // Col AF
      // Preserved Engineering SPC Columns (Cols AG to AN)
      RAW_COMP_SD: 33,
      RAW_REB_SD: 34,
      ORIGINAL_COMP_BASE: 35,
      ORIGINAL_REB_BASE: 36,
      COMP_RANGE_WIDTH: 37,
      REB_RANGE_WIDTH: 38,
      COMP_DRIFT_PCT: 39,
      REB_DRIFT_PCT: 40,
      HEALTH_STAMP: 41,      // Col AO
      CONTROL_MODE: 42       // Col AP
    })
  }),

  OPERATOR_STATION: Object.freeze({
    RANGES: Object.freeze({
      BARCODE_INPUT: "C2",
      FILE_LINK_OUTPUT: "C3",
      BOM_REV_OUTPUT: "C4",
      BASE_MODEL_OUTPUT: "C5",
      CROSS_CHECK_OUTPUT: "C6",
      CUSTOMER_ACCOUNT_OUTPUT: "C14",
      VEHICLE_SPEC_OUTPUT: "C15",
      PROGRAM_NAME_OUTPUT: "C16",
      VALVING_VERSION_OUTPUT: "C17",
      ADJUSTER_TARGETS_OUTPUT: "C18",
      CACHED_FILE_ID: "Z1",
      LIMIT_COMP_1_MIN: "B22",
      LIMIT_COMP_1_MAX: "C22",
      LIMIT_REB_1_MIN: "D22",
      LIMIT_REB_1_MAX: "E22",
      LIMIT_COMP_2_MIN: "B23",
      LIMIT_COMP_2_MAX: "C23",
      LIMIT_REB_2_MIN: "D23",
      LIMIT_REB_2_MAX: "E23",
      LIMIT_SLOPE_1_MIN: "F22",
      STATUS_BANNER: "A8",
      CLEAR_METADATA_RANGE: "C3:C6",
      CLEAR_PANEL_RANGE: "C14:C18",
      CLEAR_LIMITS_RANGE: "B22:F23",
      CLEAR_RESULTS_RANGE: "A27:L100",
      RESULTS_START_ROW: 27,
      RESULTS_START_COL: 1,
      RESULTS_COL_COUNT: 12
    })
  }),

  STATUS: Object.freeze({
    PASS: "PASS",
    FAIL: "FAIL",
    FAIL_BLUEPRINT: "FAIL (BLUEPRINT)",
    FAIL_OUTLIER: "FAIL (OUTLIER)",
    OVERRIDE_PASS: "PASS (OVERRIDE)",
    INITIALIZING: "INITIALIZING",
    NOT_RUN: "NOT RUN",
    NOT_TESTED: "NOT TESTED YET"
  }),

  DIAGNOSTIC_TAGS: Object.freeze({
    ROD_FORCE_FAIL: "[RF_FAIL]",
    COMP_1_FAIL: "[C1_FAIL]",
    REB_1_FAIL: "[R1_FAIL]",
    COMP_2_FAIL: "[C2_FAIL]",
    REB_2_FAIL: "[R2_FAIL]",
    SLOPE_FAIL: "[SLOPE_FAIL]"
  })
});