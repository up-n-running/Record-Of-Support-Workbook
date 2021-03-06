SHEETS["1.5"] = {
  INPUT: {
      NAME: 'Input',
      REFS: {
          LSA_NAME: { ROW_NO: 1, COL_NO: 5 },
          LESSON_DATE: { ROW_NO: 1, COL_NO: 2 },
          COL_NO_RECORD_1: 2,
          COL_NO_RECORD_LAST: 26,
          ROW_NO_STATUS_BAR: 3,
          ROW_NO_RECORD_NOS: 4,
          ROW_NO_PERIOD: 5,
          ROW_NO_LESSON_NAME: 6,
          ROW_NO_LEARNER_NAME: 7,
          ROW_NO_ATTENDED: 8,
          ROW_NO_LATE: 9,
          ROW_NO_ONLINE: 10,
          ROW_NO_LESSON_TARGETS: 12,
          ROW_NO_SUPPORT_STRAT_FIRST: 14,
          ROW_NO_SUPPORT_STRAT_LAST: 20,
          ROW_NO_SUPPORT_STRAT_EXTRA: 21,
          ROW_NO_RESOURCES_USED: 22,
          ROW_NO_OUT_REQUESTS: 24,
          ROW_NO_OUT_TASKS: 25,
          ROW_NO_OUT_TARGETS: 26,
          ROW_NO_OUT_TEXT: 27,
          ROW_NO_LSA_COMMENTS: 29,
          ROW_NO_AUTOSIGN_COMMENTS: 34,
          ROW_NO_EMAIL_ADDRESS: 35,
          ROW_NO_LEARNER_ID: 36,
          ROW_NO_LEARNER_EMAILED: 38,
          ROW_NO_AUTOSIGN_MANUALENTRY: 39,
          ROW_NO_FILE_ID: 40,
          ROW_NO_FILE_NAME: 41,
          ROW_NO_FILE_CREATED: 42,
          ROW_NO_FILE_UPDATED: 43,
          REF_DATE: "B1",
          REF_LSA_NAME: "E1"
      },
      STATUSES: {
          SIGNED: "Signed",
          SIGNED_AUTOSIGN: "Auto-Signed",
          UNSIGNED: "Sent to Learner",
          UNSENT: "To Send to Learner",
          SAVED_EMAILWAIT: "Saved, Awaiting Email",
          UNSENT_AUTOSIGN: "To Save (Auto-Sign)",
          BLANK: ""
      },
      ATTENDED_STATUSES: {
          ATTENDED: "Yes",
          UNAUTH_ABSENCE: "No: Unauthorised Absence",
          AUTH_ABSENCE: "No: Authorised Absence (No Email)"
      },
      MOBILE_COPY: [
          [["PERIOD", ""], ["LESSON_NAME", ""], ["LEARNER_NAME", ""], ["ATTENDED", "Yes"], ["LATE", "No"], ["ONLINE", "No"]],
          [["LESSON_TARGETS", ""]],
          [["SUPPORT_STRAT_FIRST", null], ["SUPPORT_STRAT_EXTRA", ""], ["RESOURCES_USED", ""]],
          [["OUT_REQUESTS", "Yes"], ["OUT_TASKS", "Yes"], ["OUT_TARGETS", "Most"], ["OUT_TEXT", ""]],
          [["LSA_COMMENTS", ""]]
      ],
      DEFAULTS: [
          [["PERIOD", ""], ["LESSON_NAME", ""], ["LEARNER_NAME", ""], ["ATTENDED", "Yes"], ["LATE", "No"], ["ONLINE", "No"]],
          [["LESSON_TARGETS", ""]],
          [["SUPPORT_STRAT_FIRST", null], ["SUPPORT_STRAT_EXTRA", ""], ["RESOURCES_USED", ""]],
          [["OUT_REQUESTS", "Yes"], ["OUT_TASKS", "Yes"], ["OUT_TARGETS", "Most"], ["OUT_TEXT", ""]],
          [["LSA_COMMENTS", ""]],
          [["LEARNER_EMAILED", ""], ["AUTOSIGN_COMMENTS", ""], ["FILE_ID", ""], ["FILE_NAME", ""], ["FILE_CREATED", ""], ["FILE_UPDATED", ""]]
      ],
      DEFAULTS_ATTENDANCE: {
          NO: [
              [["LATE", "N/A"]],
              [["SUPPORT_STRAT_FIRST", null], ["SUPPORT_STRAT_EXTRA", ""], ["ROW_NO_RESOURCES_USED", ""]],
              [["OUT_REQUESTS", "N/A"], ["OUT_TASKS", "N/A"], ["OUT_TARGETS", "N/A"]],
              [["AUTOSIGN_MANUALENTRY", ""]]
          ],
          YES: [
              [["LATE", "No"]],
              [["OUT_REQUESTS", "Yes"], ["OUT_TASKS", "Yes"], ["OUT_TARGETS", "Most"]],
              [["AUTOSIGN_MANUALENTRY", ""]]
          ]
      },
      COPY_TO_ROS_SHEETS: {
          START_COPY_ROW: 6,
          START_COPY_COL: 2,
          FIELDS: ["PERIOD", "LESSON_NAME", "LEARNER_NAME", "ATTENDED", "LATE", "ONLINE", "LESSON_TARGETS",
              "SUPPORT_STRAT_EXTRA", "RESOURCES_USED", "OUT_REQUESTS", "OUT_TASKS", "OUT_TARGETS", "OUT_TEXT", "LSA_COMMENTS",
              "AUTOSIGN_COMMENTS", "LEARNER_ID"],
          RANGES: ["SUPPORT_STRAT"]
      }
  },
  MY_FILES: {
      NAME: 'My Files',
      REFS: {
          ROW_NO_FIRST_FILE: 3,
          ROW_NO_LAST_FILE: 252,
          COL_NO_STATUS_BAR: 1,
          COL_NO_FILE_NAME: 8,
          COL_NO_USERS_FILE_ID: 9,
          COL_NO_CREATED_DATE: 10,
          COL_NO_UPDATED_DATE: 11,
          COL_NO_DELETED_DATE: 12,
          COL_NO_LEARNER_NAME: 13,
          COL_NO_LEARNER_ID: 14,
          COL_NO_LEARNER_EMAIL: 15,
          COL_NO_LESSON_NAME: 16,
          COL_NO_LESSON_DATE: 17,
          COL_NO_PERIOD: 18,
          COL_NO_AUTOSIGN_CMTS: 19
      },
      STATUSES: {
          SIGNED: "Signed",
          AUTOSIGNED: "Auto-Signed",
          UNSIGNED: "Unsigned",
          TRASHED: "In Trash Bin",
          DELETED: "Deleted"
      }
  },
  SETTINGS_LEARNERS: {
      NAME: 'Settings - Learners',
      HANDLE: "SETTINGS_LEARNERS",
      REFS: {
          COL_NO_LEARNER_NAME: 1,
          COL_NO_EDITABLE_NICKNAME: 4,
          COL_NO_SUPPORT_NEED: 5,
          COL_NO_SUPPORT_STRAT_FIRST: 6,
          COL_NO_SUPPORT_STRAT_LAST: 10,
          COL_NO_EXTRA_SUPPORT_TEXT: 11,
          HANDLE_EDITABLE_NICKNAME: "EDITABLE_NICKNAME",
          COL_NO_FORENAME: 12,
          HANDLE_FORENAME: "FORENAME",
          COL_NO_NICKNAME: 13,
          HANDLE_NICKNAME: "NICKNAME",
          COL_NO_SURNAME: 14,
          HANDLE_SURNAME: "SURNAME",
          COL_NO_LEARNER_ID: 15,
          HANDLE_LEARNER_ID: "LEARNER_ID",
          COL_NO_CATEGORY: 16,
          HANDLE_CATEGORY: "CATEGORY",
          COL_NO_EMAIL_ADDRESS: 17,
          HANDLE_EMAIL_ADDRESS: "EMAIL_ADDRESS",
          COL_NO_EXTERNAL_ID_1: 18,
          HANDLE_EXTERNAL_ID_1: "EXTERNAL_ID_1",
          COL_NO_LEARNER_DIR: 19,
          HANDLE_LEARNER_DIR: "LEARNER_DIR",
          ROW_NO_FIRST_LEARNER: 3,
          ROW_NO_LAST_LEARNER: 52
      }
  },
  SETTINGS_LESSONS: {
      NAME: 'Settings - Lessons',
      REFS: {
          COL_NO_LESSON_NAME_READONLY: 1,
          COL_NO_EQUIPMENT_USED: 2,
          COL_NO_SUPPORT_STRAT_FIRST: 3,
          COL_NO_SUPPORT_STRAT_LAST: 4,
          COL_NO_LESSON_NAME: 5,
          ROW_NO_FIRST_LESSON: 3,
          ROW_NO_LAST_LESSON: 22
      }
  },
  SETTINGS_TARGET_GRADES: {
      NAME: 'Settings - Target Grades',
      REFS: {
          COL_NO_LEARNER_NAMES: 1,
          ROW_NO_LESSON_NAMES: 1
      }
  },
  SETTINGS_LESSON_TARGETS: {
      NAME: 'Settings - Ongoing Targets',
      REFS: {
          COL_NO_LEARNER_NAMES: 1,
          ROW_NO_LESSON_NAMES: 1
      }
  },
  GLOBAL_SETTINGS: {
      NAME: 'Global Settings',
      REFS: {
          COL_NO: 2,
          ROW_NO_VERSION_NO: 1,
          HANDLE_VERSION_NO: "[VERSION_NO]",
          ROW_NO_THIS_FILES_ID: 2,
          HANDLE_THIS_FILES_ID: "[THIS_FILES_ID]",
          ROW_NO_DIR_ID_ROS: 5,
          ROW_NO_DIR_ID_LSAS: 6,
          ROW_NO_WEBAPP_ALL_DEPLOYID: 7,
          HANDLE_WEBAPP_ALL_DEPLOYID: "[WEBAPP_ALL_DEPLOYID]",
          ROW_NO_VALID_EMAIL_REGEX: 8,
          ROW_NO_MAIN_USERS_EMAIL: 9,
          HANDLE_MAIN_USERS_EMAIL: "[MAIN_USERS_EMAIL]",
          ROW_NO_MASTER_FILE_ID: 10,
          HANDLE_MASTER_FILE_ID: "[MASTER_FILE_ID]",
          ROW_NO_MASTERS_LINK_TO_THIS_CHILD: 11,
          HANDLE_MASTERS_LINK_TO_THIS_CHILD: "[MASTERS_LINK_TO_THIS_CHILD]",
          ROW_NO_MASTERS_VERSION: 12,
          HANDLE_MASTERS_VERSION: "[MASTERS_VERSION]",
          ROW_NO_DEVS_EMAIL: 13,
          ROW_NO_LSA_GROUP_EMAIL: 14,
          ROW_NO_LSA_GROUP_ADMIN_URL: 15,
          ROW_NO_MAX_PAST_DAYS_ROS: 16,
          ROW_NO_MAX_ROS_DEL_MINS: 17,
          ROW_NO_SUPPORT_EMAIL_TO: 18,
          ROW_NO_SUPPORT_EMAIL_CCS: 19,
          ROW_NO_SUPPORT_EMAIL_BODY: 20,
          ROW_NO_UPDATE_AVAILABLE_TXT: 21,
          ROW_NO_JUST_UPDATED_TXT: 22,
          ROW_NO_RELEASE_NOTES: 23,
          ROW_NO_LAST_TEMPLATE_EDIT: 24,
          ROW_NO_LAST_MTR_LRNR_EDIT: 25,
          ROW_NO_LAST_TARGETS_EDIT: 26,
          ROW_NO_LAST_TARGETS_NAG: 27,
          ROW_NO_PENDING_ANNOUNCEMENT: 28,
          HANDLE_PENDING_ANNOUNCEMENT: "[PENDING_ANNOUNCEMENT]",
          ANNOUNCEMENT_QUEUE_LENGTH: 10
      }
  },
  MASTER_TEMPLATE: {
      NAME: 'MasterRosTemplate',
      REFS: {}
  },
  MASTER_UPGRADE_SCRIPT: {
      NAME: 'Master - Upgrade Script',
      REFS: {
          ROW_NO_FIRST_SCRIPT_ROW: 2,
          ROW_NO_LAST_SCRIPT_ROW: 501,
          COL_NO_SOURCE_VERSION: 1,
          COL_NO_SHEET_NAME: 2,
          COL_NO_SOURCE_RANGE: 3,
          COL_NO_DEST_PASTE_AT_COL: 4,
          COL_NO_DEST_PASTE_AT_ROW: 5,
          COL_NO_RENAME_SHEET_TO: 6,
          COL_NO_FIND_REPLACE_JSON: 7
      }
  },
  MASTER_ASSETS: {
      NAME: 'Master - Assets',
      REFS: {}
  },
  MASTER_LSAS: {
      NAME: 'Master - LSAs',
      REFS: {
          ROW_NO_FIRST_LSA: 2,
          ROW_NO_LAST_LSA: 101,
          COL_NO_STATUS_BAR: 1,
          COL_NO_LSA_NOS: 2,
          COL_NO_LSA_EMAIL: 3,
          COL_NO_LSA_NAME: 4,
          COL_NO_USERS_VERSION: 5,
          COL_NO_WORKBOOK_HYPERLINK: 6,
          COL_NO_PUSH_SVC_STATUS: 7,
          COL_NO_USERS_FILE_ID: 8
      }
  },
  MASTER_LEARNERS: {
      NAME: 'Master - Learners',
      HANDLE: "MASTER_LEARNERS",
      REFS: {
          ROW_NO_FIRST_LEARNER: 3,
          //COL_NO_STATUS_BAR     : 1,
          COL_NO_FORENAME: 1,
          HANDLE_FORENAME: "FORENAME",
          COL_NO_NICKNAME: 2,
          HANDLE_NICKNAME: "NICKNAME",
          COL_NO_SURNAME: 3,
          HANDLE_SURNAME: "SURNAME",
          COL_NO_LEARNER_ID: 4,
          HANDLE_LEARNER_ID: "LEARNER_ID",
          COL_NO_CATEGORY: 5,
          HANDLE_CATEGORY: "CATEGORY",
          COL_NO_EMAIL_ADDRESS: 6,
          HANDLE_EMAIL_ADDRESS: "EMAIL_ADDRESS",
          COL_NO_EXTERNAL_ID_1: 7,
          HANDLE_EXTERNAL_ID_1: "EXTERNAL_ID_1",
          COL_NO_LEARNER_DIR: 8,
          HANDLE_LEARNER_DIR: "LEARNER_DIR"
      }
  },
  MASTER_ANNOUNCEMENTS: {
      NAME: 'Master - Announcements',
      REFS: {
          ROW_NO_PROPOSED_ANCMNT: 6,
          COL_NO_PROPOSED_ANCMNT: 2
      }
  },
  MASTER_HELP: {
      NAME: 'Help Links',
      REFS: {
          ROW_NO_FIRST_RECORD: 3,
          ROW_NO_LAST_RECORD: 22,
          COL_NO_LINK_TEXT: 1,
          COL_NO_WHICH_SHEET: 2,
          COL_NO_LINK_URL: 3
      }
  },
  MOBILE_THIS_RECORD: {
      NAME: 'Mobile - Input',
      REFS: {
          RECORD_DATA_OFFSET: 1,
          COL_NO_RECORD_DATA: 2,
          COL_NO_CHECKBOXES: 2,
          COL_NO_SNAPSHOT: 4,
          ROW_NO_CBX_RELOAD: 1,
          ROW_NO_CBX_SAVE: 3,
          ROW_NO_STATUS_OVERRIDE: 3,
          ROW_NO_RECORD_NO: 2,
          ROW_NO_STATUS: 4
      }
  },
  MOBILE_MAIN: {
      NAME: 'Mobile - Main',
      REFS: {
          ALERT_BOX: {
              TOP_ROW: 5,
              B1_COL: 3,
              B2_COL: 5,
              HIDDEN_COL: 9
          },
          COL_NO_ALERT_BOX: 999999,
          ROW_NO_ALERT_BOX: 9999999,
          LESSON_DATE: { ROW_NO: 2, COL_NO: 7 },
          COL_NO_CHECKBOXES: 1,
          COL_NO_RECORD_NO: 2,
          COL_NO_STATUSES: 3,
          COL_NO_HIDDEN: 9,
          ROW_NO_NEW_DAY_CLEAN: 1,
          ROW_NO_GENERATE_ROS: 2,
          ROW_NO_UN_TICK_ALL: 9,
          ROW_NO_FIRST_RECORD: 10,
          ROW_NO_LAST_RECORD: 34
      }
  }
};