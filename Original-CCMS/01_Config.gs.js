/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * CENTRAL APPLICATION CONFIGURATION
 * 01_Config.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Single authoritative location for:
 *
 *   - Sheet names
 *   - Google Form IDs
 *   - Google Doc template IDs
 *   - Drive folder IDs
 *   - Standard application terminology
 *   - Shared source labels
 *
 * IMPORTANT
 * ---------
 * This file contains CONFIGURATION ONLY.
 *
 * Do NOT:
 *   - read spreadsheet data here
 *   - modify forms here
 *   - calculate grades here
 *   - calculate requirements here
 *   - generate documents here
 *   - send emails here
 *
 * Other scripts should READ values from APP_CONFIG.
 * =====================================================================
 */

const APP_CONFIG = {
  SHEETS: {
    STUDENTS: 'Students',
    CLINICAL_SITES: 'Clinical Sites',
    EVALUATORS: 'Evaluators',
    EXAMINATIONS: 'Examinations',
    ANATOMICAL_AREAS: 'Anatomical Areas',
    SEMESTER_RULES: 'Semester Rules',
    PROGRAM_RULES: 'Program Rules',
    SCORING_WEIGHTS: 'Scoring Weights',
    IMAGE_SCORING_RULES: 'Image Scoring Rules',
    CLINICAL_TRACKING: 'Clinical Tracking',
    STUDENT_REQUIREMENT_STATUS: 'Student Requirement Status',
    HISTORICAL_IMPORT: 'Historical Import',
    FORM_RESPONSES_1: 'Form Responses 1',
    FORM_RESPONSES_2: 'Form Responses 2',
    CLINICAL_CALENDAR: 'Clinical Calendar',
    STUDENT_SITE_ASSIGNMENTS: 'Student Site Assignments',
    STUDENT_ROOM_ROTATIONS: 'Student Room Rotations',
    STUDENT_SITE_PREFERENCES: 'Student Site Preferences',
    CI_AVAILABILITY: 'CI Availability',
    CI_ASSIGNMENTS: 'CI Assignments',
    ATTENDANCE_EVENTS: 'Attendance Events',
    ATTENDANCE_RULES: 'Attendance Rules',
    USERS: 'Users'
  },

  FORMS: {
    DIGITAL_CLINICAL_EVALUATION: '1TGm4ciMz3YkKv902gON-FmnofNUORqMmKySVA5pPrsY',
    QR_SUBMISSION: '1LYI9PMAVyEX-A0XciEl72nWDJ99Knt0yI9HlDq-Ql_c'
  },

  TEMPLATES: {
    CLINICAL_COMPETENCY_EVALUATION: '1w2nHGKLtUynCfP-mcEb3LS9gXas-_BKgpkrcwvj1cTc',
    SEMESTER_REQUIREMENTS: {
      CLINICAL_1: '',
      CLINICAL_2: '1KrZL7kH_LtdqLmSe9BP9SFF-XRfWRe0D6o8eN3v9Gcc',
      CLINICAL_3: '',
      CLINICAL_4: '',
      CLINICAL_5: ''
    },
    PROGRAM_REQUIREMENTS: '1Pa31xYCjEjwODytBHHw1UK29bHhbqgQ8ZWVNYi094Y0'
  },

  NOTIFICATIONS: {
    COORDINATOR_RECIPIENTS: [
      'philipfoleyward@sunyorange.edu','nicolemurray@sunyorange.edu','micheleterracino@sunyorange.edu'
    ],
    DASHBOARD_URL: '',
    SENDER_NAME: 'Clinical Evaluation System'
  },

  DRIVE: {
    CURRENT_EVALUATION_OUTPUT_FOLDER_ID: '17yQMK3lrLYdC97kA-UxxcSdmqoU7FDNl',
    CLINICAL_EVALUATIONS_ROOT_ID: '',
    CLINICAL_EVALUATIONS_ROOT_NAME: 'Clinical Evaluations'
  },

  TERMINOLOGY: {
    EVALUATOR_TYPES: {
      COLLEGE_CI: 'College CI',
      HOSPITAL_CI: 'Hospital CI'
    },
    EVALUATION_TYPES: {
      COMPETENCY: 'Competency',
      ISO: 'ISO',
      SSO: 'SSO',
      SIMULATION: 'Simulation'
    },
    RESULTS: {
      PASS: 'Pass',
      FAIL: 'Fail'
    },
    REVIEW_STATUS: {
      PENDING: 'Pending',
      APPROVED: 'Approved',
      REJECTED: 'Rejected',
      REVISION_REQUIRED: 'Revision Required'
    }
  },

  SOURCES: {
    DIGITAL_EVALUATION: 'Digital Clinical Evaluation',
    QR_SUBMISSION: 'Clinical Evaluation Submission Form',
    HISTORICAL_IMPORT: 'Historical Import'
  },

  SETTINGS: {
    CREATE_MISSING_STUDENT_FOLDERS: true,
    DEFAULT_REVIEW_STATUS: 'Pending',
    HISTORICAL_IMPORT_REQUIRES_LIVE_REVIEW: false
  }
};
