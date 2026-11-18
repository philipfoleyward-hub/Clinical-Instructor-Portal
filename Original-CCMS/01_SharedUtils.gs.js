/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * SHARED UTILITIES
 * 01_SharedUtils.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Generic helper functions used throughout the application.
 *
 * These functions should NOT contain:
 *   - grading rules
 *   - competency rules
 *   - HCI rules
 *   - remediation rules
 *   - document-specific logic
 *   - form-specific logic
 *
 * They should perform small, reusable technical tasks only.
 *
 * All shared utility functions use the prefix:
 *
 *   app_
 *
 * This helps prevent collisions with legacy functions while the
 * old scripts are still present.
 * =====================================================================
 */


/**
 * =====================================================================
 * TEXT NORMALIZATION
 * =====================================================================
 *
 * Example:
 *
 *   "  Clinical   2 "
 *
 * becomes:
 *
 *   "clinical 2"
 */
function appNormalize_(value) {

  return String(
    value === null ||
    value === undefined
      ? ''
      : value
  )
    .trim()
    .toLowerCase()
    .replace(
      /\s+/g,
      ' '
    );
}


/**
 * =====================================================================
 * TRUE / FALSE NORMALIZATION
 * =====================================================================
 *
 * Accepts:
 *
 *   true
 *   TRUE
 *   Yes
 *   yes
 *   Y
 *   1
 *   Active
 *
 */
function appIsTrue_(value) {

  if (value === true) {
    return true;
  }

  const normalized =
    appNormalize_(value);

  return (
    normalized === 'true' ||
    normalized === 'yes' ||
    normalized === 'y' ||
    normalized === '1' ||
    normalized === 'active'
  );
}


/**
 * =====================================================================
 * REQUIRED SHEET LOOKUP
 * =====================================================================
 *
 * Use when a missing sheet should stop the operation.
 */
function appGetSheet_(ss, sheetName) {

  if (!ss) {
    throw new Error(
      'Spreadsheet is required.'
    );
  }

  const sheet =
    ss.getSheetByName(
      sheetName
    );

  if (!sheet) {

    throw new Error(
      'Required sheet not found: "' +
      sheetName +
      '".'
    );
  }

  return sheet;
}


/**
 * =====================================================================
 * OPTIONAL SHEET LOOKUP
 * =====================================================================
 *
 * Use when it is acceptable for a sheet not to exist.
 *
 * Returns:
 *
 *   Sheet object
 *
 * or
 *
 *   null
 */
function appFindSheet_(ss, sheetName) {

  if (!ss) {
    return null;
  }

  return (
    ss.getSheetByName(
      sheetName
    ) || null
  );
}


/**
 * =====================================================================
 * SHEET → OBJECTS
 * =====================================================================
 *
 * Converts:
 *
 *   Student ID | Student | Active
 *   123        | Jane    | TRUE
 *
 * into:
 *
 * {
 *   "Student ID": 123,
 *   "Student": "Jane",
 *   "Active": true
 * }
 *
 */
function appSheetToObjects_(sheet) {

  if (!sheet) {
    return [];
  }

  const values =
    sheet
      .getDataRange()
      .getValues();

  if (values.length < 2) {
    return [];
  }

  const headers =
    values[0].map(
      header =>
        String(
          header || ''
        ).trim()
    );

  return values
    .slice(1)
    .filter(
      row =>
        row.some(
          cell =>
            cell !== '' &&
            cell !== null &&
            cell !== undefined
        )
    )
    .map(row => {

      const object = {};

      headers.forEach(
        (header, index) => {

          if (!header) {
            return;
          }

          object[header] =
            row[index];
        }
      );

      return object;
    });
}


/**
 * =====================================================================
 * HEADER LOOKUP
 * =====================================================================
 *
 * Returns an object such as:
 *
 * {
 *   "student id": 0,
 *   "student": 1,
 *   "clinical term": 2
 * }
 */
function appHeaderMap_(headerRow) {

  const map = {};

  if (!Array.isArray(headerRow)) {
    return map;
  }

  headerRow.forEach(
    (header, index) => {

      const key =
        appNormalize_(header);

      if (key) {
        map[key] = index;
      }
    }
  );

  return map;
}


/**
 * =====================================================================
 * DATE ONLY
 * =====================================================================
 *
 * Returns a Date with the time removed.
 *
 * Invalid / blank values return null.
 */
function appDateOnly_(value) {

  if (!value) {
    return null;
  }

  let date;

  if (value instanceof Date) {

    date =
      new Date(
        value.getTime()
      );

  } else {

    date =
      new Date(value);
  }

  if (
    isNaN(
      date.getTime()
    )
  ) {

    return null;
  }

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );
}


/**
 * =====================================================================
 * DATE VALUE
 * =====================================================================
 *
 * Useful for sorting or comparing dates.
 */
function appDateValue_(value) {

  const date =
    appDateOnly_(value);

  return date
    ? date.getTime()
    : 0;
}


/**
 * =====================================================================
 * DATE WITHIN RANGE
 * =====================================================================
 *
 * Inclusive:
 *
 *   start <= date <= end
 */
function appDateWithin_(
  value,
  start,
  end
) {

  const date =
    appDateOnly_(value);

  const startDate =
    appDateOnly_(start);

  const endDate =
    appDateOnly_(end);

  if (
    !date ||
    !startDate ||
    !endDate
  ) {

    return false;
  }

  return (
    date.getTime() >=
      startDate.getTime() &&
    date.getTime() <=
      endDate.getTime()
  );
}


/**
 * =====================================================================
 * SAFE FILE NAME TEXT
 * =====================================================================
 *
 * Removes characters Google Drive / Windows filenames may dislike.
 */
function appSafeFileText_(value) {

  return String(
    value === null ||
    value === undefined
      ? ''
      : value
  )
    .replace(
      /[\\\/:*?"<>|]/g,
      ''
    )
    .replace(
      /\s+/g,
      ' '
    )
    .trim();
}


/**
 * =====================================================================
 * REGEX ESCAPE
 * =====================================================================
 *
 * Useful when replacing literal template tokens such as:
 *
 * {{NAME}}
 */
function appEscapeRegex_(value) {

  return String(
    value === null ||
    value === undefined
      ? ''
      : value
  )
    .replace(
      /[.*+?^${}()|[\]\\]/g,
      '\\$&'
    );
}


/**
 * =====================================================================
 * APPEND OBJECT TO SHEET BY HEADER NAME
 * =====================================================================
 *
 * Allows:
 *
 * {
 *   "Student": "Jane Doe",
 *   "Result": "Pass"
 * }
 *
 * to be written into the correct columns regardless of column order.
 *
 * Headers not present in the object receive a blank value.
 */
function appAppendByHeaders_(
  sheet,
  record
) {

  if (!sheet) {

    throw new Error(
      'A sheet is required.'
    );
  }

  if (!record) {

    throw new Error(
      'A record is required.'
    );
  }

  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {

    throw new Error(
      'Sheet "' +
      sheet.getName() +
      '" has no headers.'
    );
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(
        header =>
          String(
            header || ''
          ).trim()
      );

  const row =
    headers.map(header => {

      if (
        Object.prototype
          .hasOwnProperty
          .call(
            record,
            header
          )
      ) {

        return record[header];
      }

      return '';
    });

  sheet.appendRow(row);
}