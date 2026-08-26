/**
 * ===========================================================================
 * Therapist Onboarding — Automated Welcome Email
 * ===========================================================================
 *
 * Replaces a manual "someone remembers to email the new hire" process with a
 * Sheet-driven one. A recruiter sets a row's Status to "New"; the script sends
 * the welcome email, flips Status to "Emailed" and stamps the send time.
 *
 * The sheet is edited by several people at once, so every write path assumes
 * concurrency: the send is guarded by a document lock and re-checks Status
 * inside that lock before it does anything irreversible.
 *
 * Single file, no external dependencies. Read top to bottom:
 *   1. CONFIG / constants
 *   2. Sheet + schema helpers     <- you are here
 *   3. onEdit path (send the email)
 *   4. Error logging
 *   5. Daily stuck-row scan + admin alert
 *   6. Heartbeat / System Health
 *   7. Trigger installation + menu
 */

/**
 * Tunables. Anything an admin might want to change after deploy can also be
 * set as a Script Property (Project Settings > Script Properties), which wins
 * over the value here — that way you can point the alerts at a different
 * person without editing and redeploying code.
 */
var CONFIG = {
  // Tab that holds the therapist roster.
  ROSTER_SHEET_NAME: 'Onboarding',

  // Tabs the script creates on demand if they are missing.
  ERROR_LOG_SHEET_NAME: 'Error Log',
  HEALTH_SHEET_NAME: 'System Health',

  // Where stuck-row alerts go. Override with Script Property ADMIN_EMAIL.
  ADMIN_EMAIL: 'onboarding-admin@example.com',

  // Optional. Set Script Property SLACK_WEBHOOK_URL to also post alerts to
  // Slack. Left empty, the script just sends the admin email.
  SLACK_WEBHOOK_URL: '',

  // A row sitting in "New" longer than this is considered stuck.
  STALE_NEW_ROW_HOURS: 24,

  // How long an edit will wait for the document lock before giving up.
  LOCK_TIMEOUT_MS: 30000,

  // Refuse to send (and alert instead) when the daily Gmail quota drops below
  // this. Leaves headroom so the automation degrades loudly, not silently.
  MIN_QUOTA_BUFFER: 5,

  // Keep the health log from growing without bound.
  HEALTH_LOG_MAX_ROWS: 400,

  // Branding used in the email template.
  COMPANY_NAME: 'Bright Path Therapy',
  ONBOARDING_CONTACT: 'onboarding@brightpaththerapy.example'
};

/** The only three values the Status column is allowed to hold. */
var STATUS = {
  NEW: 'New',
  EMAILED: 'Emailed',
  ERROR: 'Error'
};

/**
 * Column headers, referenced by name rather than by index on purpose. In a
 * shared sheet someone will eventually insert a column in the middle, and
 * hard-coded indexes would start emailing the wrong field.
 */
var COLUMNS = {
  NAME: 'Name',
  EMAIL: 'Email',
  START_DATE: 'Start Date',
  STATUS: 'Status',
  SENT_AT: 'Email Sent Timestamp'
};

/* =========================================================================
 * 2. Sheet + schema helpers
 * ========================================================================= */

/**
 * Reads a setting, preferring the Script Property of the same name so that
 * runtime config (admin address, Slack URL) never has to live in the repo.
 *
 * @param {string} key Key in CONFIG, and the Script Property name.
 * @return {string} The property value if set and non-empty, else CONFIG[key].
 */
function getSetting_(key) {
  var override;
  try {
    override = PropertiesService.getScriptProperties().getProperty(key);
  } catch (err) {
    // Property store is unavailable in some restricted contexts; fall back
    // to the in-code default rather than failing the whole run.
    override = null;
  }
  if (override && String(override).trim()) {
    return String(override).trim();
  }
  return CONFIG[key];
}

/**
 * Fetches a tab by name, or null if it does not exist. Callers decide whether
 * a missing tab is fatal (the roster) or something to create (the log tabs).
 *
 * @param {string} name Tab name.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSheet_(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

/**
 * Fetches a tab, creating it with the given header row if absent. Used for the
 * Error Log and System Health tabs so a fresh copy of the spreadsheet is
 * self-provisioning — an admin cannot forget to add them.
 *
 * @param {string} name Tab name.
 * @param {!Array<string>} headers Header row to write on creation.
 * @return {!GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    return sheet;
  }
  sheet = ss.insertSheet(name);
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Builds a header-name -> 1-based column index map from row 1.
 *
 * Lookup keys are lower-cased and trimmed so that "Email", "email " and
 * " EMAIL" all resolve. Co-editors retype headers more often than you would
 * expect, and a trailing space should not take the automation down.
 *
 * @param {!GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to inspect.
 * @return {!Object<string, number>} Normalized header name to column index.
 */
function getHeaderMap_(sheet) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    return {};
  }
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = normalizeKey_(headers[i]);
    // First occurrence wins, so a stray duplicate header further right cannot
    // silently shadow the real column.
    if (key && !map.hasOwnProperty(key)) {
      map[key] = i + 1;
    }
  }
  return map;
}

/**
 * Resolves a column index from a header map, throwing a message that names the
 * missing header. Surfacing "Status column not found" beats a generic
 * undefined-index error three functions deeper.
 *
 * @param {!Object<string, number>} headerMap From getHeaderMap_.
 * @param {string} header Header to look up.
 * @return {number} 1-based column index.
 */
function requireColumn_(headerMap, header) {
  var index = headerMap[normalizeKey_(header)];
  if (!index) {
    throw new Error(
      'Required column "' + header + '" was not found in the header row of "' +
      CONFIG.ROSTER_SHEET_NAME + '". Check for a renamed or deleted column.');
  }
  return index;
}

/**
 * Lower-cases and trims a header or status value for tolerant comparison.
 *
 * @param {*} value Raw cell value.
 * @return {string} Normalized string.
 */
function normalizeKey_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Maps a raw Status cell to one of the canonical STATUS values, or '' when the
 * cell is blank or holds something unrecognized. Comparing through this keeps
 * "new", "New " and "NEW" all meaning the same thing.
 *
 * @param {*} value Raw Status cell value.
 * @return {string} A STATUS value, or '' if not recognized.
 */
function canonicalStatus_(value) {
  var normalized = normalizeKey_(value);
  var keys = Object.keys(STATUS);
  for (var i = 0; i < keys.length; i++) {
    if (normalizeKey_(STATUS[keys[i]]) === normalized) {
      return STATUS[keys[i]];
    }
  }
  return '';
}

/**
 * Reads one roster row into a plain object keyed by the COLUMNS names.
 *
 * Always reads live from the sheet rather than trusting a cached copy, because
 * with several people editing at once the values captured when an edit event
 * fired may already be out of date by the time we act on them.
 *
 * @param {!GoogleAppsScript.Spreadsheet.Sheet} sheet Roster sheet.
 * @param {!Object<string, number>} headerMap From getHeaderMap_.
 * @param {number} rowNumber 1-based sheet row.
 * @return {!Object} Row values plus the resolved rowNumber.
 */
function readRosterRow_(sheet, headerMap, rowNumber) {
  var lastColumn = sheet.getLastColumn();
  var values = sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];

  /** Pulls a single named column out of the row we just read. */
  function valueOf(header) {
    var index = headerMap[normalizeKey_(header)];
    return index ? values[index - 1] : '';
  }

  return {
    rowNumber: rowNumber,
    name: String(valueOf(COLUMNS.NAME) || '').trim(),
    email: String(valueOf(COLUMNS.EMAIL) || '').trim(),
    startDate: valueOf(COLUMNS.START_DATE),
    status: canonicalStatus_(valueOf(COLUMNS.STATUS)),
    rawStatus: valueOf(COLUMNS.STATUS),
    sentAt: valueOf(COLUMNS.SENT_AT)
  };
}

/**
 * Conservative email syntax check. This does not prove the address exists — it
 * only rejects the obvious typos ("jane.doe@", "jane doe@x.com") before they
 * become a failed send. A well-formed but wrong address will still bounce, and
 * bounces are not visible to Apps Script; see README.
 *
 * @param {string} email Candidate address.
 * @return {boolean} True when the address is plausibly well formed.
 */
function isValidEmail_(email) {
  return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]{2,}$/.test(String(email || '').trim());
}

/**
 * Formats a Start Date for display in the email body. Sheet date cells arrive
 * as Date objects; anything else (a typed string like "TBD") is passed through
 * untouched so the email still reads sensibly.
 *
 * @param {*} value Raw Start Date cell value.
 * @return {string} Human-readable date, or '' when blank.
 */
function formatStartDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
  }
  return String(value === null || value === undefined ? '' : value).trim();
}

/**
 * Formats a timestamp for log tabs and alert bodies, in the script's timezone
 * so every reader sees the same wall-clock time.
 *
 * @param {!Date} date Timestamp to format.
 * @return {string} e.g. "2026-08-26 14:03:11 EDT".
 */
function formatTimestamp_(date) {
  return Utilities.formatDate(
    date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss z');
}
