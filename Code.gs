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

  // Most rows one edit event will process. A bulk paste of more than this
  // leaves the remainder in "New"; the daily scan then flags them rather than
  // the trigger dying mid-way against the execution time limit.
  MAX_ROWS_PER_EDIT: 25,

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

/* =========================================================================
 * 3. onEdit path — send the welcome email
 * ========================================================================= */

/**
 * Edit entry point.
 *
 * IMPORTANT — this must be installed as an *installable* trigger (run
 * `installTriggers` once, or use the Onboarding menu). A function named
 * onEdit also runs automatically as a simple trigger, and simple triggers
 * execute without authorization, so they are not permitted to call MailApp at
 * all. Google fires both, so we tell them apart and let the simple one fall
 * straight through: only installable trigger events carry `triggerUid`.
 *
 * Without that check the simple invocation would throw a permissions error on
 * every single edit and bury the real executions in noise.
 *
 * @param {!Object} e Edit event object supplied by Apps Script.
 */
function onEdit(e) {
  if (!e || !e.range) {
    return;
  }
  if (!e.triggerUid) {
    // Simple-trigger invocation: unauthorized for MailApp. The installable
    // trigger handles this same edit a moment later.
    return;
  }
  handleRosterEdit_(e);
}

/**
 * Decides whether an edit is one we care about, and processes the rows it
 * touched.
 *
 * We act only when the edited range overlaps the Status column, which covers
 * both ways a row becomes actionable: a recruiter typing "New" into Status,
 * and a whole row being pasted in with Status already filled.
 *
 * @param {!Object} e Edit event object.
 */
function handleRosterEdit_(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== getSetting_('ROSTER_SHEET_NAME')) {
    return;
  }

  var headerMap = getHeaderMap_(sheet);
  var statusColumn = headerMap[normalizeKey_(COLUMNS.STATUS)];
  if (!statusColumn) {
    // No Status column means this is not the roster we expect. Stay quiet
    // rather than throwing on every edit to an unrelated sheet.
    return;
  }

  // Did this edit touch the Status column at all?
  var firstColumn = e.range.getColumn();
  var lastColumn = firstColumn + e.range.getNumColumns() - 1;
  if (statusColumn < firstColumn || statusColumn > lastColumn) {
    return;
  }

  var firstRow = e.range.getRow();
  var rowCount = e.range.getNumRows();

  for (var offset = 0; offset < rowCount; offset++) {
    var rowNumber = firstRow + offset;

    // Row 1 is headers.
    if (rowNumber < 2) {
      continue;
    }
    if (offset >= CONFIG.MAX_ROWS_PER_EDIT) {
      // Remaining rows keep Status "New" and are picked up by the daily scan.
      console.warn(
        'Edit touched more than ' + CONFIG.MAX_ROWS_PER_EDIT + ' rows; ' +
        'stopping at row ' + rowNumber + '. Remaining rows will be reported ' +
        'by the daily stuck-row check.');
      break;
    }

    // Cheap pre-filter outside the lock. The authoritative check happens
    // again inside sendWelcomeForRow_ once we actually hold the lock.
    var status = canonicalStatus_(sheet.getRange(rowNumber, statusColumn).getValue());
    if (status !== STATUS.NEW) {
      continue;
    }

    sendWelcomeForRow_(sheet, headerMap, rowNumber);
  }
}

/**
 * Sends the welcome email for one row, then marks it Emailed.
 *
 * Concurrency is the real hazard here. Several people edit this sheet at once,
 * and two edits to the same row can have their trigger executions overlap. So:
 *
 *   1. Take a document-scoped lock (spreadsheet-wide, across all users).
 *   2. Re-read Status *inside* the lock and bail if it is no longer "New".
 *   3. Send, write Status + timestamp, then flush before releasing.
 *
 * Step 2 is what makes this idempotent: whoever loses the race sees "Emailed"
 * and skips instead of sending a second copy. Step 3 makes sure that write is
 * committed and visible before the next contender is allowed in.
 *
 * @param {!GoogleAppsScript.Spreadsheet.Sheet} sheet Roster sheet.
 * @param {!Object<string, number>} headerMap From getHeaderMap_.
 * @param {number} rowNumber 1-based sheet row.
 * @return {string} 'sent' or 'skipped'.
 */
function sendWelcomeForRow_(sheet, headerMap, rowNumber) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    // Someone else holds the lock. Their execution is handling this row, or
    // the row stays "New" and the daily scan reports it. Never send without
    // the lock — that is exactly how duplicates happen.
    console.warn('Could not acquire lock for row ' + rowNumber + '; skipping.');
    return 'skipped';
  }

  try {
    // Authoritative re-read. The values in the edit event may already be
    // stale by the time we get here.
    var row = readRosterRow_(sheet, headerMap, rowNumber);

    // --- Idempotency gate -------------------------------------------------
    if (row.status === STATUS.EMAILED) {
      return 'skipped';
    }
    if (row.status !== STATUS.NEW) {
      return 'skipped';
    }
    if (row.sentAt) {
      // Belt and braces: a send timestamp means mail already went out, even
      // if the Status cell was hand-edited back to "New" afterwards.
      return 'skipped';
    }

    // --- Validation -------------------------------------------------------
    if (!row.email) {
      throw new Error('Email column is empty.');
    }
    if (!isValidEmail_(row.email)) {
      throw new Error('Malformed email address: "' + row.email + '"');
    }
    if (!row.name) {
      throw new Error('Name column is empty.');
    }

    // Refuse to start a send we cannot finish. Hitting the quota mid-run
    // throws an opaque error; checking first produces a message an admin can
    // act on.
    var remaining = MailApp.getRemainingDailyQuota();
    if (remaining < CONFIG.MIN_QUOTA_BUFFER) {
      throw new Error(
        'Gmail daily send quota nearly exhausted (' + remaining +
        ' remaining). Not sending; retry after the quota resets.');
    }

    // --- Send -------------------------------------------------------------
    var message = buildWelcomeEmail_(row);
    MailApp.sendEmail({
      to: row.email,
      subject: message.subject,
      body: message.body,
      htmlBody: message.htmlBody,
      name: getSetting_('COMPANY_NAME'),
      replyTo: getSetting_('ONBOARDING_CONTACT')
    });

    markRowEmailed_(sheet, headerMap, rowNumber, new Date());

    // Commit the status write before releasing the lock, so a concurrent
    // execution waiting on it reads "Emailed" and not a stale "New".
    SpreadsheetApp.flush();
    return 'sent';
  } catch (err) {
    // Any failure — validation, quota, a Gmail hiccup — lands here. The row
    // is parked in "Error" so it is visible in the sheet and picked up by the
    // daily scan, and the details go to the Error Log tab. We deliberately do
    // not rethrow: one bad row must not abort the other rows in this edit.
    markRowError_(sheet, headerMap, rowNumber);
    logError_(rowNumber, err, row ? row.email : '');
    SpreadsheetApp.flush();
    return 'failed';
  } finally {
    lock.releaseLock();
  }
}

/**
 * Builds the welcome email from a row. Kept as a pure function of the row so
 * the wording can be reviewed and changed without touching send logic.
 *
 * @param {!Object} row From readRosterRow_.
 * @return {{subject: string, body: string, htmlBody: string}}
 */
function buildWelcomeEmail_(row) {
  var company = getSetting_('COMPANY_NAME');
  var contact = getSetting_('ONBOARDING_CONTACT');
  var startDate = formatStartDate_(row.startDate);

  // First name only reads warmer than the full legal name in the sheet.
  var firstName = row.name.split(/\s+/)[0];

  var startLine = startDate
    ? 'Your first day is ' + startDate + '.'
    : 'We will confirm your start date shortly.';

  var subject = 'Welcome to ' + company + ', ' + firstName + '!';

  var body = [
    'Hi ' + firstName + ',',
    '',
    'Welcome to ' + company + '. We are glad to have you joining the clinical team.',
    '',
    startLine,
    '',
    'Before then, a few things to expect:',
    '  - Your onboarding packet and credentialing paperwork, sent separately.',
    '  - Access to the scheduling and EHR systems, set up in your first week.',
    '  - An introduction to your supervising clinician.',
    '',
    'If anything above looks wrong, or your start date has changed, just reply',
    'to this message and we will sort it out.',
    '',
    'Warmly,',
    'The ' + company + ' Onboarding Team',
    contact
  ].join('\n');

  var htmlBody = [
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#222;">',
    '<p>Hi ' + escapeHtml_(firstName) + ',</p>',
    '<p>Welcome to ' + escapeHtml_(company) + '. We are glad to have you joining the clinical team.</p>',
    '<p><strong>' + escapeHtml_(startLine) + '</strong></p>',
    '<p>Before then, a few things to expect:</p>',
    '<ul>',
    '<li>Your onboarding packet and credentialing paperwork, sent separately.</li>',
    '<li>Access to the scheduling and EHR systems, set up in your first week.</li>',
    '<li>An introduction to your supervising clinician.</li>',
    '</ul>',
    '<p>If anything above looks wrong, or your start date has changed, just reply ',
    'to this message and we will sort it out.</p>',
    '<p>Warmly,<br>The ' + escapeHtml_(company) + ' Onboarding Team<br>',
    '<a href="mailto:' + encodeURI(contact) + '">' + escapeHtml_(contact) + '</a></p>',
    '</div>'
  ].join('');

  return {subject: subject, body: body, htmlBody: htmlBody};
}

/**
 * Escapes text interpolated into the HTML email body. Names come from a sheet
 * that many people can type into, so treat them as untrusted input rather than
 * letting a stray angle bracket break the markup.
 *
 * @param {string} text Raw text.
 * @return {string} HTML-escaped text.
 */
function escapeHtml_(text) {
  return String(text === null || text === undefined ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Marks a row as successfully emailed: Status becomes "Emailed" and the
 * Email Sent Timestamp column is stamped. Together these are what later runs
 * read to decide the row is already done.
 *
 * @param {!GoogleAppsScript.Spreadsheet.Sheet} sheet Roster sheet.
 * @param {!Object<string, number>} headerMap From getHeaderMap_.
 * @param {number} rowNumber 1-based sheet row.
 * @param {!Date} when Send time.
 */
function markRowEmailed_(sheet, headerMap, rowNumber, when) {
  sheet.getRange(rowNumber, requireColumn_(headerMap, COLUMNS.STATUS))
    .setValue(STATUS.EMAILED);
  sheet.getRange(rowNumber, requireColumn_(headerMap, COLUMNS.SENT_AT))
    .setValue(when);
}

/* =========================================================================
 * 4. Error logging
 * ========================================================================= */

/** Header row for the Error Log tab, created on first use. */
var ERROR_LOG_HEADERS = [
  'Timestamp', 'Row', 'Therapist Email', 'Error Message', 'Details'
];

/**
 * Parks a row in "Error" status so the failure is visible in the sheet itself,
 * not only in a log tab nobody opens.
 *
 * Note the Email Sent Timestamp is left blank on purpose: an errored row has
 * not been emailed, and leaving it empty means a retry (setting Status back to
 * "New") is not blocked by the already-sent guard.
 *
 * Deliberately tolerant — this runs from a catch block, and a throw here would
 * mask the original error.
 *
 * @param {!GoogleAppsScript.Spreadsheet.Sheet} sheet Roster sheet.
 * @param {!Object<string, number>} headerMap From getHeaderMap_.
 * @param {number} rowNumber 1-based sheet row.
 */
function markRowError_(sheet, headerMap, rowNumber) {
  try {
    var statusColumn = headerMap[normalizeKey_(COLUMNS.STATUS)];
    if (statusColumn) {
      sheet.getRange(rowNumber, statusColumn).setValue(STATUS.ERROR);
    }
  } catch (err) {
    console.error(
      'Could not set Error status on row ' + rowNumber + ': ' + err);
  }
}

/**
 * Appends one failure to the Error Log tab, creating the tab if needed.
 *
 * Like markRowError_, this must never throw. It is the last link in the chain:
 * if logging the error also threw, the failure would vanish entirely, which is
 * the silent-failure mode this whole project exists to prevent. A failure here
 * falls back to console.error, which is still visible in the Apps Script
 * execution log.
 *
 * @param {number} rowNumber Roster row the failure relates to.
 * @param {*} error The caught exception.
 * @param {string=} email Address involved, when known — useful context.
 */
function logError_(rowNumber, error, email) {
  var message = (error && error.message) ? error.message : String(error);
  var details = (error && error.stack) ? String(error.stack) : '';

  try {
    var sheet = getOrCreateSheet_(
      getSetting_('ERROR_LOG_SHEET_NAME'), ERROR_LOG_HEADERS);
    sheet.appendRow([
      new Date(),
      rowNumber,
      email || '',
      message,
      // Stack traces can be long; the cell limit is 50k characters.
      details.slice(0, 4000)
    ]);
  } catch (err) {
    console.error(
      'Failed to write to the Error Log tab (' + err + '). Original error on ' +
      'row ' + rowNumber + ': ' + message);
  }

  // Always mirror to the execution log so Stackdriver has the full picture
  // even when the spreadsheet write succeeded.
  console.error('Row ' + rowNumber + ': ' + message);
}
