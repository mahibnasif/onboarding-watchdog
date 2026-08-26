# Therapist Onboarding — Automated Welcome Email

A Google Apps Script bound to the onboarding spreadsheet. It replaces the manual
"someone remembers to email the new hire" step, which was unreliable precisely
because several people edit the roster at once and each assumed another had sent
it.

Two files, no dependencies:

| File | What it is |
| --- | --- |
| `Code.gs` | The whole system, in seven numbered sections. |
| `appsscript.json` | Manifest pinning the runtime and OAuth scopes. |

---

## What it does

The roster tab (`Onboarding` by default) has five columns:

| Name | Email | Start Date | Status | Email Sent Timestamp |
| --- | --- | --- | --- | --- |
| Freya Ashworth | freya@example.com | 2026-09-08 | New | |

A recruiter sets **Status** to `New`. From there:

1. An installable `onEdit` trigger fires and sees the Status change.
2. It takes a document lock, re-reads the row, and validates name, address and
   remaining send quota.
3. It sends a personalized welcome email — first name, formatted start date,
   what to expect before day one — via `MailApp`.
4. It writes Status `Emailed` and stamps **Email Sent Timestamp**.

If anything fails, Status becomes `Error` and the exception is appended to an
**Error Log** tab with timestamp, row number, address and stack.

Once a day, a time-driven check scans for rows that never made it to `Emailed`,
emails an admin about them, and appends a row to a **System Health** tab.

The Error Log and System Health tabs are created automatically on first use.

---

## Setup

1. Open the spreadsheet → **Extensions → Apps Script**. Paste `Code.gs`. Enable
   the manifest under **Project Settings → Show `appsscript.json`** and paste
   that too.
2. Set `timeZone` in `appsscript.json` to the clinic's timezone. It drives the
   daily run hour and every timestamp shown to a human.
3. In **Project Settings → Script Properties**, set:
   - `ADMIN_EMAIL` — where stuck-row alerts go. **Required.**
   - `SLACK_WEBHOOK_URL` — optional; mirrors alerts to Slack.
   - Any other `CONFIG` key you want to override without editing code.
4. Reload the spreadsheet, then use **Onboarding → Install / re-install
   triggers**, and accept the authorization prompt.
5. Verify with **Onboarding → Show system status**.

**Run step 4 from the account that should appear as the sender.** Installable
triggers execute as whoever created them, and that account's Gmail quota is what
gets consumed.

### Why `installTriggers` is mandatory, not a convenience

A function named `onEdit` runs automatically as a *simple* trigger — but simple
triggers execute without authorization and are **not permitted to call `MailApp`
at all**. Nothing will ever be sent by the simple trigger.

Google fires both the simple and the installable trigger for the same edit, so
`onEdit` checks `e.triggerUid` (present only on installable events) and returns
immediately on the simple invocation. Without that guard, every keystroke in the
sheet would throw a permissions error and bury the real executions in noise.

If you skip step 4, the sheet looks completely normal and no email is ever sent.

---

## What could go wrong

### Duplicate sends from concurrent edits

The real hazard, given how this sheet is used. Two people editing the same row
can have their trigger executions overlap, and both could read Status `New` and
both send.

Handled by double-checked locking in `sendWelcomeForRow_`: take a
document-scoped lock (spreadsheet-wide, across all users), then **re-read Status
from the sheet inside the lock**, and abandon if it is no longer `New`. Whoever
loses the race sees `Emailed` and skips. The status write is flushed before the
lock is released, so the next contender cannot read a stale value. A row that
already carries a send timestamp is skipped even if its Status was hand-edited
back to `New`.

If the lock cannot be acquired within 30 seconds the row is left alone rather
than sent without protection — it stays `New` and the daily scan reports it.

### Gmail sending quota limits

Consumer accounts get roughly 100 recipients/day, Workspace accounts 1,500. A
bulk paste of new hires can exhaust that.

Every send checks `MailApp.getRemainingDailyQuota()` first and refuses below a
small buffer, producing an `Error` row that says so rather than an opaque
mid-send exception. The daily alert also carries an explicit warning when the
quota is nearly gone. Quota resets on a rolling 24-hour basis; set the affected
rows back to `New` to retry.

### Malformed email addresses

Typed by hand into a shared sheet, so expect `jane.doe@`, stray spaces, two
addresses in one cell. Addresses are syntax-checked before the send is
attempted; failures become `Error` rows naming the bad address.

**Limitation worth naming:** this catches malformed addresses, not wrong ones. A
well-formed address for a person who does not exist will be accepted, sent, and
bounce — and Apps Script cannot see bounces. The row will read `Emailed`. The
only defence is a human noticing the bounce in the sending account's inbox.

### The trigger silently failing after a permissions change

The dangerous one. If someone revokes the script's authorization, the owning
account's password changes in a way that invalidates the grant, admin policy
changes, the account is suspended, or a trigger is deleted — **no code of ours
runs**. No error is thrown, because nothing executes. The roster looks perfectly
normal. Rows just sit at `New` forever while everyone assumes the automation
handled them.

Also in this category: an OAuth scope change in `appsscript.json` forces every
editor to re-authorize, and until they do, their edits do nothing.

### Others handled in passing

- **Someone reorders or renames a column.** Columns are resolved by header name,
  not index, with trimming and case-insensitive matching. A missing required
  column produces an error naming it.
- **A bulk paste of hundreds of rows.** Capped at 25 rows per edit event so the
  trigger cannot die part-way against the execution time limit. The remainder
  stay `New` and are reported by the daily scan — never silently dropped.
- **A hire added with Status left blank.** Invisible to `onEdit`, which only
  reacts to the Status column. The daily scan flags rows that have a name and
  address but no status after 24 hours.

---

## How silent failure is caught

Two mechanisms, deliberately covering different failure classes.

### 1. The daily alert — the script is running, but rows are stuck

`dailyHealthCheck` scans the roster and emails `ADMIN_EMAIL` (and Slack, if
configured) a list of every row that is:

- in `Error` status, or
- in `New` status for more than 24 hours, or
- name + email filled with Status blank for more than 24 hours.

The 24-hour clock needs a reference time the sheet does not have. Rather than
add a "row created" column for editors to maintain, the scan records when it
first saw each pending row in a Script Property, keyed by email address so that
inserting or deleting rows above does not reset the clock. First sighting starts
the clock rather than alerting immediately, so a row entered minutes before the
scan is not a false alarm.

The check is itself wrapped in a try/catch that sends a `DAILY CHECK FAILED`
email — a watchdog that dies quietly is worse than no watchdog.

### 2. The heartbeat — the script is not running at all

The alert above has a hard limit: **it can only fire while the script still
runs.** It cannot report its own non-execution. That is exactly the permissions
failure described above, and it is invisible from inside the sheet.

So every successful daily check appends a row to **System Health** with the
time, rows scanned, stuck count, whether an alert fired, remaining quota and run
duration. The contents are secondary. What matters is that the tab keeps
growing.

**If the newest timestamp in System Health is more than about a day old, the
automation is dead** — regardless of how healthy the roster looks. A failure
turns into a visible absence rather than silence.

Only successful runs write a heartbeat; a run that threw sends the failure email
and deliberately leaves a gap, so the two signals never mask each other.

That staleness comparison is intentionally **not** done by this script. A script
cannot be trusted to detect that it is not running. Check it via:

- **Onboarding → Show system status**, which flags a heartbeat over 36 hours old
  as `STALE`, or
- an external monitor — a second Apps Script in a different account, or a
  calendar reminder for a weekly glance at the tab.

A recurring reminder for one person to open that tab weekly is low-tech and
genuinely sufficient. Pick something and write down who owns it; an unowned
monitoring tab is the same silent failure one level up.

---

## Operating it

**Retry a failed row:** fix the underlying problem, then set Status back to
`New`. Errored rows have a blank send timestamp, so the already-sent guard does
not block the retry.

**Stop the automation:** **Onboarding → Remove triggers**.

**Never** hand-edit a Status to `Emailed` to suppress an alert without checking
whether mail actually went out — that is the one action that makes the roster
lie.

---

## Five-minute review map

`Code.gs` is ordered so it reads top to bottom:

1. **Config and constants** — tunables, status values, column names.
2. **Sheet and schema helpers** — header resolution, row reads, validation.
3. **onEdit path** — the send, the lock, the idempotency gate, the template.
4. **Error logging** — `Error` status and the Error Log tab.
5. **Daily scan** — stuck-row detection and the admin alert.
6. **Heartbeat** — the System Health tab.
7. **Trigger installation and menu** — setup and diagnostics.

The parts most worth your attention are `sendWelcomeForRow_` (section 3), where
concurrency and idempotency are decided, and `findStuckRows_` (section 5), where
"stuck" is defined.
