/**
 * Allo Innoware — Recurring Task Tracking, Pending Reports & Escalation Automation
 *
 * SETUP (run ONCE, by the sheet owner): Paste this entire file into
 * Extensions > Apps Script (Code.gs), save, then from the Apps Script
 * toolbar run the function "setupTriggers" once and accept the
 * authorization prompt.
 *
 * Why this one manual step is unavoidable, and what it buys you:
 * Google will never let a script act with someone else's authority
 * unless a human explicitly grants that once. Running "setupTriggers"
 * one time, as the OWNER, makes every automated action from then on run
 * with YOUR permissions — which is exactly what lets an employee type
 * into Employee Pending View (which they CAN access) and have the
 * script update Task Master, Manager Dashboard, etc. (which they CANNOT
 * access directly — those tabs are locked to the owner only). Without
 * this step, only your own edits as the owner would flow through
 * automatically; other people's edits would silently do nothing on the
 * locked tabs, because Apps Script's "simple trigger" (onEdit/onOpen)
 * always runs with the permissions of whoever is actually typing, not
 * the script owner's permissions.
 *
 * This single setupTriggers() run also turns on the daily email check
 * (manager alerts + leadership escalation) at the hour set below.
 */

const SHEET_TASK_MASTER   = 'Task Master';
const SHEET_SLA           = 'SLA Settings';
const SHEET_EMPLOYEE_VIEW = 'Employee Pending View';
const SHEET_MGR_DASHBOARD = 'Manager Dashboard';
const SHEET_LOG           = 'Escalation Log';

// Task Master column indexes (0-based)
const TM = {
  TASK_ID: 0, NAME: 1, ASSIGNED_TO: 2, FREQ_DAYS: 3, REQUIRES_TICKET: 4,
  CYCLE_START: 5, LAST_COMPLETED: 6, COMPLETED_BY: 7, TICKET_PROOF: 8,
  NEXT_DUE: 9, STATUS: 10, DAYS_OVERDUE: 11, ALERT_SENT: 12,
  MGR_ACK: 13, ACK_DATE: 14, ESCALATED: 15
};

// ---------- SETUP (run ONCE, by the owner) ----------
function setupTriggers() {
  // Remove any previously installed triggers first, so re-running this is
  // always safe and never creates duplicates.
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  const ss = SpreadsheetApp.getActive();
  const sla = getSlaSettings();

  ScriptApp.newTrigger('installableOnEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('installableOnOpen').forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger('dailyCheckRun').timeBased().everyDays(1).atHour(sla.checkRunHour).create();

  // Flag so the reserved onEdit/onOpen (below) know the installable
  // triggers are now handling everything, and step aside to avoid
  // processing the exact same edit twice.
  PropertiesService.getScriptProperties().setProperty('AUTOMATION_LIVE', 'true');

  SpreadsheetApp.getUi().alert(
    'Full automation is now live, running with YOUR authorization:\n\n' +
    '• Employees can type into Employee Pending View, and the script will ' +
    'now correctly update Task Master and Manager Dashboard on their ' +
    'behalf, even though those tabs are locked to them directly.\n' +
    '• Manager acknowledgments on the Dashboard now sync the same way.\n' +
    '• Daily email alerts & leadership escalations are turned on, at the ' +
    'hour set in SLA Settings (B7).\n\n' +
    'Re-run this once any time you re-share the sheet with new employees, ' +
    'change the daily check-run time, or if automation seems to stop ' +
    'working after editing the script.'
  );
}

// ---------- RESERVED SIMPLE TRIGGERS ----------
// These run instantly for everyone with zero authorization, but — critically —
// they run with the permissions of whoever is actually editing, not the
// owner's. Before setupTriggers() has been run, this lets the OWNER test
// everything immediately. Once setupTriggers() has been run, these simply
// step aside (see the AUTOMATION_LIVE check) and let the installable
// triggers below do the real work with full owner authority instead.

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Task Automation')
    .addItem('Refresh views now', 'refreshViews')
    .addItem('Run daily check now (test — sends real email)', 'dailyCheckRun')
    .addItem('Turn on full automation (run once)', 'setupTriggers')
    .addToUi();

  if (PropertiesService.getScriptProperties().getProperty('AUTOMATION_LIVE') === 'true') return;
  try { refreshViews(); } catch (err) { /* non-owner, automation not set up yet — ignore */ }
}

function onEdit(e) {
  if (PropertiesService.getScriptProperties().getProperty('AUTOMATION_LIVE') === 'true') return;
  try { handleAnyEdit(e); } catch (err) { /* non-owner, automation not set up yet — ignore */ }
}

// ---------- INSTALLABLE TRIGGERS ----------
// Created once by setupTriggers(). These always run with the OWNER's full
// authority, regardless of who actually made the edit — this is what lets
// the script write into tabs the acting employee cannot open themselves.

function installableOnEdit(e) {
  handleAnyEdit(e);
}

function installableOnOpen(e) {
  try { refreshViews(); } catch (err) { /* ignore */ }
}

// ---------- SHARED EDIT ROUTER ----------
function handleAnyEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheetName === SHEET_EMPLOYEE_VIEW && row === 3 && col === 2) {
    refreshEmployeeView();   // name dropdown changed — reload that employee's pending list
    return;
  }

  if (sheetName === SHEET_EMPLOYEE_VIEW && row >= 5 && col === 8) {
    handleEmployeeCompletion(sheet, row);
  }

  if (sheetName === SHEET_MGR_DASHBOARD && row >= 3 && col === 8) {
    handleManagerAcknowledgment(sheet, row);
  }
}

// ---------- SHARED HELPERS ----------
function getTaskMasterData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_TASK_MASTER);
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  return { sheet, values };
}

function findRowByTaskId(values, taskId) {
  for (let i = 0; i < values.length; i++) {
    if (values[i][TM.TASK_ID] === taskId) return i; // 0-based index into values
  }
  return -1;
}

function logEvent(taskId, taskName, employee, eventType, details) {
  const log = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  log.appendRow([new Date(), taskId, taskName, employee, eventType, details]);
}

function getSlaSettings() {
  const sla = SpreadsheetApp.getActive().getSheetByName(SHEET_SLA);
  const timeStr = sla.getRange('B7').getDisplayValue(); // e.g. "18:00"
  const hour = parseInt(String(timeStr).split(':')[0], 10);
  return {
    managerEmail:   sla.getRange('B5').getValue(),
    leadershipEmail: sla.getRange('B6').getValue(),
    checkRunHour: isNaN(hour) ? 18 : hour,
    ackWindowHours: Number(sla.getRange('B8').getValue()) || 24,
    maxWindowHours: Number(sla.getRange('B9').getValue()) || 72
  };
}

function hoursBetween(a, b) {
  return (b.getTime() - a.getTime()) / (1000 * 60 * 60);
}

// ---------- SINGLE-POINT COMPLETION (Tests 1, 2, 3, 8) ----------
function handleEmployeeCompletion(viewSheet, row) {
  const entryCell = viewSheet.getRange(row, 8);
  let entry = entryCell.getValue();
  if (!entry || String(entry).trim() === '') return; // Test 8: empty edit does nothing

  // Auto-correct case: "done", "Done", "dOnE", etc. all become "DONE" in the
  // cell itself, so the sheet always shows a consistent uppercase keyword
  // no matter how the employee typed it.
  if (String(entry).trim().toUpperCase() === 'DONE' && entry !== 'DONE') {
    entry = 'DONE';
    entryCell.setValue('DONE');
  }

  const taskId = viewSheet.getRange(row, 1).getValue();
  if (!taskId) return; // row has no task mapped — nothing to do (prevents stray writes)

  const { sheet: tmSheet, values } = getTaskMasterData();
  const idx = findRowByTaskId(values, taskId);
  if (idx === -1) return;

  const tmRow = idx + 2; // +2: header row + 1-based
  const requiresTicket = values[idx][TM.REQUIRES_TICKET] === 'Yes';
  const employeeName = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEE_VIEW).getRange('B3').getValue();
  const now = new Date();

  // Validation: a portal task needs something that isn't just "DONE".
  // Uses a toast (not a cell overwrite) so the employee's typed entry is
  // never silently erased by the next refresh — they can see exactly what
  // they typed and fix it, rather than wondering where it went.
  if (requiresTicket && String(entry).trim().toUpperCase() === 'DONE') {
    SpreadsheetApp.getActive().toast(
      'This task requires a Ticket ID, not DONE. Enter the actual Ticket ID instead.',
      '⚠ Update not saved', 8);
    return;
  }

  // Write completion back to Task Master — the ONE place state changes.
  // This write happens under the OWNER's authority (installable trigger),
  // so it succeeds even though the employee themselves cannot open Task
  // Master directly.
  tmSheet.getRange(tmRow, TM.LAST_COMPLETED + 1).setValue(now);
  tmSheet.getRange(tmRow, TM.COMPLETED_BY + 1).setValue(employeeName);
  tmSheet.getRange(tmRow, TM.TICKET_PROOF + 1).setValue(requiresTicket ? entry : 'Daily update confirmed');
  // Clear prior-cycle flags so the new cycle starts clean (Test 1: cycle closes and resets)
  tmSheet.getRange(tmRow, TM.ALERT_SENT + 1).setValue('');
  tmSheet.getRange(tmRow, TM.MGR_ACK + 1).setValue(false);
  tmSheet.getRange(tmRow, TM.ACK_DATE + 1).setValue('');
  tmSheet.getRange(tmRow, TM.ESCALATED + 1).setValue('');

  logEvent(taskId, values[idx][TM.NAME], employeeName, 'Task Completed',
    requiresTicket ? `Ticket ${entry} logged by ${employeeName}` : `Daily update confirmed by ${employeeName}`);

  SpreadsheetApp.getActive().toast(`${taskId} marked complete. Next cycle calculated automatically.`, '✅ Saved', 5);

  // Refresh so the completed row disappears from the pending list immediately
  refreshEmployeeView();
  refreshManagerDashboard();
}

function handleManagerAcknowledgment(dashSheet, row) {
  const checked = dashSheet.getRange(row, 8).getValue();
  if (checked !== true) return;

  const taskId = dashSheet.getRange(row, 1).getValue();
  if (!taskId) return;

  const { sheet: tmSheet, values } = getTaskMasterData();
  const idx = findRowByTaskId(values, taskId);
  if (idx === -1) return;
  if (values[idx][TM.MGR_ACK] === true) return; // Test 8: already acknowledged, do nothing

  const tmRow = idx + 2;
  const now = new Date();
  tmSheet.getRange(tmRow, TM.MGR_ACK + 1).setValue(true);
  tmSheet.getRange(tmRow, TM.ACK_DATE + 1).setValue(now);
  dashSheet.getRange(row, 9).setValue(now);

  logEvent(taskId, values[idx][TM.NAME], values[idx][TM.ASSIGNED_TO], 'Manager Acknowledged',
    'Manager ticked Acknowledged on the dashboard. 72-hour unresolved ceiling still applies.');
}

// ---------- TEST 4, 5, 6, 7: DAILY ALERT + ESCALATION ----------
function dailyCheckRun() {
  const { sheet: tmSheet, values } = getTaskMasterData();
  const sla = getSlaSettings();
  const now = new Date();

  values.forEach((rowVals, i) => {
    if (!rowVals[TM.TASK_ID]) return;
    const status = rowVals[TM.STATUS];
    if (status !== 'Overdue' && status !== 'Due Today') return;

    const tmRow = i + 2;
    const taskId = rowVals[TM.TASK_ID];
    const taskName = rowVals[TM.NAME];
    const employee = rowVals[TM.ASSIGNED_TO];
    const alertSent = rowVals[TM.ALERT_SENT];
    const ack = rowVals[TM.MGR_ACK];
    const escalated = rowVals[TM.ESCALATED];
    const dueDate = rowVals[TM.NEXT_DUE];

    // --- Test 4: same-day manager alert, sent once ---
    if (!alertSent) {
      sendManagerAlertEmail(sla.managerEmail, taskId, taskName, employee, dueDate);
      tmSheet.getRange(tmRow, TM.ALERT_SENT + 1).setValue(now);
      logEvent(taskId, taskName, employee, 'Manager Alert Sent',
        `Automated email sent to ${sla.managerEmail} same day task became due/overdue.`);
      return; // don't also evaluate escalation on the very first run
    }

    // --- Test 6 / 7: escalation logic, fires once ---
    if (escalated) return; // already escalated — never resend (Test 8)

    const hoursSinceAlert = hoursBetween(new Date(alertSent), now);
    const hoursSinceDue = dueDate ? hoursBetween(new Date(dueDate), now) : 0;

    const noAckInTime = !ack && hoursSinceAlert >= sla.ackWindowHours;
    const unresolvedTooLong = hoursSinceDue >= sla.maxWindowHours;

    if (noAckInTime || unresolvedTooLong) {
      const reason = noAckInTime
        ? `No manager acknowledgment within ${sla.ackWindowHours}-hour window.`
        : `Task still unresolved ${Math.floor(hoursSinceDue)} hours after due date (ceiling: ${sla.maxWindowHours}h), regardless of acknowledgment.`;
      sendLeadershipEscalationEmail(sla.leadershipEmail, taskId, taskName, employee, dueDate, reason);
      tmSheet.getRange(tmRow, TM.ESCALATED + 1).setValue(now);
      logEvent(taskId, taskName, employee, 'Leadership Escalation',
        `Escalated to ${sla.leadershipEmail}. ${reason}`);
    }
  });

  refreshEmployeeView();
  refreshManagerDashboard();
}

function sendManagerAlertEmail(to, taskId, taskName, employee, dueDate) {
  if (!to) return;
  MailApp.sendEmail({
    to: to,
    subject: `[Allo Innoware] Pending task: ${taskName} (${employee})`,
    body: `${taskId} — ${taskName}, assigned to ${employee}, was due on ${dueDate} and is not yet marked complete.\n\n` +
          `Open the Manager Dashboard tab to review and acknowledge:\n${SpreadsheetApp.getActive().getUrl()}`
  });
}

function sendLeadershipEscalationEmail(to, taskId, taskName, employee, dueDate, reason) {
  if (!to) return;
  MailApp.sendEmail({
    to: to,
    subject: `[Allo Innoware] ESCALATION: ${taskName} (${employee}) unresolved`,
    body: `${taskId} — ${taskName}, assigned to ${employee}, due ${dueDate}, remains unresolved.\n\n` +
          `Reason for escalation: ${reason}\n\n` +
          `Full detail: ${SpreadsheetApp.getActive().getUrl()}`
  });
}

// ---------- VIEW REFRESH (keeps Employee View and Manager Dashboard current) ----------
function refreshViews() {
  refreshEmployeeView();
  refreshManagerDashboard();
}

function refreshEmployeeView() {
  const ss = SpreadsheetApp.getActive();
  const viewSheet = ss.getSheetByName(SHEET_EMPLOYEE_VIEW);
  const employee = viewSheet.getRange('B3').getValue();
  if (!employee) return;

  const { values } = getTaskMasterData();
  const rows = values
    .filter(r => r[TM.TASK_ID] && r[TM.ASSIGNED_TO] === employee &&
                 (r[TM.STATUS] === 'Overdue' || r[TM.STATUS] === 'Due Today' || r[TM.STATUS] === 'Upcoming'))
    .sort((a, b) => new Date(a[TM.NEXT_DUE]) - new Date(b[TM.NEXT_DUE]))
    .map(r => [
      r[TM.TASK_ID], r[TM.NAME], r[TM.FREQ_DAYS] === 1 ? 'Daily' : `Every ${r[TM.FREQ_DAYS]} days`,
      r[TM.NEXT_DUE], r[TM.STATUS], r[TM.DAYS_OVERDUE],
      r[TM.REQUIRES_TICKET] === 'Yes' ? 'Enter Ticket ID' : 'Type DONE when posted to WhatsApp group',
      ''
    ]);

  // Clear old list, then write fresh rows — prevents stale/duplicate rows (Test 8)
  const maxRows = Math.max(viewSheet.getLastRow() - 4, 1);
  viewSheet.getRange(5, 1, maxRows, 8).clearContent();
  if (rows.length > 0) {
    viewSheet.getRange(5, 1, rows.length, 8).setValues(rows);
  }
}

function refreshManagerDashboard() {
  // The dashboard's A3 cell holds a live QUERY formula over Task Master,
  // so columns A–G refresh automatically. This function only re-syncs the
  // editable H–J columns (Acknowledged?, Acknowledged Date, Escalated) so
  // they line up with whatever rows the QUERY is currently showing.
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SHEET_MGR_DASHBOARD);
  const lastRow = dash.getLastRow();
  if (lastRow < 3) return;

  const taskIds = dash.getRange(3, 1, lastRow - 2, 1).getValues();
  const { values } = getTaskMasterData();

  taskIds.forEach((idRow, i) => {
    const taskId = idRow[0];
    const r = 3 + i;
    if (!taskId) {
      // No task on this row anymore (QUERY shrank) — clear any stale
      // checkbox/date left behind from a row that used to be here.
      dash.getRange(r, 8, 1, 3).clearContent();
      return;
    }
    const idx = findRowByTaskId(values, taskId);
    if (idx === -1) return;
    dash.getRange(r, 8).setValue(values[idx][TM.MGR_ACK] === true);
    dash.getRange(r, 9).setValue(values[idx][TM.ACK_DATE] || '');
    dash.getRange(r, 10).setValue(values[idx][TM.ESCALATED] || '');
  });
}
