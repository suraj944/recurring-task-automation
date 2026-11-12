/**
 * Allo Innoware — Recurring Task Tracking, Pending Reports & Escalation
 *
 * One script, two front doors onto the same engine:
 *
 *   • THE SHEET — a "Task Automation" menu, typing DONE in Employee Pending
 *     View, ticking Acknowledged on the Manager Dashboard, and a daily
 *     time-driven run that emails without anyone pressing anything.
 *   • THE WEBSITE — the same actions over a JSON API (doGet / doPost).
 *
 * Both paths call the same functions, so the sheet and the site can never
 * drift apart or apply different rules.
 *
 * ── SETUP (run ONCE, by the sheet owner) ─────────────────────────────────
 * 1. Sheet > Extensions > Apps Script, paste this whole file in, Save.
 * 2. Run "setupTriggers" once from the toolbar and accept the prompt.
 *      This installs the daily email check and the owner-authority edit
 *      triggers. Without it the sheet still works for YOU, but an
 *      employee's edit cannot write to tabs locked to the owner.
 * 3. Deploy > New deployment > Web app, Execute as: Me,
 *    Who has access: Anyone. Copy the /exec URL for the website.
 *
 * Why step 2 cannot be skipped: Apps Script's simple onEdit runs with the
 * permissions of whoever is typing. Task Master is locked to the owner, so
 * an employee's edit would silently fail. An installable trigger, authorised
 * once by the owner, always runs with the OWNER's permissions — which is
 * what lets an employee's entry update tabs they cannot open themselves.
 *
 * After editing this file, redeploy (Deploy > Manage deployments > edit >
 * Version: New version) or the web app keeps serving the old code.
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

const EMPLOYEE_LIST = ['Employee A','Employee B','Employee C','Employee D','Employee E','Employee F'];

/* ========================================================================
   SETUP — run once, by the owner
   ===================================================================== */

function setupTriggers() {
  // Clear existing triggers first so re-running is always safe and never
  // stacks duplicates.
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });

  const ss = SpreadsheetApp.getActive();
  const sla = getSlaSettings();

  ScriptApp.newTrigger('installableOnEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('installableOnOpen').forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger('dailyCheckRun').timeBased().everyDays(1).atHour(sla.checkRunHour).create();

  PropertiesService.getScriptProperties().setProperty('AUTOMATION_LIVE', 'true');

  const summary =
    'Automation is live, running with YOUR authorisation:\n\n' +
    '• Employees type into Employee Pending View and the script updates ' +
    'Task Master on their behalf, even though that tab is locked to them.\n' +
    '• Manager acknowledgments on the Dashboard sync the same way.\n' +
    '• The daily check runs by itself at ' + sla.checkRunHour + ':00 ' +
    '(set in SLA Settings B7), sending manager alerts and leadership ' +
    'escalations with nobody pressing anything.\n\n' +
    'Re-run this after changing the check-run time, or if automation ever ' +
    'seems to stop.';

  // getUi() only exists when this runs from the spreadsheet. Run from the
  // Apps Script editor it throws, which would look like the setup failed
  // even though every trigger above was created. Log instead.
  try {
    SpreadsheetApp.getUi().alert(summary);
  } catch (err) {
    Logger.log(summary);
  }

  return { ok: true, triggers: ScriptApp.getProjectTriggers().length };
}

/* ========================================================================
   TRIGGERS
   ===================================================================== */

// Simple triggers: instant, no authorisation, but they run as whoever is
// editing. They let the OWNER test before setupTriggers() has been run, then
// step aside once the installable ones take over.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Task Automation')
    .addItem('Refresh dashboard now', 'refreshViews')
    .addItem('Run daily check now (sends real email)', 'dailyCheckRun')
    .addItem('Turn on full automation (run once)', 'setupTriggers')
    .addToUi();

  if (automationLive()) return;
  try { refreshViews(); } catch (err) { /* not the owner yet — ignore */ }
}

function onEdit(e) {
  if (automationLive()) return;
  try { handleAnyEdit(e); } catch (err) { /* not the owner yet — ignore */ }
}

// Installable triggers: created by setupTriggers(), always run with the
// OWNER's authority no matter who made the edit.
function installableOnEdit(e) { handleAnyEdit(e); }
function installableOnOpen(e) { try { refreshViews(); } catch (err) {} }

function automationLive() {
  return PropertiesService.getScriptProperties().getProperty('AUTOMATION_LIVE') === 'true';
}

function handleAnyEdit(e) {
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Employee typed into the update column of their pending list
  if (sheetName === SHEET_EMPLOYEE_VIEW && row >= 5 && col === 8) {
    handleSheetCompletion(row);
    return;
  }

  // Manager ticked Acknowledged on the dashboard
  if (sheetName === SHEET_MGR_DASHBOARD && row >= 3 && col === 8) {
    handleSheetAcknowledgment(row);
  }
}

/* ========================================================================
   WEB APP ENTRY POINTS
   ===================================================================== */

function doGet(e) {
  try {
    return json({ ok: true, tasks: readTasks(), log: readLog(), employees: EMPLOYEE_LIST });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    // Sent as text/plain so the browser treats it as a simple request and
    // skips the CORS preflight, which Apps Script cannot answer.
    const body = JSON.parse(e.postData.contents);
    let result;
    switch (body.action) {
      case 'complete':    result = apiComplete(body);    break;
      case 'add':         result = apiAdd(body);         break;
      case 'acknowledge': result = apiAcknowledge(body); break;
      case 'dailycheck':  result = dailyCheckRun();      break;
      default: return json({ ok: false, error: 'Unknown action: ' + body.action });
    }
    result.tasks = readTasks();
    result.log = readLog();
    return json(result);
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ========================================================================
   SHARED HELPERS
   ===================================================================== */

function sheetOf(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function iso(d) {
  if (!d) return null;
  if (!(d instanceof Date)) return String(d);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function readTasks() {
  const sheet = sheetOf(SHEET_TASK_MASTER);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const out = [];
  values.forEach(function (r, i) {
    if (!r[TM.TASK_ID] || !r[TM.NAME]) return;
    out.push({
      row: i + 2,
      id: String(r[TM.TASK_ID]),
      name: String(r[TM.NAME]),
      employee: String(r[TM.ASSIGNED_TO]),
      freq: Number(r[TM.FREQ_DAYS]) || 1,
      ticket: String(r[TM.REQUIRES_TICKET]).trim().toLowerCase() === 'yes',
      cycleStart: iso(r[TM.CYCLE_START]),
      lastCompleted: iso(r[TM.LAST_COMPLETED]),
      completedBy: r[TM.COMPLETED_BY] ? String(r[TM.COMPLETED_BY]) : null,
      ticketProof: r[TM.TICKET_PROOF] ? String(r[TM.TICKET_PROOF]) : null,
      due: iso(r[TM.NEXT_DUE]),
      status: String(r[TM.STATUS]),
      daysOverdue: Number(r[TM.DAYS_OVERDUE]) || 0,
      alertSent: iso(r[TM.ALERT_SENT]),
      mgrAck: r[TM.MGR_ACK] === true,
      ackDate: iso(r[TM.ACK_DATE]),
      escalated: iso(r[TM.ESCALATED])
    });
  });
  return out;
}

function readLog() {
  const sheet = sheetOf(SHEET_LOG);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const start = Math.max(2, lastRow - 49);
  const values = sheet.getRange(start, 1, lastRow - start + 1, 6).getValues();
  return values.map(function (r) {
    return {
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      taskId: String(r[1]), taskName: String(r[2]),
      employee: String(r[3]), type: String(r[4]), detail: String(r[5])
    };
  }).reverse();
}

function getSlaSettings() {
  const sla = sheetOf(SHEET_SLA);
  if (!sla) return { managerEmail:'', leadershipEmail:'', checkRunHour:18, ackWindowHours:24, maxWindowHours:72 };
  const timeStr = sla.getRange('B7').getDisplayValue();     // e.g. "18:00"
  const hour = parseInt(String(timeStr).split(':')[0], 10);
  return {
    managerEmail:    sla.getRange('B5').getValue(),
    leadershipEmail: sla.getRange('B6').getValue(),
    checkRunHour:    isNaN(hour) ? 18 : hour,
    ackWindowHours:  Number(sla.getRange('B8').getValue()) || 24,
    maxWindowHours:  Number(sla.getRange('B9').getValue()) || 72
  };
}

function logEvent(taskId, taskName, employee, eventType, details) {
  sheetOf(SHEET_LOG).appendRow([new Date(), taskId, taskName, employee, eventType, details]);
}

function hoursBetween(a, b) { return (b.getTime() - a.getTime()) / 36e5; }

function findTask(taskId) {
  const tasks = readTasks();
  for (let i = 0; i < tasks.length; i++) if (tasks[i].id === taskId) return tasks[i];
  return null;
}

/* ========================================================================
   COMPLETION — one engine, called from the sheet and from the website
   ===================================================================== */

/**
 * The single place a completion is recorded. Everything else is a caller.
 * Returns {ok, message} rather than touching any UI, so the sheet path can
 * raise a toast and the web path can return JSON from the same result.
 */
function recordCompletion(taskId, rawEntry, employeeName) {
  const task = findTask(taskId);
  if (!task) return { ok: false, error: 'Task not found: ' + taskId };

  let entry = String(rawEntry || '').trim();
  if (!entry) return { ok: false, error: 'Nothing entered.' };

  // "done" in any case becomes DONE
  if (entry.toUpperCase() === 'DONE') entry = 'DONE';

  if (task.ticket && entry === 'DONE') {
    return { ok: false, error: 'This task requires a Ticket ID, not DONE.' };
  }

  const sheet = sheetOf(SHEET_TASK_MASTER);
  const now = new Date();
  const who = employeeName || task.employee;

  sheet.getRange(task.row, TM.LAST_COMPLETED + 1).setValue(now);
  sheet.getRange(task.row, TM.COMPLETED_BY + 1).setValue(who);
  sheet.getRange(task.row, TM.TICKET_PROOF + 1).setValue(task.ticket ? entry : 'Daily update confirmed');
  // Clear the previous cycle's flags so the new cycle starts clean
  sheet.getRange(task.row, TM.ALERT_SENT + 1).setValue('');
  sheet.getRange(task.row, TM.MGR_ACK + 1).setValue(false);
  sheet.getRange(task.row, TM.ACK_DATE + 1).setValue('');
  sheet.getRange(task.row, TM.ESCALATED + 1).setValue('');

  logEvent(task.id, task.name, who, 'Task Completed',
    task.ticket ? 'Ticket ' + entry + ' logged by ' + who
                : 'Daily update confirmed by ' + who);

  SpreadsheetApp.flush();
  return { ok: true, message: task.id + ' marked complete. Next cycle calculated automatically.',
           entry: entry, taskName: task.name };
}

// --- called from the SHEET (employee typed in column H) ---
function handleSheetCompletion(row) {
  const view = sheetOf(SHEET_EMPLOYEE_VIEW);
  const entryCell = view.getRange(row, 8);
  const raw = entryCell.getValue();
  if (!raw || String(raw).trim() === '') return;      // empty edit does nothing

  const taskId = view.getRange(row, 1).getValue();
  if (!taskId) { entryCell.clearContent(); return; }  // typed beside no task

  // Show the corrected casing immediately in the cell the employee typed in.
  if (String(raw).trim().toUpperCase() === 'DONE' && raw !== 'DONE') entryCell.setValue('DONE');

  const employee = view.getRange('B3').getValue();
  const res = recordCompletion(taskId, raw, employee);

  if (!res.ok) {
    SpreadsheetApp.getActive().toast(res.error, '⚠ Update not saved', 8);
    return;
  }

  // The pending list is a live formula that re-sorts the moment Task Master
  // changes, so anything left in the input column would end up beside a
  // different task. Clear it.
  view.getRange(5, 8, Math.max(view.getLastRow() - 4, 1), 1).clearContent();

  SpreadsheetApp.getActive().toast(res.message, '✅ Saved', 5);
  refreshManagerDashboard();
}

// --- called from the WEBSITE ---
function apiComplete(body) {
  const res = recordCompletion(body.taskId, body.entry, body.employee);
  if (res.ok) refreshManagerDashboard();
  return res;
}

/* ========================================================================
   ACKNOWLEDGMENT
   ===================================================================== */

function recordAcknowledgment(taskId) {
  const task = findTask(taskId);
  if (!task) return { ok: false, error: 'Task not found: ' + taskId };
  if (task.mgrAck) return { ok: true, message: task.id + ' was already acknowledged.' };

  const sheet = sheetOf(SHEET_TASK_MASTER);
  const now = new Date();
  sheet.getRange(task.row, TM.MGR_ACK + 1).setValue(true);
  sheet.getRange(task.row, TM.ACK_DATE + 1).setValue(now);

  logEvent(task.id, task.name, task.employee, 'Manager Acknowledged',
    'Manager acknowledged. The unresolved ceiling still applies.');

  SpreadsheetApp.flush();
  return { ok: true, message: task.id + ' acknowledged.' };
}

// --- called from the SHEET (manager ticked the checkbox) ---
function handleSheetAcknowledgment(row) {
  const dash = sheetOf(SHEET_MGR_DASHBOARD);
  const cell = dash.getRange(row, 8);
  const taskId = dash.getRange(row, 1).getValue();
  if (!taskId) return;

  // Acknowledging is a one-way record of a decision: once it is on the
  // audit trail it cannot be taken back. Un-ticking puts the box straight
  // back, so the sheet can never disagree with the log.
  if (cell.getValue() !== true) {
    const existing = findTask(taskId);
    if (existing && existing.mgrAck) {
      cell.setValue(true);
      SpreadsheetApp.getActive().toast(
        taskId + ' was acknowledged on ' + (existing.ackDate || 'an earlier date') +
        '. An acknowledgment cannot be undone.', '🔒 Locked', 6);
    }
    return;
  }

  const res = recordAcknowledgment(taskId);
  if (res.ok) {
    refreshManagerDashboard();
    SpreadsheetApp.getActive().toast(res.message, '✅ Acknowledged', 4);
  }
}

// --- called from the WEBSITE ---
function apiAcknowledge(body) {
  const res = recordAcknowledgment(body.taskId);
  if (res.ok) refreshManagerDashboard();
  return res;
}

/* ========================================================================
   ADDING A TASK
   ===================================================================== */

function apiAdd(body) {
  const name = String(body.name || '').trim();
  if (!name) return { ok: false, error: 'Task name is required.' };
  const freq = Number(body.freq) || 1;
  if (freq < 1) return { ok: false, error: 'Frequency must be at least 1 day.' };
  const employee = String(body.employee || '').trim();
  if (!employee) return { ok: false, error: 'Assigned employee is required.' };
  const startStr = String(body.cycleStart || '').trim();
  if (!startStr) return { ok: false, error: 'Cycle start date is required.' };

  // Auto-generate the next unused Task ID so a duplicate is impossible —
  // a duplicate would make every lookup here ambiguous.
  const tasks = readTasks();
  let max = 0;
  tasks.forEach(function (t) {
    const n = parseInt(String(t.id).replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  const id = 'TASK-' + ('00' + (max + 1)).slice(-3);

  const sheet = sheetOf(SHEET_TASK_MASTER);
  const row = sheet.getLastRow() + 1;
  const parts = startStr.split('-');
  const start = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));

  sheet.getRange(row, 1, 1, 6).setValues([[id, name, employee, freq, body.ticket ? 'Yes' : 'No', start]]);

  // The three calculated columns, matching the rows above.
  sheet.getRange(row, 10).setFormula(
    '=IF($B' + row + '="","",IF($G' + row + '="",$F' + row + ',$G' + row + '+$D' + row + '))');
  sheet.getRange(row, 11).setFormula(
    '=IF($B' + row + '="","",IF($J' + row + '="","",IF(TODAY()>$J' + row + ',"Overdue",' +
    'IF(TODAY()=$J' + row + ',"Due Today","Upcoming"))))');
  sheet.getRange(row, 12).setFormula(
    '=IF($B' + row + '="",0,IF($J' + row + '="",0,MAX(0,TODAY()-$J' + row + ')))');
  sheet.getRange(row, 14).insertCheckboxes().setValue(false);

  logEvent(id, name, employee, 'Task Created',
    'Created — ' + (freq === 1 ? 'daily' : 'every ' + freq + ' days') + ', ' +
    (body.ticket ? 'ticket ID required' : 'DONE confirmation only') + '.');

  SpreadsheetApp.flush();
  refreshManagerDashboard();
  return { ok: true, message: id + ' — ' + name + ', assigned to ' + employee + '.' };
}

/* ========================================================================
   DAILY CHECK — the time-driven trigger, the menu item, and the website
   button all land here
   ===================================================================== */

function dailyCheckRun() {
  const sla = getSlaSettings();
  const sheet = sheetOf(SHEET_TASK_MASTER);
  const tasks = readTasks();
  const now = new Date();
  let alerts = 0, escalations = 0;

  tasks.forEach(function (t) {
    if (t.status !== 'Overdue' && t.status !== 'Due Today') return;

    // --- same-day manager alert, once per cycle ---
    if (!t.alertSent) {
      if (sla.managerEmail) {
        MailApp.sendEmail({
          to: sla.managerEmail,
          subject: '[Allo Innoware] Pending task: ' + t.name + ' (' + t.employee + ')',
          body: t.id + ' — ' + t.name + ', assigned to ' + t.employee + ', was due on ' +
                t.due + ' and is not yet marked complete.\n\n' +
                SpreadsheetApp.getActive().getUrl()
        });
      }
      sheet.getRange(t.row, TM.ALERT_SENT + 1).setValue(now);
      logEvent(t.id, t.name, t.employee, 'Manager Alert Sent',
        'Automated email sent to ' + sla.managerEmail + ' the same day the task became due/overdue.');
      alerts++;
      return;   // don't also evaluate escalation on the first run
    }

    // --- escalation, fires once ---
    if (t.escalated) return;

    const hoursSinceAlert = hoursBetween(new Date(t.alertSent), now);
    const hoursSinceDue   = t.due ? hoursBetween(new Date(t.due), now) : 0;
    const noAckInTime     = !t.mgrAck && hoursSinceAlert >= sla.ackWindowHours;
    const unresolvedTooLong = hoursSinceDue >= sla.maxWindowHours;

    if (noAckInTime || unresolvedTooLong) {
      const reason = noAckInTime
        ? 'No manager acknowledgment within the ' + sla.ackWindowHours + '-hour window.'
        : 'Task still unresolved ' + Math.floor(hoursSinceDue) + ' hours after its due date ' +
          '(ceiling: ' + sla.maxWindowHours + 'h), regardless of acknowledgment.';
      if (sla.leadershipEmail) {
        MailApp.sendEmail({
          to: sla.leadershipEmail,
          subject: '[Allo Innoware] ESCALATION: ' + t.name + ' (' + t.employee + ') unresolved',
          body: t.id + ' — ' + t.name + ' remains unresolved.\n\nReason: ' + reason + '\n\n' +
                SpreadsheetApp.getActive().getUrl()
        });
      }
      sheet.getRange(t.row, TM.ESCALATED + 1).setValue(now);
      logEvent(t.id, t.name, t.employee, 'Leadership Escalation',
        'Escalated to ' + sla.leadershipEmail + '. ' + reason);
      escalations++;
    }
  });

  SpreadsheetApp.flush();
  refreshManagerDashboard();
  return {
    ok: true,
    message: alerts + ' manager alert(s), ' + escalations + ' escalation(s) sent.',
    alerts: alerts, escalations: escalations
  };
}

/* ========================================================================
   VIEW SYNC
   ===================================================================== */

function refreshViews() { refreshManagerDashboard(); }

/**
 * Employee Pending View needs nothing: columns A-G are a live FILTER over
 * Task Master, so they update themselves.
 *
 * Manager Dashboard columns A-G are the same, but H-J hold a real checkbox
 * and two dates that a formula cannot own. This re-syncs those three by
 * Task ID, so they stay beside the right row as the list re-sorts.
 */
function refreshManagerDashboard() {
  const dash = sheetOf(SHEET_MGR_DASHBOARD);
  if (!dash) return;
  const lastRow = dash.getLastRow();
  if (lastRow < 3) return;

  const ids = dash.getRange(3, 1, lastRow - 2, 1).getValues();
  const tasks = readTasks();
  const byId = {};
  tasks.forEach(function (t) { byId[t.id] = t; });

  const out = ids.map(function (r) {
    const t = byId[String(r[0])];
    if (!t) return ['', '', ''];
    return [t.mgrAck === true, t.ackDate ? new Date(t.ackDate) : '', t.escalated ? new Date(t.escalated) : ''];
  });

  dash.getRange(3, 8, out.length, 3).setValues(out);
}
