/**
 * Allo Innoware — Web App API
 *
 * Turns the Google Sheet into a JSON endpoint so the website can use it as
 * a live database: reading tasks, recording completions, adding tasks,
 * acknowledging, and running the alert/escalation check.
 *
 * ── SETUP (once) ──────────────────────────────────────────────────────────
 * 1. Open the SHEET, then Extensions > Apps Script.
 * 2. Paste this whole file in (replacing anything there), and Save.
 * 3. Click "Deploy" > "New deployment".
 * 4. Click the gear next to "Select type" and choose "Web app".
 * 5. Set:
 *       Execute as:      Me (your account)
 *       Who has access:  Anyone
 *    "Anyone" is required — the website calls this without signing in.
 *    The sheet itself stays private; only this endpoint is reachable, and
 *    it only does the four actions below.
 * 6. Click Deploy, authorize when prompted, and COPY the Web app URL.
 *    It looks like: https://script.google.com/macros/s/AKfy..../exec
 * 7. Open the website, click "Connect to sheet", paste that URL.
 *
 * NOTE: after editing this file you must deploy again (Deploy > Manage
 * deployments > edit > Version: New version) or the URL keeps serving the
 * old code.
 */

const SHEET_TASK_MASTER = 'Task Master';
const SHEET_LOG         = 'Escalation Log';
const SHEET_SLA         = 'SLA Settings';

// Task Master column indexes (0-based)
const TM = {
  TASK_ID: 0, NAME: 1, ASSIGNED_TO: 2, FREQ_DAYS: 3, REQUIRES_TICKET: 4,
  CYCLE_START: 5, LAST_COMPLETED: 6, COMPLETED_BY: 7, TICKET_PROOF: 8,
  NEXT_DUE: 9, STATUS: 10, DAYS_OVERDUE: 11, ALERT_SENT: 12,
  MGR_ACK: 13, ACK_DATE: 14, ESCALATED: 15
};

// ---------- entry points ----------

function doGet(e) {
  try {
    return json({ ok: true, tasks: readTasks(), log: readLog(), employees: EMPLOYEES() });
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
      case 'dailycheck':  result = apiDailyCheck();      break;
      default: return json({ ok: false, error: 'Unknown action: ' + body.action });
    }
    // Always hand back fresh state so the page never has to guess.
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

function EMPLOYEES() {
  return ['Employee A','Employee B','Employee C','Employee D','Employee E','Employee F'];
}

// ---------- reading ----------

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
  const start = Math.max(2, lastRow - 49);            // last 50 entries
  const values = sheet.getRange(start, 1, lastRow - start + 1, 6).getValues();
  return values.map(function (r) {
    return {
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      taskId: String(r[1]), taskName: String(r[2]),
      employee: String(r[3]), type: String(r[4]), detail: String(r[5])
    };
  }).reverse();                                        // newest first
}

function getSla() {
  const sla = sheetOf(SHEET_SLA);
  if (!sla) return { managerEmail:'', leadershipEmail:'', ackWindowHours:24, maxWindowHours:72 };
  return {
    managerEmail:    sla.getRange('B5').getValue(),
    leadershipEmail: sla.getRange('B6').getValue(),
    ackWindowHours:  Number(sla.getRange('B8').getValue()) || 24,
    maxWindowHours:  Number(sla.getRange('B9').getValue()) || 72
  };
}

function logEvent(taskId, taskName, employee, eventType, details) {
  sheetOf(SHEET_LOG).appendRow([new Date(), taskId, taskName, employee, eventType, details]);
}

function findRow(taskId) {
  const tasks = readTasks();
  for (let i = 0; i < tasks.length; i++) if (tasks[i].id === taskId) return tasks[i];
  return null;
}

// ---------- actions ----------

function apiComplete(body) {
  const task = findRow(body.taskId);
  if (!task) return { ok: false, error: 'Task not found: ' + body.taskId };

  let entry = String(body.entry || '').trim();
  if (!entry) return { ok: false, error: 'Nothing entered.' };

  // "done" in any case becomes DONE
  if (entry.toUpperCase() === 'DONE') entry = 'DONE';

  if (task.ticket && entry === 'DONE') {
    return { ok: false, error: 'This task requires a Ticket ID, not DONE.' };
  }

  const sheet = sheetOf(SHEET_TASK_MASTER);
  const now = new Date();
  const who = body.employee || task.employee;

  sheet.getRange(task.row, TM.LAST_COMPLETED + 1).setValue(now);
  sheet.getRange(task.row, TM.COMPLETED_BY + 1).setValue(who);
  sheet.getRange(task.row, TM.TICKET_PROOF + 1).setValue(task.ticket ? entry : 'Daily update confirmed');
  // clear the previous cycle's flags so the new cycle starts clean
  sheet.getRange(task.row, TM.ALERT_SENT + 1).setValue('');
  sheet.getRange(task.row, TM.MGR_ACK + 1).setValue(false);
  sheet.getRange(task.row, TM.ACK_DATE + 1).setValue('');
  sheet.getRange(task.row, TM.ESCALATED + 1).setValue('');

  logEvent(task.id, task.name, who, 'Task Completed',
    task.ticket ? 'Ticket ' + entry + ' logged by ' + who
                : 'Daily update confirmed by ' + who);

  SpreadsheetApp.flush();
  return { ok: true, message: task.id + ' marked complete. Next cycle calculated automatically.' };
}

function apiAdd(body) {
  const name = String(body.name || '').trim();
  if (!name) return { ok: false, error: 'Task name is required.' };
  const freq = Number(body.freq) || 1;
  if (freq < 1) return { ok: false, error: 'Frequency must be at least 1 day.' };
  const employee = String(body.employee || '').trim();
  if (!employee) return { ok: false, error: 'Assigned employee is required.' };
  const startStr = String(body.cycleStart || '').trim();
  if (!startStr) return { ok: false, error: 'Cycle start date is required.' };

  // Auto-generate the next unused Task ID, so a duplicate is impossible.
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

  sheet.getRange(row, 1, 1, 6).setValues([[
    id, name, employee, freq, body.ticket ? 'Yes' : 'No', start
  ]]);

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
    'Created from the web app — ' + (freq === 1 ? 'daily' : 'every ' + freq + ' days') + ', ' +
    (body.ticket ? 'ticket ID required' : 'DONE confirmation only') + '.');

  SpreadsheetApp.flush();
  return { ok: true, message: id + ' — ' + name + ', assigned to ' + employee + '.' };
}

function apiAcknowledge(body) {
  const task = findRow(body.taskId);
  if (!task) return { ok: false, error: 'Task not found: ' + body.taskId };
  if (task.mgrAck) return { ok: true, message: task.id + ' was already acknowledged.' };

  const sheet = sheetOf(SHEET_TASK_MASTER);
  const now = new Date();
  sheet.getRange(task.row, TM.MGR_ACK + 1).setValue(true);
  sheet.getRange(task.row, TM.ACK_DATE + 1).setValue(now);

  logEvent(task.id, task.name, task.employee, 'Manager Acknowledged',
    'Manager acknowledged from the dashboard. The unresolved ceiling still applies.');

  SpreadsheetApp.flush();
  return { ok: true, message: task.id + ' acknowledged.' };
}

function apiDailyCheck() {
  const sla = getSla();
  const sheet = sheetOf(SHEET_TASK_MASTER);
  const tasks = readTasks();
  const now = new Date();
  let alerts = 0, escalations = 0;

  tasks.forEach(function (t) {
    if (t.status !== 'Overdue' && t.status !== 'Due Today') return;

    // same-day manager alert, once per cycle
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
      return;
    }

    if (t.escalated) return;   // never escalate the same cycle twice

    const hoursSinceAlert = (now - new Date(t.alertSent)) / 36e5;
    const hoursSinceDue   = t.due ? (now - new Date(t.due)) / 36e5 : 0;
    const noAck      = !t.mgrAck && hoursSinceAlert >= sla.ackWindowHours;
    const tooLong    = hoursSinceDue >= sla.maxWindowHours;

    if (noAck || tooLong) {
      const reason = noAck
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
  return {
    ok: true,
    message: alerts + ' manager alert(s), ' + escalations + ' escalation(s) sent.',
    alerts: alerts, escalations: escalations
  };
}
