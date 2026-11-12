# Recurring Task Automation

A recurring task tracking, pending report and escalation system built on Google Sheets + Apps Script, with an interactive web demo.

**[▶ Open the live demo](https://suraj944.github.io/recurring-task-automation/)**

## The problem

Recurring operational work (daily reports, weekly compliance checks, ticket-based portal tasks) gets tracked in chat threads and memory. Things slip silently, and the founder becomes the escalation path for everything.

## How it works

One sheet holds the data. Everything else is a view or an automation on top of it.

| Tab | Role |
|---|---|
| **Task Master** | Single source of truth — 16 columns, locked to the owner's account |
| **Employee Pending View** | Each employee picks their name, sees only their pending tasks, and submits one update |
| **Manager Dashboard** | Every pending/overdue task across all employees, with an acknowledgment checkbox |
| **SLA Settings** | Alert time, acknowledgment window, escalation ceiling — read live by the script, never hardcoded |
| **Escalation Log** | Append-only audit trail of every completion, alert, acknowledgment and escalation |

### Automation rules

- **Single point of update** — an employee types `DONE` (or a ticket ID) in one cell. The script writes the completion to Task Master, recalculates the next cycle, and clears the previous cycle's flags. No one updates status by hand.
- **Same-day manager alert** — a daily time-driven trigger emails the manager the day a task becomes due or overdue, once per cycle.
- **Escalation on two independent conditions** — no manager acknowledgment within the acknowledgment window, *or* the task still unresolved past the hard ceiling regardless of acknowledgment.
- **No duplicate emails** — each alert and escalation writes a date stamp back to Task Master and is never re-sent for the same cycle.
- **Ticket validation** — a task flagged as requiring a ticket rejects a bare `DONE`.

### Design decisions worth noting

- **Installable triggers, not simple triggers.** Apps Script's `onEdit` runs with the *editing user's* permissions. Since Task Master is locked to the owner, an employee's edit would silently fail to write. `setupTriggers()` is run once by the owner so every automated write afterwards carries owner authority.
- **Task IDs are validated unique.** A duplicate ID makes the script's row lookup ambiguous and completions land on the wrong row. Column A enforces uniqueness, with red highlighting as a visual backstop.
- **Timing lives in the sheet, not the code.** Changing the alert hour or escalation window means editing a cell, not redeploying a script.

## Repo contents

- `index.html` — the interactive demo (no build step, no dependencies, runs from a file)
- `AppsScript.gs` — the complete automation script

## The demo

Runs entirely in the browser on mock data — nothing connects to a real sheet. You can:

- complete a task as an employee and watch the next cycle calculate
- run the daily check and see manager alerts, then leadership escalations, fire in sequence
- acknowledge as a manager and watch the escalation path change
- add a new task and see it flow through to every view

## Running locally

Open `index.html` in a browser. That's it.
