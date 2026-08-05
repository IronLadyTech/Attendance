# LEP Checkpoint Attendance — Zoho Flow Wiring

**IL LEP Sessions - 1 & 2 Days** — CRM updated at **every check** (1, 2, 3) and **final majority** (4).

Wire a separate **LEP_Attendance** Flow (not Mc_Attendance or 100BM_Attendance).

## Schedule

| Day | Check 1 | Check 2 | Check 3 | Final (majority) |
|-----|---------|---------|---------|------------------|
| Day 1 (Sat) | 9:15 AM | 3:30 PM | 6:15 PM | 6:30 PM |
| Day 2 (Sun) | 9:15 AM | 12:30 PM | 4:15 PM | 4:30 PM |

Saturday = Day 1. Sunday = Day 2. `batch_date` = Saturday reg date (Day 2 uses session date − 1, same as MC).

## What updates when

| When | Bridge event | CRM update |
|------|--------------|------------|
| Check 1, 2, 3 | `attendance.lep_check` | Each person in meeting → Present/Absent |
| Check 1, 2, 3 | `attendance.lep_batch_check` | Cohort not in room → **Absent** (sales calls) |
| Final | `attendance.lep_final` | Majority Present/Absent (overwrites field) |

**One function** handles all three events: `mark_lep_attendance_final`.

## Decision branches (top to bottom)

| Condition | Event | Action |
|-----------|-------|--------|
| condition1 | `attendance.lep_check` | `mark_lep_attendance_final` |
| condition2 | `attendance.lep_batch_check` | `mark_lep_attendance_final` |
| condition3 | `attendance.lep_final` | `mark_lep_attendance_final` |
| Default | other | *(no action)* |

## CRM module & fields

**Module:** `Session_Attendance` (Session Attendance — not Leads)

| Field (UI) | API name |
|------------|----------|
| Email | `Email` |
| Lead Email | `Lead_Email` |
| Lead Full Name | `Name` |
| LEP Start Date | `LEP_Start_Date` → Saturday batch (cohort) |
| LEP Day 1 Session | `LEP_Day_1_Session` → Present / Absent |
| LEP Day 2 Session | `LEP_Day_2_Session` → Present / Absent |

## Parameter mapping — all branches use same function

### lep_check / lep_final

| Param | Value |
|-------|--------|
| participant_email | `${webhookTrigger.payload.participant_email}` |
| participant_name | `${webhookTrigger.payload.participant_name}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |
| session_day | `${webhookTrigger.payload.session_day}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_number | `${webhookTrigger.payload.check_number}` |
| present_emails | *(leave blank)* |
| attendance_result | `${webhookTrigger.payload.attendance_result}` |

### lep_batch_check

| Param | Value |
|-------|--------|
| participant_email | *(leave blank)* |
| participant_name | *(leave blank)* |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |
| session_day | `${webhookTrigger.payload.session_day}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_number | `${webhookTrigger.payload.check_number}` |
| present_emails | `${webhookTrigger.payload.present_emails}` |
| attendance_result | *(leave blank)* |

## Record matching

**Individual:** `LEP_Start_Date` = batch first; else any match on `Email` / `Lead_Email` / `Name`.

**Batch (sales):** `LEP_Start_Date` = batch, `Email` or `Lead_Email` **not** in `present_emails` → Absent.

## Render

| Variable | Purpose |
|----------|---------|
| `ZOOM_WEBHOOK_SECRET_TOKEN_LEP` | Zoom `/lep` |
| `ZOHO_WEBHOOK_FORWARD_URL_LEP` | This Flow webhook |
| `BRIDGE_STATE_PATH` | Persist roster/checks across restarts |

## No Day 2 blueprint

LEP does not run MC Completed–style blueprint on meeting end.
