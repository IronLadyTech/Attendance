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

| When | Bridge event | CRM update | Report |
|------|--------------|------------|--------|
| Check 1, 2, 3 | `attendance.lep_check` | Each Zoom person → Present/Absent | — |
| Check 1, 2, 3 | `attendance.lep_batch_check` | Cohort **not in room** → **Absent** | **`send_lep_attendance_report`** |
| Final | `attendance.lep_final` | Majority Present/Absent (people who joined Zoom) | — |
| Final | `attendance.lep_final_batch` | Cohort who **never joined Zoom** → **Absent** | — |
| Final | `attendance.lep_final_report` | — (summary webhook) | **`send_lep_attendance_report`** |

**CRM function:** `mark_lep_attendance_final` (10 params — includes `present_names`)  
**Report function:** `send_lep_attendance_report`

## Decision branches (top to bottom)

| Condition | Event | Actions |
|-----------|-------|---------|
| condition1 | `attendance.lep_check` | `mark_lep_attendance_final` |
| condition2 | `attendance.lep_batch_check` | `mark_lep_attendance_final` **then** `send_lep_attendance_report` |
| condition3 | `attendance.lep_final` | `mark_lep_attendance_final` |
| condition4 | `attendance.lep_final_batch` | `mark_lep_attendance_final` |
| condition5 | `attendance.lep_final_report` | `send_lep_attendance_report` |
| Default | other | *(no action)* |

## CRM module & fields

**Module:** `Session_Attendance`

| Field (UI) | API name |
|------------|----------|
| Email | `Email` |
| Lead Email | `Lead_Email` |
| Lead Full Name | `Name` |
| LEP Start Date | `LEP_Start_Date` |
| LEP Day 1 Session | `LEP_Day_1_Session` |
| LEP Day 2 Session | `LEP_Day_2_Session` |

## Parameter mapping — mark_lep_attendance_final

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
| present_names | *(leave blank)* |

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
| present_names | `${webhookTrigger.payload.present_names}` |

### lep_final_batch (never joined → Absent)

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
| present_names | `${webhookTrigger.payload.present_names}` |

At final, `present_emails` / `present_names` = everyone who **ever joined** Zoom. Anyone else in the CRM cohort is set to **Absent**.

## Parameter mapping — send_lep_attendance_report

| Param | Value |
|-------|--------|
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |
| session_day | `${webhookTrigger.payload.session_day}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_number | `${webhookTrigger.payload.check_number}` |
| present_emails | `${webhookTrigger.payload.present_emails}` |
| ever_joined_emails | `${webhookTrigger.payload.ever_joined_emails}` |

## Final attendance rules

| Person | Result |
|--------|--------|
| Joined Zoom, majority Present (2+ of 3) | **Present** |
| Joined Zoom, majority Absent | **Absent** |
| In today’s batch, **never joined Zoom** | **Absent** (via `lep_final_batch`) |

## Render

| Variable | Purpose |
|----------|---------|
| `ZOOM_WEBHOOK_SECRET_TOKEN_LEP` | Zoom `/lep` |
| `ZOHO_WEBHOOK_FORWARD_URL_LEP` | This Flow webhook |
| `BRIDGE_STATE_PATH` | Persist roster/checks |

Redeploy Render after bridge changes (adds `present_names` + `lep_final_batch`).

## No Day 2 blueprint

LEP does not run MC Completed–style blueprint on meeting end.
