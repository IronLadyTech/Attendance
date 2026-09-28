# Fresh LEP FT Checkpoint Attendance — Zoho Flow Wiring

Mirrors the 100BM FT checkpoint model (T+15 / T+30 / T+60 with Yes / No / Absent).
**New Flow** — do not reuse `100BM_Attendance`.

| | Value |
|--|--------|
| Zoom topic | **Fast Track Session - Iron Lady** |
| Anchor | **8:00 PM IST** (fixed; ignore early/late Zoom start) |
| Check 1 | **8:15 PM** (T+15) |
| Check 2 | **8:30 PM** (T+30) |
| Check 3 | **9:00 PM** (T+60) |
| Eligibility | `Fresh_LEP_FT_Confirmation` = **Yes** |
| Writes | `Fresh_LEP_FT_Attendance` (Yes / No / Absent) |

> Verify API names in CRM Setup → Leads → Fields. If different, update the Deluge scripts.

## Create Flow: `Fresh_LEP_FT_Attendance`

1. Trigger: **Incoming webhook**
2. Decision node (conditions top → bottom)

| Condition | Event | Action |
|-----------|-------|--------|
| condition1 | `attendance.mark_yes` | `mark_fresh_lep_ft_attendance_yes` |
| condition2 | `attendance.mark_no` | `mark_fresh_lep_ft_attendance_no` |
| condition3 | `attendance.lookup` | `mark_fresh_lep_ft_attendance_lookup` |
| condition4 | `attendance.first_check` | `mark_fresh_lep_ft_batch_checkpoint` (`check_type` = **first**) |
| condition5 | `attendance.final_check` | `mark_fresh_lep_ft_batch_checkpoint` (`check_type` = **final**) |
| condition6 | `attendance.hour_check` | `mark_fresh_lep_ft_batch_checkpoint` (`check_type` = **hour**) |
| Default | other | *(no action)* |

## Parameter mapping

### mark_fresh_lep_ft_attendance_yes / _no

| Param | Value |
|-------|--------|
| meeting_id | `${webhookTrigger.payload.meeting_id}` |
| participant_email | `${webhookTrigger.payload.participant_email}` |
| participant_name | `${webhookTrigger.payload.participant_name}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |

### mark_fresh_lep_ft_attendance_lookup

Same email / name / topic / session_date as above.

### mark_fresh_lep_ft_batch_checkpoint (first)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |
| check_type | `first` (literal) |
| ever_joined_emails | *(blank)* |
| present_emails | *(blank)* |

### mark_fresh_lep_ft_batch_checkpoint (final / hour)

| Param | Value |
|-------|--------|
| check_type | `final` or `hour` (literal) |
| ever_joined_emails | `${webhookTrigger.payload.ever_joined_emails}` |
| present_emails | `${webhookTrigger.payload.present_emails}` |
| (+ start_time, meeting_topic, session_date) |

## Connect Zoom → Render → this Flow

```
Zoom app (Event notification URL)
  = https://YOUR-SERVICE.onrender.com/fresh-lep-ft
        │
        ▼
Render bridge env:
  ZOOM_WEBHOOK_SECRET_TOKEN_FLEP_FT  = Zoom Secret Token
  ZOHO_WEBHOOK_FORWARD_URL_FLEP_FT   = your Zoho Flow webhook URL
        │
        ▼
Zoho Flow Fresh_LEP_FT_Attendance
```

1. Render → Environment → set the two vars above → Redeploy.
2. Zoom Marketplace → Event Subscriptions → endpoint = `https://YOUR-SERVICE.onrender.com/fresh-lep-ft` → Validate.
3. Subscribe: meeting started, participant joined, participant left.
4. Meeting title must contain **Fast Track Session** (e.g. `Fast Track Session - Iron Lady`).

## Outcomes (Confirmation = Yes only)

| Situation | 8:15 | 8:30 |
|-----------|------|------|
| In room whole time | Yes | Yes |
| Late join, still there at 8:30 | No | Yes |
| Joined, left before 8:15 | No | No |
| In at 8:15, gone by 8:30 | Yes | No |
| Never joined | No | Absent |
| Confirmation ≠ Yes | ignored | ignored |

## Deluge source files

- `mark_fresh_lep_ft_attendance_yes.deluge`
- `mark_fresh_lep_ft_attendance_no.deluge`
- `mark_fresh_lep_ft_attendance_lookup.deluge`
- `mark_fresh_lep_ft_batch_checkpoint.deluge`
