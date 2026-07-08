# 100BM Checkpoint Attendance — Zoho Flow Wiring

Mirrors the MC checkpoint model (T+15 / T+30 with Yes / No / Absent), using **FT fields**.
Wire a separate **100BM_Attendance** Flow (not Mc_Attendance).

## Zoom app (100BM account)

Enable event subscriptions:

- `meeting.started`
- `meeting.participant_joined`
- `meeting.participant_left`
- `meeting.ended` (optional — no attendance action)

Webhook endpoint: bridge **`POST /100bm`** with `ZOOM_WEBHOOK_SECRET_TOKEN_100BM`.

## Decision branches

| Condition | Event | Action |
|-----------|-------|--------|
| condition1 | `attendance.mark_yes` | `mark_100bm_attendance_yes` |
| condition2 | `attendance.mark_no` | `mark_100bm_attendance_no` |
| condition3 | `attendance.lookup` | `mark_100bm_attendance_lookup` |
| condition4 | `attendance.first_check` | `mark_100bm_batch_attendance_checkpoint` (check_type = **first**) |
| condition5 | `attendance.final_check` | `mark_100bm_batch_attendance_checkpoint` (check_type = **final**) |
| Default | `meeting.ended`, etc. | *(no action)* |

Check conditions top to bottom; first match wins.

## CRM fields

| UI label | API name | Written |
|----------|----------|---------|
| FT Attendance | `FT_attendance` | Yes / No / Absent |
| FT Attended Date | `FT_attended_date` | Set on **Yes** only; cleared on **No** / **Absent** |
| FT Invite Date | `FT_Invite_Date` | Eligibility filter (= session date); not updated |

## Eligibility

Leads are processed when **`FT_Invite_Date` matches the session date** (IST date from meeting `start_time`). No `Payment_Status` filter.

## Parameter mapping

### mark_100bm_attendance_yes / mark_100bm_attendance_no

| Param | Value |
|-------|--------|
| meeting_id | `${webhookTrigger.payload.meeting_id}` |
| participant_email | `${webhookTrigger.payload.participant_email}` |
| participant_name | `${webhookTrigger.payload.participant_name}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |

### mark_100bm_attendance_lookup

| Param | Value |
|-------|--------|
| participant_email | `${webhookTrigger.payload.participant_email}` |
| participant_name | `${webhookTrigger.payload.participant_name}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |

### mark_100bm_batch_attendance_checkpoint (first_check)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |
| check_type | `first` (literal) |
| ever_joined_emails | *(leave blank)* |
| present_emails | *(leave blank)* |

### mark_100bm_batch_attendance_checkpoint (final_check)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |
| check_type | `final` (literal) |
| ever_joined_emails | `${webhookTrigger.payload.ever_joined_emails}` |
| present_emails | `${webhookTrigger.payload.present_emails}` |

## Attendance outcomes

| Situation | T+15 | T+30 (final) | FT_attended_date |
|-----------|------|--------------|------------------|
| In room the whole time | **Yes** | **Yes** | session date |
| Joins late (after 15), stays till 30 | **No** | **Yes** | session date |
| Joins, leaves before 15 | **No** | **No** | — |
| In room at 15, leaves before 30 | **Yes** | **No** (dropout) | — |
| Never joins (eligible lead) | **No** | **Absent** | — |
| Guest (no CRM lead) | unmatched sheet | unmatched sheet | — |

## Timeline

```
T+0     meeting.started (bridge schedules T+15 / T+30)
T+15    mark_yes (in room) → Yes + FT_attended_date
        lookup (joined-left) → guests to sheet
        first_check → batch blank → No
T+30    mark_yes (in room) → Yes + FT_attended_date (upgrades No → Yes)
        mark_no (joined-left dropout) → Yes → No
        final_check → batch safety net:
            present → Yes | joined-left → No | never joined → Absent
End     meeting.ended (optional; no CRM action)
```

## Bridge env vars

| Variable | Default | Purpose |
|----------|---------|---------|
| `BM100_CHECKPOINT_1_SECONDS` | 900 | T+15 |
| `BM100_CHECKPOINT_2_SECONDS` | 1800 | T+30 |
| `ZOHO_WEBHOOK_FORWARD_URL_100BM` | — | Separate 100BM Flow URL |
