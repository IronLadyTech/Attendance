# MC Checkpoint Attendance — Zoho Flow Wiring

After deploying the updated bridge, rewire **Mc_Attendance** as follows.

## Zoom app (MC account)

Enable event subscription: **`meeting.started`** (in addition to join/leave/ended).

## Decision branches

| Condition | Event | Action |
|-----------|-------|--------|
| condition1 | `attendance.mark_yes` | `mark_attendance_yes` |
| condition2 | `attendance.lookup` | `mark_attendance_lookup` (joined but left before T+15) |
| condition3 | `attendance.first_check` | `mark_batch_attendance_checkpoint` (check_type = **first**) **+** `send_attendance_report` (check_type = **first**) |
| condition4 | `attendance.final_check` + Day 1 topic | `mark_batch_attendance_checkpoint` (final) **+** `send_attendance_report` (final) |
| condition5 | `attendance.hour_check` | `mark_batch_attendance_checkpoint` (check_type = **hour**) **+** `send_attendance_report` (check_type = **hour**) |
| condition6 | `attendance.final_check` + Day 2 topic | `mark_batch_attendance_checkpoint` (final) **+** `send_attendance_report` (final) **+** `mark_mc_completed` |
| condition5 | `meeting.ended` + Day 1 topic | *(none)* |
| condition6 | `meeting.ended` + Day 2 topic | `mark_mc_completed` only |

**Remove or disable:**
- `set_blank_automation_not_attended` on meeting.ended

**Add:**
- `attendance.mark_no` → `mark_attendance_no` (T+30: joined but not in room at final check)

## Parameter mapping

### mark_attendance_yes

| Param | Value |
|-------|--------|
| meeting_id | `${webhookTrigger.payload.meeting_id}` |
| participant_email | `${webhookTrigger.payload.participant_email}` |
| participant_name | `${webhookTrigger.payload.participant_name}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| session_date | `${webhookTrigger.payload.session_date}` |

### mark_attendance_no (T+30 — dropped: was in room at T+15, left before T+30 → **No**)

Same mapping as `mark_attendance_yes`.

| Param | Value |
|-------|--------|
| meeting_id | `${webhookTrigger.payload.meeting_id}` |
| participant_email | `${webhookTrigger.payload.participant_email}` |
| participant_name | `${webhookTrigger.payload.participant_name}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| session_date | `${webhookTrigger.payload.session_date}` |

### mark_attendance_lookup (attendance.lookup at T+15)

| Param | Value |
|-------|--------|
| participant_email | `${webhookTrigger.payload.participant_email}` |
| participant_name | `${webhookTrigger.payload.participant_name}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| session_date | `${webhookTrigger.payload.session_date}` |

## Every joiner: CRM or sheet

| Situation | What happens |
|-----------|----------------|
| In meeting at T+15 or T+30 | `mark_attendance_yes` → **Yes** in CRM, or **unmatched sheet** if no lead |
| Joined but left before T+15 | `mark_attendance_lookup` → CRM lead gets **No** via batch; guest → **sheet** |
| Never joined Zoom (CRM lead) | Batch **first** → No (sales); **final** → **Absent** |
| Joined but left before T+30 | `mark_attendance_no` → **No** |
| No email and no name on join | Bridge skips (cannot identify) |

### mark_batch_attendance_checkpoint (first_check)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_type | `first` (literal) |
| ever_joined_emails | *(leave empty)* |
| present_emails | *(leave empty)* |

### mark_batch_attendance_checkpoint (final_check)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_type | `final` (literal) |
| ever_joined_emails | `${webhookTrigger.payload.ever_joined_emails}` |
| present_emails | `${webhookTrigger.payload.present_emails}` |

### mark_batch_attendance_checkpoint + send_attendance_report (hour_check)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_type | `hour` (literal) |
| ever_joined_emails | `${webhookTrigger.payload.ever_joined_emails}` |
| present_emails | `${webhookTrigger.payload.present_emails}` |

**Hour check logic:** in room at T+60 + CRM was **No** or **Absent** → **Yes**. Already **Yes** → unchanged. Not in room → unchanged (No stays No, Absent stays Absent).

> Note: the bridge now also sends `ever_joined_emails` and `present_emails` on **first_check**. For the attendance update they can stay empty on first_check, but `send_attendance_report` needs them (see below).

### send_attendance_report (first_check AND final_check — second action on each branch)

Same signature and mapping as `mark_batch_attendance_checkpoint`. Map ALL six params from the payload (do **not** leave the email lists empty — the report needs them):

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_type | `first` on the first_check branch / `final` on the final_check branch / `hour` on the hour_check branch (literal) |
| ever_joined_emails | `${webhookTrigger.payload.ever_joined_emails}` |
| present_emails | `${webhookTrigger.payload.present_emails}` |

Emails the report to `ironladytech@gmail.com`, `yaswanthgrandhi2580@gmail.com`, and `brunda@iamironlady.com` (change `reportTo` in the function to adjust). Per-participant Yes/No/Absent + reason, counts, and a guests section for joiners with no CRM match.

### mark_mc_completed (Day 2 only)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |

## Timeline

```
T+0   meeting.started → bridge schedules sweeps
T+15  mark_yes (in meeting) + lookup (joined-left) + first_check → blank batch → No
T+30  mark_yes (in room) + mark_no (dropped → No) + final_check → never joined → Absent
T+60  mark_yes (in room) + hour_check → No/Absent in room → Yes; not in room → unchanged
End   meeting.ended → mark_mc_completed (Day 2 only)
```

## 100BM

Uses the same T+15 / T+30 checkpoint model on `/100bm`. See **`100BM_CHECKPOINT_FLOW_WIRING.md`** for Zoho Flow setup and separate Deluge functions (`mark_100bm_*`).
