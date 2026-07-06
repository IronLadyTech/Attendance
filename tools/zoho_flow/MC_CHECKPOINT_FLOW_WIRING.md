# MC Checkpoint Attendance — Zoho Flow Wiring

After deploying the updated bridge, rewire **Mc_Attendance** as follows.

## Zoom app (MC account)

Enable event subscription: **`meeting.started`** (in addition to join/leave/ended).

## Decision branches

| Condition | Event | Action |
|-----------|-------|--------|
| condition1 | `attendance.mark_yes` | `mark_attendance_yes` |
| condition2 | `attendance.lookup` | `mark_attendance_lookup` (joined but left before T+15) |
| condition3 | `attendance.first_check` | `mark_batch_attendance_checkpoint` (check_type = **first**) |
| condition4 | `attendance.final_check` | `mark_batch_attendance_checkpoint` (check_type = **final**) |
| condition5 | `meeting.ended` + Day 1 topic | *(none)* |
| condition6 | `meeting.ended` + Day 2 topic | `mark_mc_completed` only |

**Remove or disable:**
- `attendance.mark_no` branch (MC no longer sends this)
- `set_blank_automation_not_attended` on meeting.ended

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
| Never joined (CRM lead) | Batch **first** → No, **final** → Absent |
| No email and no name on join | Bridge skips (cannot identify) |

### mark_batch_attendance_checkpoint (first_check)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_type | `first` (literal) |

### mark_batch_attendance_checkpoint (final_check)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| batch_date | `${webhookTrigger.payload.batch_date}` |
| check_type | `final` (literal) |

### mark_mc_completed (Day 2 only)

| Param | Value |
|-------|--------|
| start_time | `${webhookTrigger.payload.start_time}` |

## Timeline

```
T+0   meeting.started → bridge schedules sweeps
T+15  mark_yes (in meeting) + lookup (joined-left) + first_check → blank batch → No
T+30  mark_yes (in meeting) + final_check → blank batch → Absent
End   meeting.ended → mark_mc_completed (Day 2 only)
```

## 100BM

Unchanged — `/100bm` route still uses 30-min timer and separate Flow (if configured).
