# Mc_Attendance — Decision Tree (MC Checkpoint Model)

## Visual tree

```
                        WEBHOOK (incoming from Render bridge)
                                    │
                                    ▼
                              ┌───────────┐
                              │ DECISION  │
                              └───────────┘
                                    │
     ┌──────────┬──────────┬───────┴───────┬──────────┬──────────┐
     ▼          ▼          ▼               ▼          ▼          ▼
 cond1      cond2      cond3           cond4a       cond4b     Default
 mark_yes   lookup    first_check    final D1     final D2    (no action)
     │          │          │               │          │
     ▼          ▼          ▼               ▼          ▼
 mark_      mark_      mark_batch      mark_batch   mark_batch
 attend_    attend_    checkpoint      checkpoint   checkpoint
 yes        lookup     (first)         (final)      (final)
                                              │          │
                                              ▼          ▼
                                           (stop)   mark_mc_completed
```

---

## Condition details (copy into Zoho Flow Decision)

Check conditions **top to bottom**. First match wins.

### condition1 — Someone present at checkpoint → Yes

```
Payload > Event  equals  attendance.mark_yes
```

**Action:** `mark_attendance_yes`

### condition1b — T+30: dropped (joined, left before T+30) → **No**

```
Payload > Event  equals  attendance.mark_no
```

**Action:** `mark_attendance_no` (same parameter mapping as `mark_attendance_yes`)

---


### condition2 — Joined but left before T+15 → lookup or sheet

```
Payload > Event  equals  attendance.lookup
```

**Action:** `mark_attendance_lookup`

---

### condition3 — T+15 batch sweep → blank leads get No

```
Payload > Event  equals  attendance.first_check
```

**Action:** `mark_batch_attendance_checkpoint`  
- `check_type` = `first` (literal text)

---

### condition4a — T+30 Day 1 final → never joined Zoom → **Absent**

```
Payload > Event  equals  attendance.final_check
AND
( Add Group )
  Payload > Meeting Topic  contains  BHAG
  OR
  Payload > Meeting Topic  contains  breakthrough actions
```

**Action:** `mark_batch_attendance_checkpoint`  
- `check_type` = `final` (literal text)

---

### condition4b — T+30 Day 2 final → never joined → **Absent**, then MC Completed

```
Payload > Event  equals  attendance.final_check
AND
( Add Group )
  Payload > Meeting Topic  contains  art of war
  OR
  Payload > Meeting Topic  contains  shameless pitching
```

**Actions (in order — same branch, chained):**

1. `mark_batch_attendance_checkpoint` — `check_type` = `final`, `ever_joined_emails` = `${webhookTrigger.payload.ever_joined_emails}`
2. `mark_mc_completed` — `start_time` = `${webhookTrigger.payload.start_time}`

---

### Default

No action. Events that fall through:

- `attendance.update_duration`
- `meeting.started` (handled in bridge only)
- Unknown events

---

## Remove these old branches

| Old condition | Remove? |
|---------------|---------|
| `attendance.mark_no` on **participant_left** (old timer) | **Remove** — use T+30 `mark_no` for dropouts only |
| `set_blank_automation_not_attended` | **Remove** — replaced by checkpoint |
| `meeting.ended` → mark_mc_completed | **Remove** — use T+30 Day 2 final_check instead |
| `meeting.ended` + Day 1 | **Optional** — no action needed |

---

## What the bridge sends (MC route `/`)

| When | `event` value |
|------|----------------|
| In meeting at T+15 or T+30 | `attendance.mark_yes` (one per person) |
| Joined but left before T+15 | `attendance.lookup` (one per person) |
| Joined but left before T+30 (dropped) | `attendance.mark_no` → **No** |
| Never joined Zoom (CRM lead) | T+15 batch → **No**; T+30 batch → **Absent** |
| T+15 sweep done | `attendance.first_check` |
| T+30 sweep done | `attendance.final_check` |
| Meeting ends | `meeting.ended` (optional; no attendance action) |

---

## Timeline (one session)

```
T+0     meeting.started (bridge only — not a Flow branch)
T+15    mark_yes (each in room)
        lookup (each who joined then left)
        first_check → batch blank → No
T+30    mark_yes (each in room — upgrades No → Yes)
        final_check Day 1 → batch blank → No
        final_check Day 2 → batch blank → No → mark_mc_completed
```

---

## Parameter quick reference

| Function | Key parameters |
|----------|----------------|
| mark_attendance_yes | email, name, meeting_topic, batch_date, session_date |
| mark_attendance_lookup | email, name, meeting_topic, batch_date, session_date |
| mark_batch_attendance_checkpoint | start_time, meeting_topic, batch_date, check_type, ever_joined_emails (final only) |
| mark_mc_completed | start_time |

All map from `${webhookTrigger.payload.<field>}` except `check_type` (literal `first` or `final`).
