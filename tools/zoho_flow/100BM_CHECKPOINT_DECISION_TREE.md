# 100BM_Attendance — Decision Tree (Checkpoint Model)

Mirrors the MC checkpoint model (T+15 / T+30, Yes / No / Absent) using the **FT fields**.

## Visual tree

```
                        WEBHOOK (incoming from Render bridge /100bm)
                                    │
                                    ▼
                              ┌───────────┐
                              │ DECISION  │
                              └───────────┘
                                    │
     ┌──────────┬──────────┬───────┴───────┬──────────────┬──────────┐
     ▼          ▼          ▼               ▼              ▼          ▼
 cond1      cond2      cond3           cond4          cond5      Default
 mark_yes   mark_no    lookup       first_check    final_check  (no action)
     │          │          │               │              │
     ▼          ▼          ▼               ▼              ▼
 mark_      mark_      mark_          mark_          mark_
 100bm_     100bm_     100bm_         100bm_batch_   100bm_batch_
 attend_    attend_    attend_        checkpoint     checkpoint
 yes        no         lookup         (first)        (final)
```

---

## Condition details (copy into Zoho Flow Decision)

Check conditions **top to bottom**. First match wins.

### condition1 — Present at checkpoint → Yes

```
Payload > Event  equals  attendance.mark_yes
```

**Action:** `mark_100bm_attendance_yes`
Matches lead by **email/name** (any invite date). Sets `FT_attendance = Yes` and `FT_attended_date = actual session date`. **Does not change** `FT_Invite_Date` (supports makeup sessions).

---

### condition2 — T+30 dropout → No

```
Payload > Event  equals  attendance.mark_no
```

**Action:** `mark_100bm_attendance_no`
Joined earlier but not in the room at T+30 → `FT_attendance = No` (downgrades a T+15 Yes).

---

### condition3 — Joined but left before T+15 → lookup

```
Payload > Event  equals  attendance.lookup
```

**Action:** `mark_100bm_attendance_lookup`
Eligible CRM lead → skip (batch first_check marks No). No lead → log to unmatched sheet.

---

### condition4 — T+15 batch sweep → blank leads get No

```
Payload > Event  equals  attendance.first_check
```

**Action:** `mark_100bm_batch_attendance_checkpoint`
- `check_type` = `first` (literal text)

---

### condition5 — T+30 batch safety net → Yes / No / Absent

```
Payload > Event  equals  attendance.final_check
```

**Action:** `mark_100bm_batch_attendance_checkpoint`
- `check_type` = `final` (literal text)
- `ever_joined_emails` = `${webhookTrigger.payload.ever_joined_emails}`
- `present_emails` = `${webhookTrigger.payload.present_emails}`

---

### Default

No action. Events that fall through (e.g. `meeting.ended`, unknown events).

---

## Eligibility

A lead is processed when **`FT_Invite_Date` matches the session date** (IST date from meeting `start_time`). **No `Payment_Status` filter.**

---

## What the bridge sends (100BM route `/100bm`)

| When | `event` value |
|------|----------------|
| In meeting at T+15 or T+30 | `attendance.mark_yes` (one per person) |
| Joined but left before T+15 | `attendance.lookup` (one per person) |
| Joined but not in room at T+30 | `attendance.mark_no` (one per person) |
| T+15 sweep done | `attendance.first_check` |
| T+30 sweep done | `attendance.final_check` (+ present_emails, ever_joined_emails) |
| Meeting ends | `meeting.ended` (no attendance action) |

---

## Timeline (one session)

```
T+0     meeting.started (bridge only — not a Flow branch)
T+15    mark_yes  (each in room)        → Yes + FT_attended_date
        lookup    (each joined-then-left)
        first_check → batch blank → No
T+30    mark_yes  (each in room)        → Yes + FT_attended_date (upgrades No→Yes)
        mark_no   (each joined-left)    → Yes→No dropout
        final_check → batch safety net:
            present → Yes | joined-left → No | never joined → Absent
```

---

## Per-person outcomes

```
                      PARTICIPANT JOINS SESSION?
                               │
               ┌───────────────┴───────────────┐
               │ NO                            │ YES
               ▼                               ▼
     FT_attendance:                   In room at T+15?
       T+15 → No                            │
       T+30 → Absent            ┌───────────┴───────────┐
                                │ YES                   │ NO (left early)
                                ▼                       ▼
                          T+15 → Yes              T+15 → No
                                │                       │
                          In room at T+30?        In room at T+30?
                                │                       │
                        ┌───────┴───────┐       ┌───────┴───────┐
                        │ YES           │ NO    │ YES           │ NO
                        ▼               ▼       ▼               ▼
                   T+30 → Yes      T+30 → No  T+30 → Yes   T+30 → No
                   (attended)      (dropout)  (late join)  (joined-left)
```

---

## Parameter quick reference

| Function | Key parameters |
|----------|----------------|
| mark_100bm_attendance_yes | meeting_id, participant_email, participant_name, meeting_topic, session_date |
| mark_100bm_attendance_no | meeting_id, participant_email, participant_name, meeting_topic, session_date |
| mark_100bm_attendance_lookup | participant_email, participant_name, meeting_topic, session_date |
| mark_100bm_batch_attendance_checkpoint | start_time, meeting_topic, session_date, check_type, ever_joined_emails, present_emails |

All map from `${webhookTrigger.payload.<field>}` except `check_type` (literal `first` or `final`).

---

## Difference vs MC

| | MC | 100BM |
|--|----|-------|
| Attendance field | `Day_1_Attendance` / `Day_2_Attendance` | `FT_attendance` |
| Attended date | — | `FT_attended_date` = session date on **Yes / No / Absent**; invite date unchanged |
| Eligibility | `Payment_Status = Completed` + `MC_Start_Date_Time` = batch | `FT_Invite_Date` = session date |
| Sessions | Day 1 + Day 2 (topic split) | Single FastTrack session |
| End of program | `mark_mc_completed` on Day 2 | none |
