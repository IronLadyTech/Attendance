# Iron Lady MC Attendance Automation — System Documentation

---

## What This System Does (Simple Explanation)

When a participant joins the Iron Lady Masterclass Zoom meeting, the system **automatically watches** how long they stay. Based on that, it marks their attendance in Zoho CRM — no manual work needed.

When Day 2 meeting ends, the system automatically moves all participants from **"MC Enrolled"** to **"MC Completed"** in CRM with their attendance filled in.

---

## WORKFLOW DIAGRAMS

---

### DIAGRAM 1 — How a Single Participant's Attendance is Tracked

```
Participant joins Zoom Meeting
           │
           ▼
   ┌─────────────────────────────┐
   │  Does the participant have  │
   │  an email linked to their   │
   │  Zoom account?              │
   └─────────────────────────────┘
           │
     ┌─────┴──────┐
     │ NO         │ YES
     ▼            ▼
  ┌──────┐   5-Minute Timer Starts
  │SKIP  │        │
  │(guest│        │
  │user) │   ┌────┴─────────────────────────────────┐
  └──────┘   │  Does the participant stay 5+ minutes?│
             └────┬─────────────────────────────────┘
                  │
         ┌────────┴────────┐
         │ YES             │ NO
         │ (stayed 5+ min) │ (left before 5 min)
         ▼                 ▼
   ┌───────────┐     ┌───────────┐
   │ MARK YES  │     │ MARK NO   │
   │ in CRM    │     │ in CRM    │
   └───────────┘     └─────┬─────┘
                           │
                    Does the participant
                    REJOIN the meeting?
                           │
                  ┌────────┴────────┐
                  │ NO              │ YES
                  ▼                 ▼
            Stays as "No"    New 5-min timer starts
                                    │
                           ┌────────┴────────┐
                           │ YES             │ NO
                           │ (stays 5+ min)  │ (leaves again)
                           ▼                 ▼
                      ┌─────────┐      Stays as "No"
                      │ MARK YES│
                      │(replaces│
                      │  "No")  │
                      └─────────┘
```

---

### DIAGRAM 2 — What Happens During Day 1 Meeting

```
DAY 1 MEETING (Topic contains "BHAG" or "Breakthrough Actions")
─────────────────────────────────────────────────────────────────

Participant Joins
     │
     ▼
5-Minute Timer Starts
     │
     ├──── Stays 5+ minutes ──────► Day_1_Attendance_Automation = YES ✅
     │
     └──── Leaves before 5 min ───► Day_1_Attendance_Automation = NO ❌
               │
               └── Rejoins + stays 5 min ──► Changes to YES ✅


WHAT STAYS BLANK?
─────────────────
If a participant never joined at all → Day_1_Attendance_Automation stays BLANK
  (This will default to "No" when MC Completed runs on Day 2)


WHO IS NOT TRACKED?
────────────────────
Participants who join Zoom as a Guest (no Zoom account) → No email → SKIPPED
  → Their name + email logged to Unmatched Sheet for manual review
```

---

### DIAGRAM 3 — What Happens During Day 2 Meeting

```
DAY 2 MEETING (Topic contains "Art of War" or "Shameless Pitching")
─────────────────────────────────────────────────────────────────────

Same attendance tracking as Day 1 ↑
  → Day_2_Attendance_Automation = YES / NO / BLANK


WHEN THE HOST ENDS THE MEETING:
────────────────────────────────

Meeting Ends
     │
     ▼
System finds all participants in CRM where:
  ✔ Lead Status = "MC Enrolled"
  ✔ Payment Status = "Completed"
  ✔ MC Start Date = Day 1 date of this batch
     │
     ▼
For each participant, reads their attendance:

  Day 1 Attendance  │  Day 2 Attendance  │  Result in MC Completed
  ──────────────────┼───────────────────┼──────────────────────────
  Yes               │  Yes              │  Day1=Yes, Day2=Yes
  Yes               │  No               │  Day1=Yes, Day2=No
  Yes               │  Blank            │  Day1=Yes, Day2=No (default)
  No                │  Yes              │  Day1=No,  Day2=Yes
  No                │  No               │  Day1=No,  Day2=No
  No                │  Blank            │  Day1=No,  Day2=No (default)
  Blank             │  Yes              │  Day1=No,  Day2=Yes (default)
  Blank             │  No               │  Day1=No,  Day2=No
  Blank             │  Blank            │  Day1=No,  Day2=No (default)
     │
     ▼
Blueprint Transition fires for each participant:
  MC Enrolled ──────────────────────────► MC Completed
     │
     ▼
Summary email sent to ironladytech@gmail.com
  "MC Completed - 2026-06-19 (20/20)"
```

---

### DIAGRAM 4 — How the System Finds a Participant in CRM

```
Zoom sends: Name = "Sowmya Rao", Email = "soumya.v.6@gmail.com"
                         │
                         ▼
              ┌─────────────────────────┐
              │ Search CRM by EMAIL     │
              │ soumya.v.6@gmail.com    │
              └─────────────────────────┘
                         │
               ┌─────────┴─────────┐
               │ Found?            │ Not Found
               ▼                   ▼
          MATCH ✅          Try by NAME
                                   │
                            ┌──────▼──────────────────────┐
                            │ Split "Sowmya Rao"           │
                            │ First = "Sowmya"             │
                            │ Last  = "Rao"                │
                            │ Search: First=Sowmya &       │
                            │         Last=Rao             │
                            └──────────────────────────────┘
                                   │
                         ┌─────────┴─────────┐
                         │ Found?             │ Not Found
                         ▼                    ▼
                    MATCH ✅          Try full name as Last Name
                                      Search: Last_Name = "Sowmya Rao"
                                             │
                                   ┌─────────┴─────────┐
                                   │ Found?             │ Not Found
                                   ▼                    ▼
                              MATCH ✅           LOG TO UNMATCHED SHEET
                                                 (Name, Email, Session)
                                                 for manual review
```

---

### DIAGRAM 5 — Complete End-to-End System Flow

```
┌─────────────────────────────────────────────────────────────────────┐
│                         ZOOM MEETING                                 │
│                                                                      │
│  Participant A joins ───────────────────────────────────────────────►│
│  Participant B joins ───────────────────────────────────────────────►│
│  Participant B leaves at 3 min ────────────────────────────────────►│
│  Participant B rejoins ─────────────────────────────────────────────►│
│  Participant A's 5-min timer fires ────────────────────────────────►│
│  Participant B's 5-min timer fires ────────────────────────────────►│
│  Host ends meeting ────────────────────────────────────────────────►│
└────────────────────────────────────┬────────────────────────────────┘
                                     │ Zoom sends webhooks
                                     ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    WEBHOOK BRIDGE (Render Server)                    │
│                                                                      │
│  joined  → Start 5-min timer for Participant A                      │
│  joined  → Start 5-min timer for Participant B                      │
│  left    → Cancel timer for B → Send "mark_no" for B              │
│  joined  → Restart timer for B (rejoin)                            │
│  [timer] → 5 min done for A → Send "mark_yes" for A               │
│  [timer] → 5 min done for B → Send "mark_yes" for B               │
│  ended   → Day 2 topic detected → Send "meeting.ended"             │
└────────────────────────────────────┬────────────────────────────────┘
                                     │ Posts to Zoho Flow
                                     ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    ZOHO FLOW (Mc_Attendance)                         │
│                                                                      │
│  attendance.mark_no  ──► mark_attendance_no()  → CRM: B = No       │
│  attendance.mark_yes ──► mark_attendance_yes() → CRM: A = Yes      │
│  attendance.mark_yes ──► mark_attendance_yes() → CRM: B = Yes      │
│  meeting.ended       ──► mark_mc_completed()                        │
└────────────────────────────────────┬────────────────────────────────┘
                                     │
                                     ▼
┌─────────────────────────────────────────────────────────────────────┐
│                         ZOHO CRM                                     │
│                                                                      │
│  Participant A: Day_1_Attendance = Yes  │  Day_2_Attendance = Yes  │
│  Participant B: Day_1_Attendance = No   │  Day_2_Attendance = Yes  │
│  Participant C: Day_1_Attendance = No   │  Day_2_Attendance = No   │
│                 (never joined)          │  (never joined)          │
│                                     │                              │
│  All 3 leads: MC Enrolled ──────────────────────► MC Completed    │
│                                                                      │
│  📧 Email sent: "MC Completed - 2026-06-19 (3/3)"                  │
└─────────────────────────────────────────────────────────────────────┘
```

---

### DIAGRAM 6 — All Possible Attendance Outcomes for a Participant

```
                      PARTICIPANT JOINS MEETING?
                               │
               ┌───────────────┴───────────────┐
               │ NO                            │ YES
               ▼                               ▼
     Day_X_Attendance = BLANK        Does participant have email?
     (defaults to "No" at                      │
      MC Completed stage)           ┌──────────┴──────────┐
                                    │ NO (guest)          │ YES
                                    ▼                     ▼
                               SKIPPED             Stays 5+ minutes?
                               (logged to               │
                                sheet)         ┌────────┴────────┐
                                               │ YES             │ NO
                                               ▼                 ▼
                                       Attendance = YES   Attendance = NO
                                               │                 │
                                               │         Rejoins + 5+ min?
                                               │                 │
                                               │         ┌───────┴───────┐
                                               │         │ YES           │ NO
                                               │         ▼               ▼
                                               │   Changes to YES   Stays NO
                                               │
                                    ┌──────────┴──────────────────┐
                                    │   AT END OF DAY 2 MEETING   │
                                    │   mark_mc_completed runs    │
                                    └──────────┬──────────────────┘
                                               │
                              ┌────────────────┼────────────────┐
                              │                │                │
                              ▼                ▼                ▼
                         Att = YES        Att = NO         Att = BLANK
                              │                │                │
                              ▼                ▼                ▼
                    Blueprint fills      Blueprint fills   Blueprint fills
                    Day_X_Att = Yes      Day_X_Att = No    Day_X_Att = No
                                                           (default)
```

---

## Overview

Automatically tracks Zoom meeting attendance for the Iron Lady Leadership Masterclass (MC) and updates Zoho CRM leads. When Day 2 ends, transitions all enrolled leads to "MC Completed" via blueprint with their attendance values.

---

## Architecture (Technical)

```
Zoom Meeting
    │
    │ Webhooks (participant_joined / participant_left / meeting.ended)
    ▼
Webhook Bridge (Python — hosted on Render)
    │
    │ POST attendance.mark_yes / attendance.mark_no / meeting.ended
    ▼
Zoho Flow — Mc_Attendance
    │
    ├── attendance.mark_yes  →  mark_attendance_yes()
    ├── attendance.mark_no   →  mark_attendance_no()
    └── meeting.ended        →  mark_mc_completed()
                                        │
                                        ▼
                               Zoho CRM (Leads module)
                               Blueprint: MC Enrolled → MC Completed
```

---

## Components

### 1. Zoom App (Marketplace)
**Event Subscriptions required:**
| Event | Purpose |
|-------|---------|
| Meeting → Participant/Host joined | Starts 5-min presence timer |
| Meeting → Participant/Host left | Cancels timer, marks No if early |
| Meeting → End Meeting | Triggers MC Completed transition |

---

### 2. Webhook Bridge (Render)
**File:** `tools/zoom_webhook_bridge.py`  
**Hosted on:** Render (auto-deploys from GitHub `main` branch)

**Environment Variables:**
| Variable | Description |
|----------|-------------|
| `ZOOM_WEBHOOK_SECRET_TOKEN` | Zoom app Secret Token (for HMAC validation) |
| `ZOHO_WEBHOOK_FORWARD_URL` | Zoho Flow webhook URL (single URL for all events) |
| `PRESENCE_SECONDS` | Seconds before marking Yes (default: 300 = 5 min) |
| `PORT` | Listen port (Render sets this automatically) |

**How it works:**

| Zoom Event | Bridge Action | Sends to Zoho |
|-----------|--------------|---------------|
| `meeting.participant_joined` | Starts 5-min timer per participant | Nothing yet |
| `meeting.participant_left` (timer active) | Cancels timer | `attendance.mark_no` |
| `meeting.participant_left` (no timer) | Ignores | Nothing |
| Timer fires (5 min elapsed) | — | `attendance.mark_yes` |
| `meeting.ended` (Day 2 topic) | — | `meeting.ended` |
| `meeting.ended` (Day 1 topic) | Ignores | Nothing |
| `endpoint.url_validation` | Responds with HMAC | Nothing |

**Day 2 Topic Keywords** (case-insensitive):
- `"art of war"`
- `"shameless pitching"`

**Day 1 Topic Keywords** (case-insensitive):
- `"bhag"`
- `"breakthrough actions"`

**Participant matching requires email.** Participants who join Zoom without signing in (guests) have no email and are skipped.

---

### 3. Zoho Flow — Mc_Attendance

**Decision node conditions:**

| Condition | Routes to |
|-----------|-----------|
| `Payload > Event = attendance.mark_yes` | `mark_attendance_yes` |
| `Payload > Event = attendance.mark_no` | `mark_attendance_no` |
| `Payload > Event = meeting.ended` | `mark_mc_completed` |

**Parameter mappings:**

`mark_attendance_yes` / `mark_attendance_no`:
| Parameter | Mapping |
|-----------|---------|
| `meeting_id` | `${webhookTrigger.payload.meeting_id}` |
| `participant_email` | `${webhookTrigger.payload.participant_email}` |
| `participant_name` | `${webhookTrigger.payload.participant_name}` |
| `meeting_topic` | `${webhookTrigger.payload.meeting_topic}` |

`mark_mc_completed`:
| Parameter | Mapping |
|-----------|---------|
| `start_time` | `${webhookTrigger.payload.start_time}` |

---

### 4. Zoho CRM Functions

#### `mark_attendance_yes`
**Triggered:** When participant stays 5+ minutes in meeting  
**Action:** Sets `Day_1_Attendance_Automation = Yes` or `Day_2_Attendance_Automation = Yes`

**Lead matching order:**
1. Email exact match
2. First Name + Last Name (Title Case)
3. First Name + Last Name (UPPER)
4. First Name + Last Name (lower)
5. Full name as Last_Name (handles "Sowmya Rao" stored in Last_Name field)
6. Full name as Last_Name (UPPER / lower)

**If not found:** Logs to Zoho Sheet (unmatched participants)  
**If already Yes:** Skips (no downgrade)

---

#### `mark_attendance_no`
**Triggered:** When participant leaves before 5 minutes  
**Action:** Sets `Day_1_Attendance_Automation = No` or `Day_2_Attendance_Automation = No`

**Same lead matching logic as `mark_attendance_yes`**  
**Never downgrades Yes → No** (if already Yes, skips)

---

#### `mark_mc_completed`
**Triggered:** When Day 2 meeting ends  
**Action:** Transitions all MC Enrolled leads for that batch to MC Completed

**Step-by-step:**
1. Extract Day 2 date from `start_time` → subtract 1 day → `batch_start_date` (Day 1 date)
2. Search all `MC Enrolled` leads where `MC_Start_Date_Time` date = `batch_start_date` AND `Payment_Status = Completed`
3. Get blueprint transition ID for "MC Completed" from first matched lead
4. For each lead: read `Day_1_Attendance`, `Day_2_Attendance`, `City`, `State`
   - If null → default `Day_1_Attendance = No`, `Day_2_Attendance = No`
   - If City/State blank → `"Please Check With The Participant"`
5. Fire blueprint PUT API to transition each lead
6. Send summary email to `ironladytech@gmail.com`

**Emails sent:**
| Subject | When |
|---------|------|
| `MC Completed - No batch found` | start_time date doesn't match any batch |
| `MC Completed - No leads found` | No MC Enrolled + Payment Completed leads for batch |
| `MC Completed - No transition found` | Blueprint transition unavailable |
| `MC Completed - YYYY-MM-DD (X/Y)` | Summary with success/fail count |

---

### 5. Zoho Sheet (Unmatched Participants Log)
Participants not found in CRM (by email or name) are logged here.

**Columns:** Timestamp, Name, Email, Session, Duration (mins)  
**Connection:** `zoho_sheet_mc`

---

## CRM Fields

| Field API Name | Used By | Description |
|---------------|---------|-------------|
| `Day_1_Attendance_Automation` | `mark_attendance_yes`, `mark_attendance_no` | Auto-marked during Day 1 |
| `Day_2_Attendance_Automation` | `mark_attendance_yes`, `mark_attendance_no` | Auto-marked during Day 2 |
| `Day_1_Attendance` | `mark_mc_completed` (read), Blueprint | RM team field, used in transition |
| `Day_2_Attendance` | `mark_mc_completed` (read), Blueprint | RM team field, used in transition |
| `MC_Start_Date_Time` | `mark_mc_completed` | Used to identify batch |
| `Payment_Status` | `mark_mc_completed` | Must be "Completed" to be included |
| `Lead_Status` | `mark_mc_completed` | Must be "MC Enrolled" to be included |

---

## Attendance Logic Summary

| Scenario | Result |
|----------|--------|
| Participant joins + stays 5+ min | `Yes` marked |
| Participant joins + leaves before 5 min | `No` marked immediately |
| Participant leaves early then rejoins + stays 5 min | `Yes` replaces `No` |
| Participant marked Yes, leaves, rejoins | Stays `Yes` (not changed) |
| Participant joins with no email (guest) | Skipped — not tracked |
| Participant email not in CRM, name matches | Found via name fallback |
| Participant not found by email or name | Logged to Zoho Sheet |
| Day 2 meeting ends | All batch leads → MC Completed |
| Lead has no attendance value at transition | Defaults to `No` |

---

## Deployment

**Bridge (Render):**
- GitHub repo: `IronLadyTech/Attendance`
- Branch: `main`
- Auto-deploys on push
- File: `tools/zoom_webhook_bridge.py`

**To deploy changes:**
```bash
git add tools/zoom_webhook_bridge.py
git commit -m "description"
git push origin main
```

---

## Troubleshooting

| Symptom | Check |
|---------|-------|
| No attendance marked | Render logs — is timer starting? |
| Participant not found | Check email in CRM matches Zoom account email |
| `meeting.ended` not received | Zoom app → Event Subscriptions → "End Meeting" enabled? |
| MC Completed not triggered | Zoho Flow published? `start_time` mapped correctly? |
| All transitions failed | Blueprint field names in `data` match transition form? |
| Email shows `\n` literally | Use `<br>` in sendmail message |
| Bridge not sending events | Check `ZOHO_WEBHOOK_FORWARD_URL` env var on Render |
