# LEP Checkpoint Attendance — Zoho Flow Wiring

**IL LEP Sessions - 1 & 2 Days** — majority of 3 in-room checks → **Present** / **Absent**.

Wire a separate **LEP_Attendance** Flow (not Mc_Attendance or 100BM_Attendance).

## Matching rule (important)

**Email or name only** — no batch date, no `LEP_Start_Date`, no Lead Status filter.

Previous-batch participants often join the current session; marking is only for people **actually in the Zoom meeting** at checkpoints. The bridge sends one update per participant who joined at least once.

## Zoom app (LEP account)

Enable event subscriptions:

- `meeting.started`
- `meeting.participant_joined`
- `meeting.participant_left`
- `meeting.ended` (optional — bridge cancels timers only; no CRM action)

Webhook endpoint: bridge **`POST /lep`** with `ZOOM_WEBHOOK_SECRET_TOKEN_LEP`.

## Recommended Zoom topics

| Session | Example topic |
|---------|-----------------|
| Day 1 | `IL LEP Sessions - 1 & 2 Days - Day 1` |
| Day 2 | `IL LEP Sessions - 1 & 2 Days - Day 2` |

If topic has no `Day 2` hint, bridge treats it as **Day 1**.

## Checkpoint schedule (9:00 AM IST anchor)

| Day | Check 1 | Check 2 | Check 3 | Final (majority) |
|-----|---------|---------|---------|------------------|
| Day 1 | 9:15 AM | 3:30 PM | 6:15 PM | 6:30 PM |
| Day 2 | 9:15 AM | 12:30 PM | 4:15 PM | 4:30 PM |

At each check: **Present** = in Zoom room; **Absent** = not in room.

At final: **2+ Present → Present**; tie or **2+ Absent → Absent**.

## Decision branches

| Condition | Event | Action |
|-----------|-------|--------|
| condition1 | `attendance.lep_final` | `mark_lep_attendance_final` |
| Default | `meeting.ended`, etc. | *(no action)* |

## CRM fields

| UI label | API name | Values |
|----------|----------|--------|
| LEP Day 1 Session | `LEP_Day_1_Session` | Present / Absent |
| LEP Day 2 Session | `LEP_Day_2_Session` | Present / Absent |

## Parameter mapping — mark_lep_attendance_final

| Param | Value |
|-------|--------|
| participant_email | `${webhookTrigger.payload.participant_email}` |
| participant_name | `${webhookTrigger.payload.participant_name}` |
| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |
| session_date | `${webhookTrigger.payload.session_date}` |
| session_day | `${webhookTrigger.payload.session_day}` |
| attendance_result | `${webhookTrigger.payload.attendance_result}` |

## Render / bridge env

| Variable | Purpose |
|----------|---------|
| `ZOOM_WEBHOOK_SECRET_TOKEN_LEP` | Zoom secret for `/lep` |
| `ZOHO_WEBHOOK_FORWARD_URL_LEP` | LEP Zoho Flow webhook URL |

## No Day 2 completion

Unlike MC, LEP does **not** run a blueprint / “Completed” transition when the meeting ends.
