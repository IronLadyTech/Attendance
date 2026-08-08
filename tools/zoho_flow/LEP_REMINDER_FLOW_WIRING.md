# LEP Session Reminder Emails — Zoho Flow Wiring

Scheduled reminders for **IL LEP Sessions** (Gmail via Zoho Flow). Separate from attendance (`LEP_Attendance` flow).

## Schedule (Sat + Sun, both start 9:00 AM IST)

| Reminder | When it sends | Session |
|----------|---------------|---------|
| **1 day before** | Fri **6:00 PM** | Sat Day 1 @ 9 AM |
| **1 day before** | Sat **6:00 PM** | Sun Day 2 @ 9 AM |
| **1 hour before** | Sat **~8:00 AM** | Sat Day 1 @ 9 AM |
| **1 hour before** | Sun **~8:00 AM** | Sun Day 2 @ 9 AM |
| **15 min before** | Sat **~8:45 AM** | Sat Day 1 @ 9 AM |
| **15 min before** | Sun **~8:45 AM** | Sun Day 2 @ 9 AM |

| Reminder | Flow trigger | Window |
|----------|--------------|--------|
| **1 day before** | Daily **6:00 PM IST** | ±30 min (~900 min before 9 AM) |
| **1 hour before** | Every **15 minutes** | 53–67 min before 9 AM (Sat/Sun only) |
| **15 min before** | Same flow as 1 hour | 8–22 min before 9 AM (Sat/Sun only) |

Session start anchor: **9:00 AM IST** on **Saturday (Day 1)** and **Sunday (Day 2)**.

---

## Step 1 — Email template (one template)

Use your premium HTML (`ironlady_session_reminder_premium.html`). Change **3 fixed lines** to placeholders:

| Location | Replace with |
|----------|--------------|
| Preheader (hidden preview, ~line 34) | `{{Reminder_Preheader}}` |
| Eyebrow label (~line 78) | `{{Reminder_Label}}` |
| Hero headline (~line 79) | `{{Reminder_Headline}}` |

Keep existing placeholders: `{{Participant_Name}}`, `{{Session_Date}}`, `{{Session_Time}}` (default: `9:00 AM to 7:00 PM`), `{{Time_Zone}}`, `{{Session_Duration}}` (default: `10hr (including breaks)`), `{{Session_Title}}`, `{{Session_Mode_or_Venue}}`, `{{Session_Join_Link}}`, `{{Support_Email}}`, `{{Current_Year}}`.

**Subject** (set in Deluge, not fixed in template):

| Type | Day 1 example |
|------|----------------|
| 1 day | `Tomorrow: IL LEP Day 1 — Saturday 9:00 AM to 7:00 PM IST` |
| 1 hour | `Starting in 1 hour: IL LEP Day 1` |
| 15 min | `Starting in 15 minutes: IL LEP Day 1` |

Swap **Day 1** / **Saturday** for **Day 2** / **Sunday** as needed.

See `templates/LEP_REMINDER_HERO_SNIPPET.html` for the hero block copy-paste.

---

## Step 2 — CRM fields (Session_Attendance)

**Six checkboxes** — one per reminder type **and** session day (required so Day 2 reminders still send after Day 1):

| Field (UI) | API name | Type |
|------------|----------|------|
| LEP Reminder 1d D1 Sent | `LEP_Reminder_1d_D1_Sent` | Checkbox |
| LEP Reminder 1d D2 Sent | `LEP_Reminder_1d_D2_Sent` | Checkbox |
| LEP Reminder 1h D1 Sent | `LEP_Reminder_1h_D1_Sent` | Checkbox |
| LEP Reminder 1h D2 Sent | `LEP_Reminder_1h_D2_Sent` | Checkbox |
| LEP Reminder 15m D1 Sent | `LEP_Reminder_15m_D1_Sent` | Checkbox |
| LEP Reminder 15m D2 Sent | `LEP_Reminder_15m_D2_Sent` | Checkbox |

If you previously created `LEP_Reminder_1d_Sent` / `1h` / `15m` (3 shared fields), replace with the 6 above — shared fields block Sunday reminders after Saturday sends.

Optional: **LEP Zoom Link** (`LEP_Zoom_Link`) on Session Attendance — else set default in Deluge.

Existing fields used: `Email`, `Lead_Email`, `Name`, `LEP_Start_Date` (Saturday cohort / batch date).

---

## Step 3 — Deluge function

Create **`send_lep_session_reminder`** from `send_lep_session_reminder.deluge`.

| Param | Value |
|-------|--------|
| `reminder_type` | `1day` / `1hour` / `15min` |

Connection: `zoho_flow_to_zoho_crm` (same as attendance).

---

## Step 4 — Two Zoho Flows

### Flow A — `LEP_Reminder_1_Day`

```
Trigger: Schedule → Daily 6:00 PM IST
Action:  send_lep_session_reminder
         reminder_type = 1day
```

### Flow B — `LEP_Reminder_1h_15m`

```
Trigger: Schedule → Every 15 minutes
Action:  send_lep_session_reminder
         reminder_type = 1hour

Action:  send_lep_session_reminder  (second call in same flow, or branch in Deluge)
         reminder_type = 15min
```

**Option:** One scheduled flow every 15 min that calls Deluge once; inside Deluge, run both `1hour` and `15min` windows (already supported if you pass `1hour` and call twice, or extend function).

Simplest: **two actions** in Flow B — call function with `1hour`, then with `15min`.

---

## Step 5 — Gmail

In each Flow action after Deluge, or inside Deluge via `sendmail`:

- **From:** `zoho.loginuserid` (required in Deluge `sendmail`; must be a verified CRM sender)
- **To:** participant `Email` or `Lead_Email`
- **HTML:** CRM template or body built in Deluge

---

## Placeholder text (Deluge sets these)

| reminder_type | Reminder_Headline | Reminder_Label |
|---------------|-------------------|----------------|
| `1day` | Your LEP session is tomorrow. | Session reminder |
| `1hour` | Your session starts in 1 hour. | Starting soon |
| `15min` | Your session is almost here. | Join now |

| reminder_type | Reminder_Preheader |
|---------------|-------------------|
| `1day` | Your IronLady LEP session is tomorrow. Review details and join on time. |
| `1hour` | Your session starts in one hour. |
| `15min` | Your session starts in 15 minutes. |

---

## Test

1. Session Attendance record: `LEP_Start_Date` = upcoming Saturday, email filled.
2. Manually run Flow A or call function with `reminder_type = 1day`.
3. Check inbox + `LEP_Reminder_1d_D1_Sent` = true (or matching D1/D2 field for the session day).

---

## Not used

- Render attendance bridge
- Zoom webhooks
- `LEP_Attendance` flow
