# Starter: Zoom webhook → Zoho Flow → Zoho CRM attendance

This matches your target flow: **join → webhook → match Event + attendee → update attendance**. Your Streamlit app is optional for batch/manual backup.

---

## Phase A — Zoho CRM (do this first)

1. **Events module**
   - Add a text field **Zoom Meeting ID** (API name e.g. `Zoom_Meeting_ID`) so Flow can find the event from the webhook payload.
   - Add attendance fields you need: picklist **Attendance status**, optional **Join time** / **Duration**, or an **“Needs review”** checkbox for bad email matches.

2. **Many participants**
   - If more than one join, prefer a **related list** or custom **Attendance** module (one row per Event + Contact), instead of one field on the Event.

3. **Zoho–Zoom integration** (recommended)
   - Enable if available so meetings created from Zoho already store the Zoom ID on the Event — matching becomes reliable.

---

## Phase B — Zoho Flow

1. Create a **new Flow**.
2. **Trigger** (pick one):
   - **Zoom** app trigger for meeting/participant events (if your Zoho Flow edition includes it), **or**
   - **Incoming webhook** / **Catch webhook** — Flow gives you a **URL**; use that as the Zoom webhook endpoint.
3. **Actions** (outline):
   - Parse payload: `meeting id`, participant **email** / **name**, timestamps if present.
   - **Search CRM Event** where `Zoom_Meeting_ID` equals payload meeting id.
   - **Search Lead or Contact** by email (if email missing → branch: set Event “Needs review” or create task).
   - **Update** Event or related Attendance record (status, join time, duration rules you define).
4. Turn the Flow **ON**. Run a **test** from Flow if available.

---

## Phase C — Zoom

1. **Zoom Pro or higher** (webhook availability follows Zoom’s current rules).
2. **Zoom Marketplace** → build/configure an app that supports **webhooks**.
3. Subscribe to the event you need (often named around **participant joined** / meeting participation — confirm exact names in Zoom’s webhook docs for your app type).
4. **Endpoint URL** = the webhook URL from **Zoho Flow** (not this repo).
5. Complete Zoom’s **validation** step; Zoho Flow’s Zoom connector usually documents how validation works — follow Zoho’s guide if using generic webhook.

---

## Phase D — Debug payloads locally (optional)

Use this repo’s helper only to **see** what Zoom sends **before** switching the endpoint to Zoho:

```powershell
cd D:\Attendence\Attendance
python tools/zoom_webhook_echo.py --port 8765
```

For Zoom to reach your PC, use a tunnel (e.g. ngrok) and paste that HTTPS URL into Zoom’s webhook config **temporarily**. Stop the tunnel when done.

---

## What stays in this Python project

| Tool | Role |
|------|------|
| `app.py` | Manual UI: CSV / Zoom report → Google Sheet |
| `auto_sync_attendance.py` | Scheduled batch sync (no webhooks) |
| Real-time CRM attendance | **Zoho Flow + Zoom**, not these scripts |

---

## Zoom “Validate URL” fails with Zoho Flow webhook

Zoom sends `endpoint.url_validation` and expects a JSON body with `plainToken` + `encryptedToken` (HMAC-SHA256 using your app’s **Secret Token**). A plain Zoho incoming webhook often **cannot** answer that, so validation fails.

**Fix:** Use Zoho’s **Zoom app trigger** in Flow if your account includes it, **or** run the bridge in this repo:

- Copy `env.zoom-zoho.example` → `.env` and fill in variables (or set the same names in the shell).
- `ZOOM_WEBHOOK_SECRET_TOKEN` = Zoom app → Feature → **Secret Token** (webhooks).
- `ZOHO_WEBHOOK_FORWARD_URL` = your full Zoho incoming webhook URL.
- Deploy **`tools/zoom_webhook_bridge.py`** behind **HTTPS**, put **that public URL** in Zoom (not Zoho’s URL). The bridge validates Zoom and **forwards** real events to Zoho.

Zoom requires a public **HTTPS** URL with a valid certificate — `localhost` only works via a tunnel (e.g. ngrok) for testing.

### Deploy bridge on Railway (HTTPS without ngrok)

1. Push this repo to **GitHub** (or connect the folder).
2. [Railway](https://railway.app) → **New project** → **Deploy from GitHub** → pick the repo (or upload).
3. Railway detects **`Dockerfile`** and builds `zoom_webhook_bridge` only (small image).
4. **Variables** (project → Variables):
   - `ZOOM_WEBHOOK_SECRET_TOKEN` — Zoom Secret Token  
   - `ZOHO_WEBHOOK_FORWARD_URL` — full Zoho Flow webhook URL  
   (Do **not** rely on `.env` in the repo on Railway.)
5. **Deploy** → open the generated **HTTPS** domain (Railway Settings → Networking → generate domain).
6. Paste **`https://your-app.up.railway.app`** (your real URL) into Zoom **Event notification endpoint URL**, then **Validate**.

`railway.toml` sets a GET **healthcheck** on `/` (the bridge returns 200).

---

## Next concrete action

Start with **Phase A** (CRM fields + Zoom ID on Events), then **Phase B** (one Flow with webhook trigger), then **Phase C** (Zoom points at Flow URL).
