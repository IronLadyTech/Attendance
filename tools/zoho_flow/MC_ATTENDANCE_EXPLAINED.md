# MC Attendance — Explained

---

## Overview

When an MC Zoom session runs, the system **automatically updates attendance in Zoho CRM** for paid batch leads. It checks at **two fixed times**:

| Checkpoint | When | Purpose |
|------------|------|---------|
| **First check** | 15 minutes in | Mark who is in the room; flag missing leads as **No** so sales can call |
| **Final check** | 30 minutes in | Final **Yes / No / Absent** for everyone |

Nothing is marked at join time. CRM updates only at **15 min** and **30 min**.

---

## The three attendance values in CRM

| Value | Meaning |
|-------|---------|
| **Yes** | In the Zoom room at the **30-minute** check |
| **No** | Not in the room at 15 min, **or** joined but left before 30 min |
| **Absent** | Paid batch lead who **never joined Zoom** (confirmed at 30 min) |

**Important for sales:** **No at 15 min does not mean absent.** It means *“not in the room right now”* — they may still join before 30 min and become **Yes**.

---

## Quick reference — all situations

| Situation | At 15 min | At 30 min (final) |
|-----------|-----------|-------------------|
| In Zoom the whole time | **Yes** | **Yes** |
| Joins late (after 15 min) but stays till 30 min | **No** | **Yes** |
| Joins, leaves before 15 min | **No** | **No** |
| Joins, leaves before 30 min (after 15 min) | **Yes** or **No** | **No** |
| Never joins Zoom (paid batch) | **No** | **Absent** |

**Note on row 4:** “Yes or No at 15 min” depends on whether the person was **still in Zoom at the 15-minute mark** — not on when they first joined.

---

## One participant — Priya (paid batch lead, Day 1)

Priya’s email is in Zoho CRM. The system tracks her through the session like this:

```
Minute 0          Minute 15              Minute 30
Session starts    First check            Final check
     │                 │                      │
     ▼                 ▼                      ▼
  Tracking         CRM updated            CRM final
  starts           (first result)         (final result)
```

### Example A — Stays the whole session

| Time | What Priya does | CRM (Day 1 Attendance) |
|------|-----------------|--------------------------|
| 0 min | Joins Zoom | *(blank — not marked yet)* |
| 15 min | Still in Zoom | **Yes** |
| 30 min | Still in Zoom | **Yes** |

In the room at both checks → **Yes**.

---

### Example B — Joins late (at 20 min), stays till end

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 0 min | Not joined yet | *(blank)* |
| 15 min | Not in Zoom yet | **No** |
| 20 min | Joins Zoom | Still **No** *(next check is at 30 min)* |
| 30 min | Still in Zoom | **Yes** *(upgraded from No)* |

Not in room at 15 min → **No** (sales can call). Joins in time → **Yes** at 30 min.

---

### Example C — Joins, then leaves at 10 min

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 5 min | Joins Zoom | *(blank)* |
| 10 min | Leaves Zoom | *(blank)* |
| 15 min | Not in Zoom | **No** |
| 30 min | Not in Zoom | **No** *(final)* |

Joined but left before 15 min → **No** at 15 min, stays **No** at 30 min.

---

### Example D — Joins, stays till 20 min, then leaves

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 0 min | Joins Zoom | *(blank)* |
| 15 min | Still in Zoom | **Yes** |
| 20 min | Leaves Zoom | Still **Yes** *(until 30 min check)* |
| 30 min | Not in Zoom | **No** *(downgraded from Yes)* |

In room at 15 min → **Yes**. Left before 30 min → final **No** (dropout).

---

### Example E — Never joins Zoom

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 0–15 min | Never joins | *(blank → then **No**)* |
| 15 min | Not in Zoom | **No** *(sales can call)* |
| 30 min | Still never joined | **Absent** *(final)* |

**No** at 15 min = not in room yet — please call. **Absent** at 30 min = never showed up on Zoom.

---

## Simple rule for one person

1. **15 min** — In Zoom? → **Yes**. Not in Zoom? → **No**.
2. **30 min** — In Zoom? → **Yes**. Joined but left? → **No**. Never joined? → **Absent**.

---

## What RM / sales should do

1. **At 15 min** — Filter CRM for **No**. Call them — they can still join before 30 min.
2. **At 30 min** — Attendance is final:
   - **Yes** = attended
   - **No** = joined but dropped / left early
   - **Absent** = never showed on Zoom
3. **After Day 2 (30 min)** — **MC Completed** is updated automatically for eligible leads.

---

## Day 1 vs Day 2

| Session | Zoom topic (approx.) | CRM field updated |
|---------|----------------------|-------------------|
| **Day 1** | BHAG and Breakthrough Actions | Day 1 Attendance |
| **Day 2** | Art of War, Shameless Pitching & Negotiation | Day 2 Attendance + MC Completed |

---

## How the system works (logic only)

### At session start

- Zoom session begins.
- The system starts tracking who joins, who is still in the room, and who has left.
- **No CRM updates yet.**

---

### At 15 minutes (first check)

The system runs two steps in order:

**Step 1 — Individual updates**

| Who | What CRM gets |
|-----|----------------|
| Currently in Zoom | **Yes** |
| Joined earlier but left before 15 min | **No** *(via batch sweep below)* |
| Not in Zoom yet (paid batch lead) | **No** *(via batch sweep below)* |

**Step 2 — Batch sweep (all paid batch leads for that day)**

- Any lead whose attendance is still **blank** → marked **No**
- This covers people who never joined and people who joined and left early

**Guests not in CRM:** If someone joined Zoom but their email is not in CRM, they are logged to an internal sheet for manual follow-up. CRM leads always get **No** or **Yes** — not left blank.

---

### At 30 minutes (final check)

The system runs three steps in order:

**Step 1 — Individual updates**

| Who | What CRM gets |
|-----|----------------|
| Currently in Zoom | **Yes** *(can upgrade **No → Yes** for late joiners)* |
| Joined at some point but left before 30 min | **No** *(can downgrade **Yes → No** for dropouts)* |

**Step 2 — Batch safety net (all paid batch leads for that day)**

For any lead not yet fully resolved, the system uses join history:

| Join history | Final value |
|--------------|-------------|
| In Zoom at 30 min | **Yes** |
| Joined Zoom at some point but not in room at 30 min | **No** |
| Never joined Zoom at all | **Absent** |

Also:

- **No → Yes** if they are in the room at 30 min (late joiner)
- **No → Absent** if they never joined Zoom at all

**Step 3 — Day 2 only**

After the 30-minute check on Day 2, eligible leads are marked **MC Completed** in CRM.

---

### What the system remembers

Throughout the session, the system keeps track of:

- **Who is in the room right now** — used at each checkpoint
- **Who has ever joined** — used at the 30-minute check to tell “left early” from “never joined”

This is why someone who never opens Zoom ends as **Absent**, while someone who joined and left ends as **No**.

---

### Rules the system follows

| Rule | Behaviour |
|------|-----------|
| Already **Yes** at 30 min, still in room | Stays **Yes** |
| **No** at 15 min, joins and stays till 30 min | Upgraded to **Yes** |
| **Yes** at 15 min, leaves before 30 min | Downgraded to **No** |
| **No** at 15 min, never joins | Upgraded to **Absent** at 30 min |
| Paid batch lead after 30 min | Never left blank — always **Yes**, **No**, or **Absent** |

---

## One-line summary

> **15 min:** In room = Yes; not in room = No (call them).  
> **30 min:** In room = Yes; joined but left = No; never joined = Absent.
