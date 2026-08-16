# LEP Attendance — Explained

---

## Overview

LEP runs across a **weekend**: Saturday is **Day 1**, Sunday is **Day 2**. The same Zoom links are used on both days — the system works out which day it is from the **calendar date**, not the meeting title.

Up to **four Zoom rooms** run in parallel. All four are **combined into one list** before anything is written to Zoho CRM, so a participant counts the same whichever room she is in.

Each day is sampled at **three checkpoints**, then a **final** result is worked out from those three.

| | Check 1 | Check 2 | Check 3 | Final |
|---|---|---|---|---|
| **Day 1 (Saturday)** | 9:15 AM | 3:30 PM | 6:15 PM | 6:30 PM |
| **Day 2 (Sunday)** | 9:15 AM | 12:30 PM | 4:15 PM | 4:30 PM |

All times IST, measured from a 9:00 AM anchor. Nothing is marked when someone joins — CRM is updated **only at these four moments**.

---

## The two attendance values in CRM

| Value | Meaning |
|-------|---------|
| **Present** | In one of the Zoom rooms at that moment |
| **Absent** | Not in any Zoom room at that moment |

There is no third value. Someone who never joined and someone who joined and left both read as **Absent** — the difference is visible in the checkpoint history, not in the CRM field.

---

## The final result is a majority vote

This is the most important rule and the biggest difference from MC.

**A person is Present for the day if she was in a room at 2 of the 3 checkpoints.**

| Check 1 | Check 2 | Check 3 | Final |
|---------|---------|---------|-------|
| Present | Present | Present | **Present** |
| Present | Present | Absent | **Present** |
| Present | Absent | Present | **Present** |
| Absent | Present | Present | **Present** |
| Present | Absent | Absent | **Absent** |
| Absent | Present | Absent | **Absent** |
| Absent | Absent | Present | **Absent** |
| Absent | Absent | Absent | **Absent** |

Missing one checkpoint does not make someone Absent:

| Check 1 | Check 2 | Check 3 | Final |
|---------|---------|---------|-------|
| Present | Present | *(never ran)* | **Present** |
| Absent | Absent | *(never ran)* | **Absent** |
| Present | Absent | *(never ran)* | **Unresolved — CRM not changed** |

**Unresolved means the system refuses to guess.** With only one Present and one Absent and no third vote, it leaves the record alone rather than marking someone Absent on incomplete evidence. If you see Unresolved in the logs, that record needs a human decision.

---

## Which CRM field gets updated

**Module:** `Session_Attendance`

| Day | Field |
|-----|-------|
| Saturday | `LEP_Day_1_Session` |
| Sunday | `LEP_Day_2_Session` |

**The cohort is always the Saturday date.** A record belongs to the weekend if its `LEP Start Date` matches that Saturday — including on Sunday. Sunday's checkpoints look up the Saturday date, not their own.

---

## What each checkpoint actually writes

Every checkpoint writes to **every record in the cohort**, in two steps:

1. **Everyone in a room** → marked **Present**, one record at a time
2. **Everyone else in the cohort** → marked **Absent**, in one batch

So an attendance value is never left blank once a checkpoint has run.

---

## Why the CRM timeline often looks empty

Zoho only records a timeline entry when a value **changes**. Check 2 writes `Present` over `Present` for anyone who stayed in the room, so nothing appears.

**A quiet timeline means nobody's status changed — not that the check failed.** You will only see entries for people who crossed over: joined late (Absent → Present) or left early (Present → Absent).

The per-checkpoint history is kept separately and is not visible in the CRM record. Ask for it if you need to see how someone was scored at each individual check.

---

## How a person is matched to a CRM record

Most participants join as guests, so Zoom sends **no email address** — only a display name. The system tries these in order:

| Order | Match on | Example |
|-------|----------|---------|
| 1 | Email exactly | `priya@example.com` → same email in CRM |
| 2 | Full name exactly | `Priya Sharma` → `Priya Sharma` |
| 3 | CRM name contains the Zoom name | `Priya` → `Priya Sharma` |
| 4 | First name, **only if one person in the cohort has it** | `Priya` → the only Priya |
| 5 | Two or more people share that first name | **Not updated** — logged for review |

Step 5 is deliberate. Marking the wrong person present is worse than marking nobody, so the system stops and records it instead.

Anyone who matches nothing is written to the **unmatched sheet** with their name, the check number, and the batch date.

---

## One participant — Priya (Day 2, Sunday)

### Example A — Attends the whole day

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 9:15 AM | In the room | **Present** |
| 12:30 PM | Still in the room | **Present** *(no timeline entry — unchanged)* |
| 4:15 PM | Still in the room | **Present** |
| 4:30 PM | Final: Present / Present / Present | **Present** |

### Example B — Joins late, stays

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 9:15 AM | Not joined yet | **Absent** |
| 12:30 PM | In the room | **Present** *(timeline shows the change)* |
| 4:15 PM | Still in the room | **Present** |
| 4:30 PM | Final: Absent / Present / Present → majority Present | **Present** |

Missing the first check does not cost her the day.

### Example C — Leaves at lunch

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 9:15 AM | In the room | **Present** |
| 12:30 PM | Left | **Absent** |
| 4:15 PM | Still away | **Absent** |
| 4:30 PM | Final: Present / Absent / Absent → majority Absent | **Absent** |

### Example D — Steps out briefly

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 9:15 AM | In the room | **Present** |
| 12:30 PM | Out for 20 minutes | **Absent** |
| 4:15 PM | Back in the room | **Present** |
| 4:30 PM | Final: Present / Absent / Present → majority Present | **Present** |

One missed check does not lose the day. That is the point of sampling three times.

### Example E — Never joins

| Time | What Priya does | CRM |
|------|-----------------|-----|
| All three checks | Never joins | **Absent** at each |
| 4:30 PM | Final: Absent / Absent / Absent | **Absent** |

### Example F — Moves between rooms

| Time | What Priya does | CRM |
|------|-----------------|-----|
| 9:15 AM | In room 1 | **Present** |
| 12:30 PM | Moved to room 3 | **Present** |

All four rooms are combined first, so moving rooms changes nothing.

---

## Simple rule

> **At each check:** in a room = Present, not in a room = Absent.
> **For the day:** Present at 2 of the 3 checks = Present.

---

## What the team should do

**During the day** — nothing. Attendance settles itself at each checkpoint.

**After the final** (6:30 PM Saturday, 4:30 PM Sunday):

1. The Day 1 / Day 2 column is complete for the whole cohort
2. Check the **unmatched sheet** — anyone there attended but could not be tied to a CRM record
3. Look for **Unresolved** in the logs — those need a human decision

**Before the next batch** — make sure every participant's name in CRM matches the name she uses in Zoom. That single thing prevents most problems below.

---

## Known limits

**Renaming in Zoom splits a person in two.** Someone who joins as `Namita Dutt` and later rejoins as `Namita's` becomes two separate people, each with half the attendance, and both may end up Absent. Ask participants to join with their registered name.

**Spelling differences cannot be bridged.** CRM `Sushmita` and Zoom `Sushmitha P` will not match, and the system will not guess. Fix it in one place or the other.

**Shared logins count once.** Two people on one Zoom account are one participant, and only one CRM record can be credited.

**No checkpoint means no attendance.** If a session's schedule fails to be created, rooms are still tracked but nothing is written to CRM. The bridge now warns loudly at startup and on every join when this can happen — if you see that warning, fix it before the session starts.

---

## One-line summary

> **Four rooms combined. Sampled three times a day. Present at two of three = Present for the day. A missing check is never counted as absent.**
