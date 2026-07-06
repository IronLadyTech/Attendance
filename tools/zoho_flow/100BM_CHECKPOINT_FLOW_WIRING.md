# 100BM Checkpoint Attendance — Zoho Flow Wiring



After deploying the updated bridge, wire **100BM_Attendance** Flow (separate from Mc_Attendance).



## Zoom app (100BM account)



Enable event subscriptions:



- `meeting.started`

- `meeting.participant_joined`

- `meeting.participant_left`

- `meeting.ended` (optional — no attendance action)



Webhook endpoint: bridge **`POST /100bm`** with `ZOOM_WEBHOOK_SECRET_TOKEN_100BM`.



## Decision branches (required)



| Condition | Event | Action |

|-----------|-------|--------|

| condition1 | `attendance.mark_yes` | `mark_100bm_attendance_yes` |

| condition2 | `attendance.lookup` | `mark_100bm_attendance_lookup` |

| Default | everything else | *(no action)* |



## Decision branches (optional — audit log only)



| Condition | Event | Action |

|-----------|-------|--------|

| condition3 | `attendance.first_check` | `mark_100bm_batch_attendance_checkpoint` (check_type = **first**) |

| condition4 | `attendance.final_check` | `mark_100bm_batch_attendance_checkpoint` (check_type = **final**) |



The batch function **does not write CRM**. It only logs how many eligible leads are Yes vs still blank. You can omit these branches entirely.



**Remove or disable (old timer model):**



- `attendance.mark_no` branch

- `attendance.update_duration` branch



## CRM fields



| UI label | API name | Written when |

|----------|----------|--------------|

| FT Attendance | `FT_attendance` | **Yes only** — present at T+15 or T+30 |

| FT Attended Date | `FT_attended_date` | **Yes only** — session date (or current date fallback) |

| FT Invite Date | `FT_Invite_Date` | Eligibility filter only (not updated) |



**Policy:** No / Absent are **not** written. Non-attendees keep `FT_attendance` and `FT_attended_date` **blank/untouched**.



## Eligibility



Participants are processed when **`FT_Invite_Date` matches the session date** (IST date from meeting `start_time`). No `Payment_Status` filter.



## Parameter mapping



### mark_100bm_attendance_yes



| Param | Value |

|-------|--------|

| meeting_id | `${webhookTrigger.payload.meeting_id}` |

| participant_email | `${webhookTrigger.payload.participant_email}` |

| participant_name | `${webhookTrigger.payload.participant_name}` |

| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |

| session_date | `${webhookTrigger.payload.session_date}` |



### mark_100bm_attendance_lookup



| Param | Value |

|-------|--------|

| participant_email | `${webhookTrigger.payload.participant_email}` |

| participant_name | `${webhookTrigger.payload.participant_name}` |

| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |

| session_date | `${webhookTrigger.payload.session_date}` |



### mark_100bm_batch_attendance_checkpoint (optional audit)



| Param | Value |

|-------|--------|

| start_time | `${webhookTrigger.payload.start_time}` |

| meeting_topic | `${webhookTrigger.payload.meeting_topic}` |

| session_date | `${webhookTrigger.payload.session_date}` |

| check_type | `first` or `final` (literal) |



## Attendance outcomes



| Situation | CRM `FT_attendance` | CRM `FT_attended_date` |

|-----------|---------------------|------------------------|

| In meeting at T+15 or T+30 | **Yes** | Session date |

| Joined but left before T+15 | **Blank** (untouched) | **Blank** |

| Never joined (eligible lead) | **Blank** (untouched) | **Blank** |

| Guest (no CRM lead) | — | Unmatched sheet only |



## Timeline



```

T+0     meeting.started (bridge schedules T+15 / T+30)

T+15    mark_yes (each in room) → Yes + FT_attended_date

        lookup (each who joined then left) → sheet for guests only

T+30    mark_yes (each in room) → Yes + FT_attended_date (if rejoined)

End     meeting.ended (optional; no CRM action)

```



## Bridge env vars



| Variable | Default | Purpose |

|----------|---------|---------|

| `BM100_CHECKPOINT_1_SECONDS` | 900 | T+15 |

| `BM100_CHECKPOINT_2_SECONDS` | 1800 | T+30 |

| `ZOHO_WEBHOOK_FORWARD_URL_100BM` | — | Separate 100BM Flow URL |


