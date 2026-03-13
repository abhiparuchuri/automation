# Google Sheets → Gmail Email Scheduler

Sends semi-personalized outreach emails from Gmail based on a Google Sheet. Scrapes role names from job links, fills a template, and sends when `send_at` is due.

---

## Quick Start

1. **Sheet**: Tab named `Emails123`, header row with columns below
2. **Script**: Extensions → Apps Script → paste `email-scheduler-script.js`
3. **Trigger**: Triggers → Add → `runMorningSend`, Time-driven, Day timer, 8am–9am

---

## Sheet Columns

| Column | Notes |
|--------|--------|
| `email` | Required |
| `first_name` | Used in greeting |
| `company` | In body |
| `job_link` | Scraped for role name |
| `send_at` | DateTime cell (when to send) |
| `role_name` | Auto-filled by scrape, or manual |
| `status` | PENDING / SENDING / SENT / SKIPPED |
| `skip_reason` | Why skipped |
| `attempt_count` | Retry tracking |
| `last_attempt_at` | Timestamp |
| `sent_at` | Timestamp when sent |

---

## Flow

- Add rows with `status = PENDING`, future `send_at`
- Script runs (trigger or manual): sends rows where `send_at ≤ now` and role name is present
- Skips if `role_name` can’t be scraped; you can manually fill it and set `status = PENDING` to retry

---

## Skipped Emails

Filter by `status = SKIPPED`. Common `skip_reason`:

- `MISSING_EMAIL` / `MISSING_URL`
- `ROLE_NOT_FOUND` (couldn’t parse title)
- `FETCH_FAILED` / `HTTP_403`, `HTTP_500`, etc.

---

## Customize

| Change | Edit in script |
|--------|----------------|
| Subject | `DEFAULT_SUBJECT` |
| Body | `renderBody_` (plain), `renderHtmlBody_` (HTML) |
| Send cap | `MAX_SEND_PER_RUN` |
| Sheet tab | `SHEET_NAME` |

---

## Safety

- Uses `MailApp` (send-only, no delete/read)
- Capped at 80 sends per run
- Retries transient scrape failures (10m, 1h, 6h backoff)
