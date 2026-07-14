# cold-mail-sender

Automated, personalized HR outreach: reads a contact database, asks Gemini
(rotating across multiple models and API keys to dodge rate limits) for a
short personalized opening line per contact, sends via SMTP with a resume
attached, and throttles itself to avoid spam flags.

Runs daily via GitHub Actions (`.github/workflows/send-emails.yml`). Live
status dashboard: https://rishittandon7.github.io/cold-mail-sender/

## How it works

1. `send_hr_emails.py` loads contacts from an xlsx (fetched at runtime from
   a GitHub Release asset — the contact database and resume aren't committed
   to the repo).
2. For each contact, it calls Gemini for a one-line personalized opener,
   round-robining across `GEMINI_MODELS` and up to 16 `GOOGLE_API_KEY_*`
   secrets so no single key/model gets rate-limited.
3. Sends via Gmail SMTP with the resume attached.
4. Logs every send (with timestamp) to `sent_log.json` so reruns never
   double-email anyone, and writes `stats.json` for the dashboard.
5. Copies both into `docs/data/` so GitHub Pages can render live status.

## Configuration (repo Settings → Secrets and variables → Actions)

**Secrets:** `SMTP_USER`, `SMTP_PASS` (Gmail App Password), `GOOGLE_API_KEY_1`
through `GOOGLE_API_KEY_16` (as many as you have — rotation uses whichever
are set).

**Variables:** `DRY_RUN` — `true` (default) prints sample emails without
sending; set to `false` when ready to go live. Can also be overridden per
manual run via the "Run workflow" dispatch input. The GitHub Pages dashboard
also supports manual runs from the browser using a PAT with `Contents: write`;
its mode toggle is tab-local and does not change the repository variable.

## Local dry run

```
pip install openpyxl requests
python send_hr_emails.py
```
