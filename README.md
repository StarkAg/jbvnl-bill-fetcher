# JBVNL Bill Auto-Fetcher

Automatically logs into the [JBVNL customer portal](https://cportal.jbvnl.co.in),
downloads the **current month's** electricity bill PDF for each of your accounts,
files it by property/month, and logs every month's bill **amount** to a spreadsheet.

Built with Playwright. Your credentials never live in the code — they're read at
runtime from a git-ignored `.env`.

## Why current-month only
The portal only serves a valid PDF for the latest bill. Older months return an empty
(0-byte) response *from the server* (verified at the byte level). No tool can fetch
what the server doesn't send — so this is a **preventive** pipeline: run it monthly and
you'll never miss a bill again. (The bill *amounts* for older months are still scraped
from the table, even when their PDFs aren't available.)

## Setup
```bash
python -m venv venv && source venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
cp .env.example .env      # then fill in JBVNL_PASSWORD, BILLS_BASE, JBVNL_ACCOUNTS
```

`.env` example:
```
JBVNL_PASSWORD=your-password
BILLS_BASE=/path/to/your/Electricity Bills
JBVNL_ACCOUNTS=100000000001:Home,100000000002:Shop
```

## Run
```bash
source venv/bin/activate
HEADED=1 python fetch_bills.py          # watch the first run
python fetch_bills.py                    # normal (headless) run, all accounts
python fetch_bills.py --account 100000000001   # a single account
```

Bills are saved to `BILLS_BASE/<account> (<Label>)/M:YY.pdf`. Amounts are written to a
**"Bill Amounts"** sheet inside `BILLS_BASE/Electricity.xlsx` (existing sheets left
untouched), de-duplicated by (account, billing date).

## Automate monthly (macOS launchd)
Edit `com.stark.jbvnl-billfetch.plist.template` — replace `__PROJECT_DIR__` with the
absolute path to this folder — save it to
`~/Library/LaunchAgents/com.stark.jbvnl-billfetch.plist`, then:
```bash
launchctl load ~/Library/LaunchAgents/com.stark.jbvnl-billfetch.plist
```
Runs on the 8th of each month at 10:00 (bills issue ~5th). On Linux, use cron.

## Telegram notifications (optional)
Each run posts a summary via `notify.py`, reading `TELEGRAM_BOT_TOKEN` /
`TELEGRAM_CHAT_ID` from the environment. No token is stored in the repo.

## Security
Passwords and account numbers live only in your local `.env` (git-ignored). Nothing
personal is committed. Rotate your password anytime — just update `.env`.
