# Leave Tracker Automation — MDM BAU Team

Automates leave tracking for the MDM BAU team at Pacific Life. Reads leave emails, 
figures out who's taking leave and when, and updates the Excel tracker automatically.

## What it does

1. **Fetches emails** from an inbox (or reads local test files)
2. **Classifies** the leave type: SL (Sick), CL (Casual), PL (Planned), EL (Emergency), HD (Half Day)
3. **Fuzzy-matches** the sender's name to a team member in the Excel workbook
4. **Writes the leave code** into the correct cell (right person, right date, right month sheet)
5. **Skips** cells that already have WFH, WO, OH, HL, or an existing leave code
6. **Logs everything** to CSV files for auditing
7. **Sends a summary email** to the PM (or prints it if email isn't configured)

## Setup

### 1. Install Python

You need Python 3.8 or newer. Check with:
```
python --version
```

### 2. Install dependencies

```
cd leave_automation
pip install -r requirements.txt
```

That installs:
- `openpyxl` — reads/writes Excel files
- `rapidfuzz` — fuzzy name matching
- `dateparser` — pulls dates out of messy email text
- `msal` — Microsoft auth (only needed if you go the Graph API route)

All free, all open source. No paid APIs or subscriptions.

### 3. Set up your Excel file

The real Excel tracker should have this format:
- One sheet per month: "Jan 2026", "Feb 2026", etc.
- Column A = Stream (team name), Column B = Team Member Name
- Row 1 = day names (Mon, Tue...), Row 2 = day numbers (1, 2, 3...)
- Row 3+ = one row per team member
- Last 4 columns = totals (don't touch these)

For testing, run `python update_excel.py` to generate a test file.

### 4. Configure

Edit `config.json` to set:
- `excel_path` — path to your Excel tracker file
- `fuzzy_threshold` — how close a name match needs to be (0-100, default 75)
- `leave_keywords` — keywords for each leave type (tweak as needed)
- Email settings (see below)

## How to run

### Quick test with sample emails (no real inbox needed)

```
python main.py
```

This reads the `.txt` files in `sample_emails/` and processes them against the test Excel file.

### Process a single email file

```
python main.py --file sick_leave_01.txt
```

### Run against a real inbox (IMAP)

First, set up your email credentials in `config.json`:

```json
{
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "imap_user": "your.email@gmail.com",
    "imap_password": "your-app-password-here"
}
```

**Important:** Use an **App Password**, not your regular password!

- **Gmail:** Go to [Google Account](https://myaccount.google.com/) > Security > 2-Step Verification > App Passwords
- **Outlook:** Go to [Microsoft Account](https://account.microsoft.com/) > Security > App Passwords

Then run:
```
python main.py --imap
```

### Set up the summary email

To email the PM a daily summary, also fill in:

```json
{
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_user": "your.email@gmail.com",
    "smtp_password": "your-app-password-here",
    "pm_email": "manager@company.com"
}
```

If these aren't set, the summary just prints to the console instead.

## File structure

```
leave_automation/
  config.json           ← all settings (keywords, thresholds, email creds)
  main.py               ← run this! ties everything together
  fetch_emails.py       ← gets emails from IMAP or local files
  classify_leave.py     ← figures out leave type and dates from email text
  update_excel.py       ← writes leave codes into the Excel tracker
  send_summary.py       ← emails the daily summary to the PM
  requirements.txt      ← Python dependencies
  sample_emails/        ← test email files (no real inbox needed)
  audit_log.csv         ← log of every action (created at runtime)
  unmatched_log.csv     ← log of names that didn't match (created at runtime)
  test_leave_tracker.xlsx ← test Excel file (created by update_excel.py)
```

## How the classification works

The classifier looks for keywords in the email subject and body:

| Leave Type | Keywords |
|-----------|----------|
| SL (Sick) | sick, fever, unwell, hospital, medical, doctor, not feeling well, health, illness, infection |
| CL (Casual) | casual, personal work, errand, urgent work, out of office, short leave, half day |
| EL (Emergency) | emergency, accident, family emergency, sudden, urgent personal, bereavement, loss, critical |
| PL (Planned) | planned, vacation, holiday, travel, annual leave, pre-approved, approved leave |

**Priority:** If multiple types match, it picks: Emergency > Sick > Planned > Casual

**Half Day:** If "half day", "4 hours", "morning off" etc. are mentioned, it marks "HD" instead.

**Corrections:** If the email says "correction", "update my", etc., it'll overwrite the existing code.

**Unknown:** If nothing matches, it defaults to CL and flags it for review.

## Logs

- **audit_log.csv** — Every email processed, what happened, whether it was updated/skipped/unmatched
- **unmatched_log.csv** — Names that couldn't be fuzzy-matched to anyone in the Excel

## Troubleshooting

**"No match for name"** — The fuzzy threshold might be too high. Try lowering `fuzzy_threshold` in config.json (default is 75, try 65).

**"Sheet not found"** — The month sheet name doesn't exist in the workbook. Sheets should be named like "Jun 2026" (abbreviated month + year).

**IMAP login fails** — Make sure you're using an App Password, not your regular password. Also check that IMAP access is enabled in your email settings.

**dateparser is slow on first run** — That's normal, it loads language data the first time. Subsequent runs are faster.

## What's NOT automated (yet)

- Creating new month sheets (you still need to add those manually)
- Adding new team members to the tracker
- Handling overlapping/conflicting leave requests
- Running on a schedule (you'd need Windows Task Scheduler or a cron job)

## Tech stack (all free)

- Python 3.8+
- openpyxl (Excel read/write)
- rapidfuzz (fuzzy name matching)
- dateparser (date extraction)
- imaplib + email (built-in Python, email fetching)
- smtplib (built-in Python, sending emails)
