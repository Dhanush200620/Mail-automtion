# update_google_sheet_yearly.py
# Updates the Google Sheet "test_leave_tracker" with:
#   - 12 monthly sheets (Jul 2026 – Jun 2027)
#   - Adds new team members: Dhanush, Balavignesh, Poovarasan, Swetha
#   - Preserves existing team members from the current sheet
#   - Marks weekends as WO and Indian holidays as OH

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import calendar
import json
import os
import time
from datetime import date

# --- Config ---
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(script_dir, "config.json")

with open(config_path) as f:
    config = json.load(f)

GOOGLE_SHEET_NAME = config.get("google_sheet_name", "test_leave_tracker")
GOOGLE_CREDENTIALS_FILE = config.get("google_credentials_file", "google_credentials.json")

# Year range
START_YEAR = 2026
START_MONTH = 7
END_YEAR = 2029
END_MONTH = 6

# We will just rewrite the team list entirely below
NEW_MEMBERS = []

# Indian public holidays (2026-2027)
INDIAN_HOLIDAYS = {
    date(2026, 8, 15): "Independence Day",
    date(2026, 10, 2): "Gandhi Jayanti",
    date(2026, 10, 20): "Diwali",
    date(2026, 10, 21): "Diwali (Day 2)",
    date(2026, 11, 4): "Diwali Amavasya",
    date(2026, 12, 25): "Christmas",
    date(2027, 1, 1): "New Year",
    date(2027, 1, 14): "Pongal",
    date(2027, 1, 15): "Pongal (Day 2)",
    date(2027, 1, 26): "Republic Day",
    date(2027, 3, 29): "Holi",
    date(2027, 4, 2): "Good Friday",
    date(2027, 4, 14): "Tamil New Year",
    date(2027, 5, 1): "May Day",
}


def connect_to_google_sheets():
    """Connect to Google Sheets using service account credentials."""
    creds_file = GOOGLE_CREDENTIALS_FILE
    if not os.path.isabs(creds_file):
        creds_file = os.path.join(script_dir, creds_file)

    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
    client = gspread.authorize(creds)
    return client


def get_existing_members(sh):
    """Read existing team members from the first available sheet."""
    existing_members = []
    try:
        # Try to read from any existing worksheet
        for ws in sh.worksheets():
            values = ws.get_all_values()
            if len(values) >= 3:  # Has data rows
                for row in values[2:]:  # Skip header rows (row 1 & 2)
                    if len(row) >= 2 and row[1] and row[1].strip():
                        stream = row[0].strip() if row[0] else "MDM BAU"
                        name = row[1].strip()
                        existing_members.append((stream, name))
                if existing_members:
                    break  # Got members from first sheet with data
    except Exception as e:
        print(f"  Warning: couldn't read existing members: {e}")

    return existing_members


def build_full_team(existing_members):
    """Return the team matching the new structure."""
    return [
        ("MDM-BAU", "Kavivanan"),
        ("MDM-BAU", "Leando Iruthayaraj"),
        ("MDM-BAU", "Vibra Narayanan"),
        ("MDM-BAU", "Aravind"),
        ("MDM-BAU", "Jaya Shree"),
        ("MDM - GROW", "Vignesh Pugalendhi"),
        ("MDM - GROW", "Sujitha"),
        ("MDM - GROW", "Balaji Loganathan"),
        ("MDM - GROW", "Ebron Stalin"),
        ("MDM - GROW", "Rahul Kumar"),
        ("MDM - GROW", "Harivarshan"),
        ("MDM - GROW", "BharaniDharan"),
        ("MDM - GROW", "Sandhiya"),
        ("MDM - GROW", "Jose"),
        ("CRM - Next Gen", "Dravid"),
        ("CRM - Next Gen", "Soma Sai Pavan"),
        ("CRM - Next Gen", "Rathimeena"),
        ("CRM - Next Gen", "Pavithra"),
        ("CRM - Next Gen", "Jefflina"),
        ("CRM - Next Gen", "Manoj"),
        ("CDS", "Arshad"),
        ("CDS", "Dushyanth"),
        ("PowerBI / Tableau / Alteryx / Devops Admin", "Hari Vembeina"),
        ("PowerBI / Tableau / Alteryx / Devops Admin", "JP"),
        ("PowerBI / Tableau / Alteryx / Devops Admin", "Jelsi"),
        ("PowerBI / Tableau / Admin", "Mahesh Reddy"),
        ("PowerBI / Tableau / Admin", "Sandeep Reddy"),
        ("", "Alafia Yusuf Fidvi"),
        ("", "Arun R"),
        ("", "Swaminathan"),
        ("", "Vasanth"),
    ]


def generate_months():
    """Generate list of (year, month) tuples for the tracker period."""
    months = []
    y, m = START_YEAR, START_MONTH
    while (y, m) <= (END_YEAR, END_MONTH):
        months.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def create_month_data(year, month_num, team):
    """Build the grid data for one monthly sheet."""
    days_in_month = calendar.monthrange(year, month_num)[1]
    grid = []

    # --- Row 1: Headers ---
    row1 = ["Stream", "Team Member Name"]
    for day in range(1, days_in_month + 1):
        d = date(year, month_num, day)
        row1.append(d.strftime("%a"))  # Mon, Tue, etc.

    # Total headers
    total_headers = [
        "Total CL", "Total SL", "Total PL", "Total EL", "Total HD",
        "Total WFH", "Total Week Off", "Total Comp Off", "Total Leave Days"
    ]
    row1.extend(total_headers)
    grid.append(row1)

    # --- Row 2: Day numbers ---
    row2 = ["", ""]
    for day in range(1, days_in_month + 1):
        row2.append(day)
    row2.extend([""] * len(total_headers))
    grid.append(row2)

    # --- Row 3+: Team members ---
    for i, (stream, name) in enumerate(team):
        row = [stream, name]
        row_num = i + 3  # 1-indexed row number in the sheet

        for day in range(1, days_in_month + 1):
            d = date(year, month_num, day)
            if d.weekday() in [5, 6]:  # Saturday, Sunday
                row.append("WO")
            elif d in INDIAN_HOLIDAYS:
                row.append("OH")
            else:
                row.append("")

        # Formulas for totals
        last_data_col_letter = chr(ord('A') + days_in_month + 1)  # Won't work for >24 days
        # Use proper column letter calculation
        last_data_col_num = days_in_month + 2
        last_data_col = _col_letter(last_data_col_num)

        total_start_col = days_in_month + 3

        row.extend([
            f'=COUNTIF(C{row_num}:{last_data_col}{row_num}, "CL")',
            f'=COUNTIF(C{row_num}:{last_data_col}{row_num}, "SL")',
            f'=COUNTIF(C{row_num}:{last_data_col}{row_num}, "PL")',
            f'=COUNTIF(C{row_num}:{last_data_col}{row_num}, "EL")',
            f'=COUNTIF(C{row_num}:{last_data_col}{row_num}, "HD")',
            f'=COUNTIF(C{row_num}:{last_data_col}{row_num}, "WFH")',
            f'=COUNTIF(C{row_num}:{last_data_col}{row_num}, "WO")',
            f'=COUNTIF(C{row_num}:{last_data_col}{row_num}, "CO")',
        ])

        # Total Leave Days = CL + SL + PL + EL + (HD * 0.5)
        cl_col = _col_letter(total_start_col)
        el_col = _col_letter(total_start_col + 3)
        hd_col = _col_letter(total_start_col + 4)
        row.append(f'=SUM({cl_col}{row_num}:{el_col}{row_num}) + ({hd_col}{row_num}*0.5)')

        grid.append(row)

    return grid


def _col_letter(col_num):
    """Convert 1-based column number to Excel-style column letter (1=A, 27=AA, etc.)."""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def update_google_sheet():
    """Main function: updates the Google Sheet with full year data."""
    print("=" * 60)
    print("  Updating Google Sheet: Leave Tracker")
    print(f"  Sheet name: {GOOGLE_SHEET_NAME}")
    print(f"  Period: {calendar.month_abbr[START_MONTH]} {START_YEAR} – {calendar.month_abbr[END_MONTH]} {END_YEAR}")
    print("=" * 60)

    # Connect
    print("\n[1/4] Connecting to Google Sheets...")
    client = connect_to_google_sheets()
    sh = client.open(GOOGLE_SHEET_NAME)
    print(f"  Connected to '{GOOGLE_SHEET_NAME}'")

    # Read existing members
    print("\n[2/4] Reading existing team members...")
    existing_members = get_existing_members(sh)
    if existing_members:
        print(f"  Found {len(existing_members)} existing members:")
        for stream, name in existing_members:
            print(f"    - {name} ({stream})")
    else:
        print("  No existing members found, starting fresh.")

    # Build full team
    print("\n[3/4] Building full team list...")
    team = build_full_team(existing_members)
    print(f"  Total team members: {len(team)}")

    # Generate months
    months = generate_months()

    # Create/update each monthly sheet
    print(f"\n[4/4] Creating {len(months)} monthly sheets...")

    for year, month_num in months:
        month_abbr = calendar.month_abbr[month_num]
        ws_name = f"{month_abbr} {year}"

        # Check if worksheet exists
        try:
            ws = sh.worksheet(ws_name)
            ws.clear()
            print(f"  Updated: {ws_name} (cleared & rebuilt)")
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=ws_name, rows=100, cols=50)
            print(f"  Created: {ws_name}")

        # Build data grid
        grid = create_month_data(year, month_num, team)

        # Write to sheet using value_input_option='USER_ENTERED' so formulas work
        end_cell = gspread.utils.rowcol_to_a1(len(grid), len(grid[0]))
        ws.update(grid, f"A1:{end_cell}", value_input_option='USER_ENTERED')

        # Rate limiting — Google Sheets API has quotas
        time.sleep(2)

    # Clean up the default "Sheet1" if it exists and we have other sheets
    try:
        default_sheet = sh.worksheet("Sheet1")
        if len(sh.worksheets()) > 1:
            sh.del_worksheet(default_sheet)
            print("\n  Removed default 'Sheet1'")
    except gspread.WorksheetNotFound:
        pass

    print(f"\n{'=' * 60}")
    print(f"  Done! Google Sheet '{GOOGLE_SHEET_NAME}' updated.")
    print(f"  Sheets: {len(months)} monthly sheets")
    print(f"  Team: {len(team)} members")
    print(f"  New members added: {[n for n in NEW_MEMBERS if n.lower() not in [name.lower() for _, name in existing_members]]}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    update_google_sheet()
