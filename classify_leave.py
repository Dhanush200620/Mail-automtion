# classify_leave.py
# This script takes an email (subject + body + sender) and figures out:
#   - what type of leave it is (SL, CL, PL, EL)
#   - what dates the leave covers
#   - whether it's a half day
#   - whether it's a correction to a previous request
#   - whether it needs manual review (couldn't classify confidently)
#
# It uses keyword matching from config.json and dateparser to pull dates
# out of messy email text.

import re
import json
import os
from datetime import datetime, date, timedelta
import dateparser

# load config
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(script_dir, "config.json")

with open(config_path) as f:
    config = json.load(f)

LEAVE_KEYWORDS = config["leave_keywords"]
HALF_DAY_KEYWORDS = config["half_day_keywords"]
LEAVE_PRIORITY = config["leave_priority"]  # EL > SL > PL > CL


def parse_email_file(filepath):
    # reads a .txt email file and splits it into subject, from, date, body
    # the format is simple: header lines at the top, then a blank line, then body
    
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()
    
    lines = content.split("\n")
    
    subject = ""
    sender = ""
    email_date = ""
    body_start = 0
    
    # parse the headers (everything before the first blank line)
    for i, line in enumerate(lines):
        if line.strip() == "":
            body_start = i + 1
            break
        
        if line.lower().startswith("subject:"):
            subject = line.split(":", 1)[1].strip()
        elif line.lower().startswith("from:"):
            sender = line.split(":", 1)[1].strip()
        elif line.lower().startswith("date:"):
            email_date = line.split(":", 1)[1].strip()
    
    body = "\n".join(lines[body_start:]).strip()
    
    # extract just the name from the "From" field
    # it usually looks like "Dhanush Kumar <dhanush.kumar@company.com>"
    sender_name = extract_name(sender)
    
    # parse the email date so we know "today" and "tomorrow" relative to when it was sent
    parsed_email_date = None
    if email_date:
        # dateparser is pretty good at parsing weird date formats
        parsed_email_date = dateparser.parse(email_date)
    
    return {
        "subject": subject,
        "sender_raw": sender,
        "sender_name": sender_name,
        "email_date": parsed_email_date,
        "body": body,
        "full_text": subject + " " + body  # combined for keyword searching
    }


def extract_name(from_field):
    # pulls the human name out of an email From field
    # "Dhanush Kumar <dhanush.kumar@company.com>" -> "Dhanush Kumar"
    # "dhanush.kumar@company.com" -> "Dhanush Kumar" (from the email address)
    
    from_field = from_field.strip()
    
    # check if there's a name before the <email> part
    match = re.match(r'^(.*?)\s*<.*?>$', from_field)
    if match:
        name = match.group(1).strip()
        if name:
            return name
    
    # if no name part, try to build one from the email address
    # like dhanush.kumar@company.com -> Dhanush Kumar
    email_match = re.search(r'([\w.]+)@', from_field)
    if email_match:
        username = email_match.group(1)
        # replace dots and underscores with spaces, then title case
        name = username.replace(".", " ").replace("_", " ").title()
        return name
    
    # couldn't figure it out, just return whatever we got
    return from_field


def classify_leave_type(text):
    # looks for leave-type keywords in the text (subject + body combined)
    # returns the leave type based on priority: EL > SL > PL > CL
    # if nothing matches, defaults to CL and flags for review
    
    text_lower = text.lower()
    
    matched_types = []
    
    for leave_type, keywords in LEAVE_KEYWORDS.items():
        for keyword in keywords:
            # for multi-word keywords, plain substring match is fine
            # for single words, use word boundaries so "loss" doesn't match "across" etc.
            if " " in keyword:
                found = keyword.lower() in text_lower
            else:
                found = bool(re.search(r'\b' + re.escape(keyword.lower()) + r'\b', text_lower))
            
            if found:
                matched_types.append(leave_type)
                print(f"    keyword hit: '{keyword}' -> {leave_type}")
                break  # one hit per leave type is enough
    
    if not matched_types:
        # nothing matched — default to CL but flag it
        print("    no keywords matched, defaulting to CL (needs review)")
        return "CL", True  # True = needs_review
    
    # if multiple types matched, use priority order
    for priority_type in LEAVE_PRIORITY:
        if priority_type in matched_types:
            needs_review = len(matched_types) > 1  # multiple matches = maybe review
            return priority_type, needs_review
    
    # shouldn't get here but just in case
    return matched_types[0], False


def check_half_day(text):
    # checks if the email mentions a half day
    text_lower = text.lower()
    
    for keyword in HALF_DAY_KEYWORDS:
        if keyword.lower() in text_lower:
            print(f"    half-day detected: '{keyword}'")
            return True
    
    return False


def check_correction(text):
    # checks if this email is correcting a previous leave request
    # looking for words like "correction", "update", "change", "modify"
    
    text_lower = text.lower()
    correction_keywords = ["correction", "correct my", "update my", "change my",
                           "modify my", "was actually", "not sick", "please update",
                           "wrong leave", "incorrect"]
    
    for keyword in correction_keywords:
        if keyword in text_lower:
            print(f"    correction detected: '{keyword}'")
            return True
    
    return False


def extract_dates(text, reference_date=None):
    # tries to pull leave dates out of the email text
    # this is the trickiest part — emails say stuff like "tomorrow", "June 23 to 25",
    # "next Monday", etc. dateparser handles most of it but not everything
    
    # TODO: this is still kinda rough, need to handle more date formats
    #       like "23rd and 24th" or "next week Mon-Wed"
    
    dates = []
    ref = reference_date or datetime.now()
    
    # settings for dateparser — tells it what "today" means
    # NOT using STRICT_PARSING because it rejects too many valid dates
    parser_settings = {
        "RELATIVE_BASE": ref,
        "PREFER_DATES_FROM": "current_period",
        "PREFER_DAY_OF_MONTH": "first",
    }
    
    # helper to add a date without duplicates
    def add_date(d):
        if isinstance(d, datetime):
            d = d.date()
        if d not in dates:
            dates.append(d)
            return True
        return False
    
    # --- Step 1: try to find date ranges like "June 29 to July 2" ---
    # these patterns try to catch various ways people write date ranges
    range_patterns = [
        # "from June 29 to July 2" or "from June 29 to July 2, 2026"
        r'from\s+(.+?)\s+to\s+(.+?)(?:\.|,\s*\d{4}|\s*$)',
        # "June 29 to July 2" (month+day to month+day)
        r'((?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:,?\s+\d{4})?)\s*(?:to|-)\s*((?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:,?\s+\d{4})?)',
        # "Jul 6 to Jul 10" (abbreviated month)
        r'((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2})\s*(?:to|-)\s*((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}(?:,?\s*\d{4})?)',
    ]
    
    for pattern in range_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            start_str, end_str = match[0].strip(), match[1].strip()
            
            start_date = dateparser.parse(start_str, settings=parser_settings)
            end_date = dateparser.parse(end_str, settings=parser_settings)
            
            if start_date and end_date:
                # if end is before start, the year might be wrong — try next year
                if end_date < start_date:
                    end_date = end_date.replace(year=end_date.year + 1)
                
                # generate all dates in the range
                current = start_date.date()
                end = end_date.date()
                
                while current <= end:
                    add_date(current)
                    current += timedelta(days=1)
                
                print(f"    date range found: {start_date.strftime('%d %b')} to {end_date.strftime('%d %b')}")
                dates.sort()
                return dates  # found a range, that's probably the main dates
    
    # --- Step 2: check for "today" and "tomorrow" ---
    text_lower = text.lower()
    
    if "today" in text_lower:
        d = ref.date() if isinstance(ref, datetime) else ref
        if add_date(d):
            print(f"    found 'today': {d.strftime('%d %b %Y')}")
    
    if "tomorrow" in text_lower:
        d = (ref + timedelta(days=1))
        d = d.date() if isinstance(d, datetime) else d
        if add_date(d):
            print(f"    found 'tomorrow': {d.strftime('%d %b %Y')}")
    
    # --- Step 3: look for individual date mentions ---
    # like "June 23", "23 June 2026", "June 23, 2026"
    date_patterns = [
        r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:,?\s+\d{4})?',
        r'\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)(?:\s+\d{4})?',
        r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}(?:,?\s+\d{4})?',
    ]
    
    for pattern in date_patterns:
        found = re.findall(pattern, text, re.IGNORECASE)
        for date_str in found:
            parsed = dateparser.parse(date_str, settings=parser_settings)
            if parsed:
                d = parsed.date()
                if add_date(d):
                    print(f"    found date: '{date_str}' -> {d.strftime('%d %b %Y')}")
    
    # --- Step 4: check for "and" between dates ---
    # like "June 23 and June 24"
    and_pattern = r'(\w+\s+\d{1,2})\s+and\s+(\w+\s+\d{1,2})'
    and_matches = re.findall(and_pattern, text, re.IGNORECASE)
    for match in and_matches:
        for date_str in match:
            parsed = dateparser.parse(date_str.strip(), settings=parser_settings)
            if parsed:
                add_date(parsed.date())
    
    if not dates:
        print("    WARNING: couldn't find any dates in the email!")
    
    dates.sort()
    return dates


def classify_email(filepath):
    # main function — takes an email file, returns everything we figured out
    
    print(f"\n  Parsing: {os.path.basename(filepath)}")
    
    # step 1: parse the email file
    email = parse_email_file(filepath)
    print(f"    From: {email['sender_name']}")
    print(f"    Subject: {email['subject']}")
    
    # step 2: classify the leave type
    print("    Classifying leave type...")
    leave_type, needs_review = classify_leave_type(email["full_text"])
    
    # step 3: check if it's a half day
    is_half_day = check_half_day(email["full_text"])
    if is_half_day:
        leave_type = "HD"
    
    # step 4: check if it's a correction
    is_correction = check_correction(email["full_text"])
    
    # step 5: extract dates
    print("    Extracting dates...")
    leave_dates = extract_dates(email["full_text"], email["email_date"])
    
    result = {
        "sender_name": email["sender_name"],
        "subject": email["subject"],
        "leave_type": leave_type,
        "leave_dates": leave_dates,
        "is_half_day": is_half_day,
        "is_correction": is_correction,
        "needs_review": needs_review,
        "reason": email["subject"],  # using subject as the reason for now
    }
    
    print(f"    -> Type: {leave_type}, Dates: {[d.strftime('%d %b') for d in leave_dates]}, "
          f"Correction: {is_correction}, Review: {needs_review}")
    
    return result


# ============================================================
# Test it against the sample emails
# ============================================================

if __name__ == "__main__":
    
    samples_dir = os.path.join(script_dir, config["sample_emails_folder"])
    
    print("=" * 60)
    print("  classify_leave.py — Test Run")
    print("=" * 60)
    
    if not os.path.exists(samples_dir):
        print(f"ERROR: sample emails folder not found: {samples_dir}")
        exit(1)
    
    # get all .txt files in the sample folder
    email_files = sorted([f for f in os.listdir(samples_dir) if f.endswith(".txt")])
    
    if not email_files:
        print("No sample emails found!")
        exit(1)
    
    print(f"\nFound {len(email_files)} sample emails\n")
    
    all_results = []
    
    for filename in email_files:
        filepath = os.path.join(samples_dir, filename)
        print("-" * 50)
        result = classify_email(filepath)
        all_results.append(result)
    
    # print summary table
    print("\n" + "=" * 60)
    print("  SUMMARY")
    print("=" * 60)
    print(f"{'Name':<20} {'Type':<6} {'Dates':<25} {'Correction':<12} {'Review'}")
    print("-" * 80)
    
    for r in all_results:
        dates_str = ", ".join([d.strftime("%d %b") for d in r["leave_dates"]]) if r["leave_dates"] else "NO DATES"
        print(f"{r['sender_name']:<20} {r['leave_type']:<6} {dates_str:<25} "
              f"{str(r['is_correction']):<12} {r['needs_review']}")
    
    print("\n" + "=" * 60)
    print("  Done!")
    print("=" * 60)
