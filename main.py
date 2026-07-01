# main.py
# This is the main script that ties everything together.
# Run this to do a full pass: fetch emails -> classify -> update Excel -> log -> summarize
#
# Usage:
#   python main.py              — fetch from local sample_emails/ folder
#   python main.py --imap       — fetch from a real IMAP inbox
#   python main.py --file sick_leave_01.txt  — process a single file
#
# It writes results to audit_log.csv and unmatched_log.csv

import os
import sys
import csv
import json
import argparse
import time
from datetime import datetime

# import our modules
from fetch_emails import fetch_local_emails, fetch_imap_emails
from classify_leave import classify_email, parse_email_file, classify_leave_type, \
    check_half_day, check_correction, extract_dates, extract_name
from update_excel import write_leave, write_leave_multi_day
from send_summary import send_summary_email

# load config
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(script_dir, "config.json")

with open(config_path) as f:
    config = json.load(f)

EXCEL_PATH = os.path.join(script_dir, config["excel_path"])
AUDIT_LOG = os.path.join(script_dir, "audit_log.csv")
UNMATCHED_LOG = os.path.join(script_dir, "unmatched_log.csv")


def init_logs():
    # create the CSV log files with headers if they don't exist yet
    
    if not os.path.exists(AUDIT_LOG):
        with open(AUDIT_LOG, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "sender", "leave_type", "dates", "status",
                             "message", "needs_review", "is_correction", "source_file"])
        print(f"Created {AUDIT_LOG}")
    
    if not os.path.exists(UNMATCHED_LOG):
        with open(UNMATCHED_LOG, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "sender", "subject", "source_file", "reason"])
        print(f"Created {UNMATCHED_LOG}")


def log_to_audit(sender, leave_type, dates, status, message, needs_review, is_correction, source_file):
    # append one row to audit_log.csv
    
    dates_str = ", ".join([d.strftime("%d %b %Y") for d in dates]) if dates else "none"
    
    with open(AUDIT_LOG, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            sender,
            leave_type,
            dates_str,
            status,
            message,
            needs_review,
            is_correction,
            source_file,
        ])


def log_to_unmatched(sender, subject, source_file, reason):
    # append one row to unmatched_log.csv
    
    with open(UNMATCHED_LOG, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            sender,
            subject,
            source_file,
            reason,
        ])


def process_email_file(filepath):
    # takes a single email file, classifies it, and writes to Excel
    # returns a result dict for the summary
    
    print(f"\n{'=' * 50}")
    print(f"Processing: {os.path.basename(filepath)}")
    print(f"{'=' * 50}")
    
    # step 1: classify the email
    classification = classify_email(filepath)
    
    sender = classification["sender_name"]
    leave_type = classification["leave_type"]
    leave_dates = classification["leave_dates"]
    is_correction = classification["is_correction"]
    needs_review = classification["needs_review"]
    source_file = os.path.basename(filepath)
    
    result = {
        "sender": sender,
        "leave_type": leave_type,
        "dates_str": [d.strftime("%d %b") for d in leave_dates],
        "status": "unknown",
        "message": "",
        "needs_review": needs_review,
    }
    
    # step 2: check if we got any dates
    if not leave_dates:
        print(f"\n  No dates found — can't update Excel")
        result["status"] = "error"
        result["message"] = "No dates found in email"
        log_to_audit(sender, leave_type, leave_dates, "error",
                     "No dates found", needs_review, is_correction, source_file)
        return result
    
    # step 3: write to Excel for each date
    print(f"\n  Writing to Excel...")
    all_success = True
    messages = []
    
    for leave_date in leave_dates:
        success, msg = write_leave(EXCEL_PATH, sender, leave_date, leave_type, is_correction)
        
        if not success:
            all_success = False
            messages.append(msg)
            
            # check if it's an unmatched name issue
            if "No match" in msg:
                log_to_unmatched(sender, classification["subject"], source_file, msg)
                result["status"] = "unmatched"
                result["message"] = msg
                log_to_audit(sender, leave_type, leave_dates, "unmatched",
                             msg, needs_review, is_correction, source_file)
                return result  # no point trying other dates if name doesn't match
    
    # figure out the overall status
    if all_success:
        result["status"] = "updated"
        result["message"] = f"Updated {len(leave_dates)} day(s)"
    else:
        result["status"] = "skipped"
        result["message"] = "; ".join(set(messages))
    
    # log it
    log_to_audit(sender, leave_type, leave_dates, result["status"],
                 result["message"], needs_review, is_correction, source_file)
    
    return result


def run_local_mode(specific_file=None):
    # process emails from the local sample_emails/ folder
    
    samples_dir = os.path.join(script_dir, config["sample_emails_folder"])
    
    if specific_file:
        # process just one file
        filepath = os.path.join(samples_dir, specific_file)
        if not os.path.exists(filepath):
            # maybe it's a full path already
            filepath = specific_file
        if not os.path.exists(filepath):
            print(f"ERROR: file not found: {specific_file}")
            return []
        return [process_email_file(filepath)]
    
    # process all .txt files in the folder
    if not os.path.exists(samples_dir):
        print(f"ERROR: sample emails folder not found: {samples_dir}")
        return []
    
    email_files = sorted([f for f in os.listdir(samples_dir) if f.endswith(".txt")])
    
    if not email_files:
        print("No email files found!")
        return []
    
    print(f"\nFound {len(email_files)} email files to process\n")
    
    results = []
    for filename in email_files:
        filepath = os.path.join(samples_dir, filename)
        result = process_email_file(filepath)
        results.append(result)
    
    return results


def run_imap_mode():
    # fetch and process emails from a real IMAP inbox
    # TODO: this is untested since I don't have IMAP creds set up yet
    
    print("Fetching emails via IMAP...")
    emails = fetch_imap_emails(mark_as_read=True)
    
    if not emails:
        print("No new leave emails found")
        return []
    
    # save fetched emails to temp files so classify_email can read them
    # (classify_email expects file paths right now)
    # TODO: maybe refactor classify_email to accept a dict directly?
    
    temp_dir = os.path.join(script_dir, "fetched_emails")
    os.makedirs(temp_dir, exist_ok=True)
    
    results = []
    for email_data in emails:
        # save to a temp file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{timestamp}.txt"
        filepath = os.path.join(temp_dir, filename)
        
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"Subject: {email_data['subject']}\n")
            f.write(f"From: {email_data['sender']}\n")
            f.write(f"Date: {email_data['date']}\n")
            f.write(f"\n{email_data['body']}\n")
        
        result = process_email_file(filepath)
        results.append(result)
    
    return results


def run_once(mode, specific_file):
    # run the appropriate mode
    if mode == "imap":
        print("\n>>> Running in IMAP mode <<<")
        results = run_imap_mode()
    else:
        if specific_file:
            print(f"\n>>> Processing single file: {specific_file} <<<")
        else:
            print("\n>>> Running in LOCAL mode (sample emails) <<<")
        results = run_local_mode(specific_file)
    
    # print a quick summary
    if results:
        print("\n" + "=" * 60)
        print("  PROCESSING COMPLETE")
        print("=" * 60)
        
        updated = sum(1 for r in results if r["status"] == "updated")
        skipped = sum(1 for r in results if r["status"] == "skipped")
        unmatched = sum(1 for r in results if r["status"] == "unmatched")
        errors = sum(1 for r in results if r["status"] == "error")
        review = sum(1 for r in results if r["needs_review"])
        
        print(f"  Updated:    {updated}")
        print(f"  Skipped:    {skipped}")
        print(f"  Unmatched:  {unmatched}")
        print(f"  Errors:     {errors}")
        print(f"  Need Review:{review}")
        print(f"\n  Audit log: {AUDIT_LOG}")
        print(f"  Unmatched log: {UNMATCHED_LOG}")
        
        # send summary email (or print to console if SMTP not configured)
        send_summary_email(results)
    else:
        print("\nNo new emails or files to process.")


# ============================================================
# Main entry point
# ============================================================

if __name__ == "__main__":
    
    print("=" * 60)
    print("  Leave Tracker — MDM BAU Team")
    print("  Pacific Life — Automated Leave Processing")
    print("=" * 60)
    print(f"  Time: {datetime.now().strftime('%d %b %Y, %I:%M %p')}")
    print(f"  Excel: {EXCEL_PATH}")
    print("=" * 60)
    
    # initialize log files
    init_logs()
    
    # get default interval from config, or fallback to 60 seconds
    default_interval = config.get("poll_interval_seconds", 60)
    
    # parse arguments
    parser = argparse.ArgumentParser(description="Leave Tracker Automation — MDM BAU Team")
    parser.add_argument("specific_file_pos", nargs="?", help="Specific email file to process")
    parser.add_argument("--imap", action="store_true", help="Run in IMAP mode (fetch from inbox)")
    parser.add_argument("--file", help="Specific email file to process (alternative format)")
    parser.add_argument("--loop", "-l", action="store_true", help="Run continuously in a loop")
    parser.add_argument("--interval", "-i", type=int, default=default_interval,
                        help="Polling interval in seconds for loop mode")
    
    args = parser.parse_args()
    
    # determine mode and specific file
    mode = "imap" if args.imap else "local"
    specific_file = args.specific_file_pos or args.file
    
    if args.loop:
        print("=" * 60)
        print(f"  Running in LOOP mode (polling every {args.interval} seconds)")
        print("  Press Ctrl+C to stop the process.")
        print("=" * 60)
        
        try:
            while True:
                print(f"\n--- Starting Polling Cycle: {datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')} ---")
                run_once(mode, specific_file)
                print(f"\nSleeping for {args.interval} seconds... (Ctrl+C to stop)")
                time.sleep(args.interval)
        except KeyboardInterrupt:
            print("\nLoop execution interrupted by user. Exiting cleanly...")
    else:
        run_once(mode, specific_file)
        print("\nDone!")
