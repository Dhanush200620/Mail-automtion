# send_summary.py
# Sends a daily summary email to the project manager.
# Lists what got updated, what was skipped, and what's still unmatched.
#
# Uses Python's built-in smtplib — no extra libraries needed.
# For Gmail: you need an app password (not your regular password).
# For Outlook: same deal, app password.

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json
import os
from datetime import datetime

# load config
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(script_dir, "config.json")

with open(config_path) as f:
    config = json.load(f)


def build_summary_text(results):
    # takes a list of processing results and builds a readable summary
    # results is a list of dicts like:
    # {"sender": "...", "leave_type": "...", "dates": [...], "status": "...", "message": "..."}
    
    now = datetime.now().strftime("%d %b %Y, %I:%M %p")
    
    lines = []
    lines.append(f"Leave Tracker — Daily Summary")
    lines.append(f"Generated: {now}")
    lines.append(f"{'=' * 50}")
    lines.append("")
    
    # split results into categories
    updated = [r for r in results if r.get("status") == "updated"]
    skipped = [r for r in results if r.get("status") == "skipped"]
    unmatched = [r for r in results if r.get("status") == "unmatched"]
    review = [r for r in results if r.get("needs_review")]
    errors = [r for r in results if r.get("status") == "error"]
    
    # --- successfully updated ---
    lines.append(f"UPDATED ({len(updated)}):")
    if updated:
        for r in updated:
            dates_str = ", ".join(r.get("dates_str", []))
            lines.append(f"  [OK] {r['sender']} - {r['leave_type']} on {dates_str}")
    else:
        lines.append("  (none)")
    lines.append("")
    
    # --- skipped ---
    lines.append(f"SKIPPED ({len(skipped)}):")
    if skipped:
        for r in skipped:
            lines.append(f"  [SKIP] {r['sender']} - {r.get('message', 'skipped')}")
    else:
        lines.append("  (none)")
    lines.append("")
    
    # --- unmatched names ---
    lines.append(f"UNMATCHED ({len(unmatched)}):")
    if unmatched:
        for r in unmatched:
            lines.append(f"  [??] {r['sender']} - could not match to any team member")
    else:
        lines.append("  (none)")
    lines.append("")
    
    # --- needs review ---
    lines.append(f"NEEDS REVIEW ({len(review)}):")
    if review:
        for r in review:
            lines.append(f"  [REVIEW] {r['sender']} - {r.get('message', 'flagged for review')}")
    else:
        lines.append("  (none)")
    lines.append("")
    
    # --- errors ---
    if errors:
        lines.append(f"ERRORS ({len(errors)}):")
        for r in errors:
            lines.append(f"  [ERR] {r['sender']} - {r.get('message', 'unknown error')}")
        lines.append("")
    
    lines.append(f"{'=' * 50}")
    lines.append("This is an automated email from the Leave Tracker.")
    
    return "\n".join(lines)


def send_summary_email(results):
    # sends the summary to the PM via SMTP
    # if SMTP creds aren't configured, just prints it to console instead
    
    summary_text = build_summary_text(results)
    
    smtp_server = config.get("smtp_server", "")
    smtp_port = config.get("smtp_port", 587)
    smtp_user = config.get("smtp_user", "")
    smtp_password = config.get("smtp_password", "")
    pm_email = config.get("pm_email", "")
    
    # if credentials aren't set, just print the summary
    if not smtp_server or not smtp_user or not smtp_password or not pm_email:
        print("\n[Summary Email] SMTP not configured — printing to console instead:")
        print("-" * 50)
        print(summary_text)
        print("-" * 50)
        return False
    
    try:
        # build the email
        msg = MIMEMultipart()
        msg["From"] = smtp_user
        msg["To"] = pm_email
        msg["Subject"] = f"Leave Tracker Summary — {datetime.now().strftime('%d %b %Y')}"
        msg.attach(MIMEText(summary_text, "plain"))
        
        # connect and send
        print(f"Sending summary email to {pm_email}...")
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, pm_email, msg.as_string())
        server.quit()
        
        print("Summary email sent!")
        return True
        
    except Exception as e:
        print(f"ERROR sending email: {e}")
        print("Check SMTP settings in config.json")
        return False


# ============================================================
# Test it with some dummy results
# ============================================================

if __name__ == "__main__":
    
    print("=" * 60)
    print("  send_summary.py — Test Run")
    print("=" * 60)
    
    # fake some results to test the summary formatting
    test_results = [
        {
            "sender": "Dhanush Kumar",
            "leave_type": "SL",
            "dates_str": ["23 Jun"],
            "status": "updated",
            "needs_review": False,
        },
        {
            "sender": "Priya Sharma",
            "leave_type": "CL",
            "dates_str": ["25 Jun"],
            "status": "updated",
            "needs_review": False,
        },
        {
            "sender": "Kavita Nair",
            "leave_type": "PL",
            "dates_str": ["29 Jun", "30 Jun"],
            "status": "skipped",
            "message": "cells already have leave codes",
            "needs_review": False,
        },
        {
            "sender": "John Doe",
            "leave_type": "CL",
            "dates_str": [],
            "status": "unmatched",
            "needs_review": False,
        },
        {
            "sender": "Neha Patel",
            "leave_type": "SL",
            "dates_str": ["23 Jun"],
            "status": "updated",
            "needs_review": True,
            "message": "correction email, multiple leave types matched (SL, PL)",
        },
    ]
    
    # this will print to console since SMTP isn't configured
    send_summary_email(test_results)
    
    print("\n" + "=" * 60)
    print("  Done!")
    print("=" * 60)
