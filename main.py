"""Main CLI entrypoint and orchestrator for PIB Daily UPSC Automation."""

import argparse
import sys
import os
import subprocess
import time
from datetime import datetime, date, timedelta
from pathlib import Path

# Reconfigure stdout for Windows terminal Unicode compatibility
sys.stdout.reconfigure(encoding="utf-8")

from config import OUTPUT_DIR
from pib_scraper import PIBScraper
from upsc_classifier import UPSCClassifier
from doc_builder import DocumentBuilder


def run_pipeline(target_date: date = None, open_files: bool = False):
    """Executes the complete scrape -> classify -> document build pipeline."""
    target = target_date if target_date else date.today()
    date_code = target.strftime("%Y-%m-%d")
    
    print("\n" + "=" * 60)
    print(f"  PIB UPSC AUTOMATION ENGINE: {date_code}")
    print("=" * 60)

    # 1. Scrape
    scraper = PIBScraper()
    releases, formatted_date = scraper.fetch_releases(target_date=target)
    
    # Smart Morning Fallback:
    # If checking today's releases and none are published yet (common before 11:00 AM IST),
    # automatically fallback to yesterday's complete daily bulletin.
    is_fallback = False
    if not releases and target == date.today():
        yesterday = target - timedelta(days=1)
        print(f"[!] No releases found for today ({formatted_date}) yet.")
        print(f"[*] Activating Smart Morning Fallback -> Compiling complete yesterday's bulletin ({yesterday.strftime('%d %B %Y')})...")
        target = yesterday
        date_code = target.strftime("%Y-%m-%d")
        releases, formatted_date = scraper.fetch_releases(target_date=target)
        is_fallback = True

    if not releases:
        print(f"[!] No releases found for {formatted_date}.")
        return None, None, None

    # 2. Classify
    classifier = UPSCClassifier()
    enriched = classifier.process_all(releases)
    
    relevant_count = sum(1 for r in enriched if r.get("is_relevant"))
    print(f"[+] Total Releases: {len(enriched)} | UPSC Relevant: {relevant_count} | Routine Filtered: {len(enriched) - relevant_count}")
    if is_fallback:
        print(f"[*] Morning Bulletin: Prepared using yesterday's complete wrap-up ({date_code}).")

    # 3. Generate Documents
    excel_filename = f"PIB_UPSC_Daily_{date_code}.xlsx"
    docx_filename = f"PIB_UPSC_Daily_{date_code}.docx"
    
    excel_path = str(OUTPUT_DIR / excel_filename)
    docx_path = str(OUTPUT_DIR / docx_filename)

    DocumentBuilder.generate_excel(enriched, excel_path, formatted_date)
    DocumentBuilder.generate_docx(enriched, docx_path, formatted_date)

    print(f"[✓] Generated Excel: {excel_path}")
    print(f"[✓] Generated Word:  {docx_path}")
    print("=" * 60 + "\n")

    if open_files and sys.platform == "win32":
        try:
            os.startfile(docx_path)
        except Exception:
            pass

    return excel_path, docx_path, date_code


def install_windows_task(run_time: str = "21:00"):
    """Registers a daily task in Windows Task Scheduler using schtasks.exe."""
    task_name = "PIB_UPSC_Daily_Automation"
    python_exe = sys.executable
    script_path = str(Path(__file__).resolve())
    
    command = f'"{python_exe}" "{script_path}" --today'
    schtasks_cmd = [
        "schtasks", "/Create",
        "/SC", "DAILY",
        "/TN", task_name,
        "/TR", command,
        "/ST", run_time,
        "/F"  # force overwrite
    ]

    print(f"Creating Windows Scheduled Task: '{task_name}' at {run_time} daily...")
    res = subprocess.run(schtasks_cmd, capture_output=True, text=True)
    if res.returncode == 0:
        # Configure battery and wake settings using PowerShell
        ps_config = (
            f'$t = Get-ScheduledTask -TaskName "{task_name}"; '
            '$s = $t.Settings; '
            '$s.DisallowStartIfOnBatteries = $false; '
            '$s.StopIfGoingOnBatteries = $false; '
            '$s.StartWhenAvailable = $true; '
            '$s.WakeToRun = $true; '
            f'Set-ScheduledTask -TaskName "{task_name}" -Settings $s'
        )
        subprocess.run(["powershell", "-Command", ps_config], capture_output=True)
        print(f"[✓] Task successfully created! It will run automatically every day at {run_time}.")
        print("[✓] Configured: Runs on battery, wakes from sleep, and catches up if PC was off.")
    else:
        print(f"[!] Error creating scheduled task: {res.stderr}")


def run_scheduler_loop(run_time: str = "21:00"):
    """Simple in-process daemon that sleeps until target time every day."""
    print(f"[*] Starting local scheduler daemon. Job will run every day at {run_time} (24-hr format).")
    print("[*] Press Ctrl+C to stop.")
    
    while True:
        now = datetime.now()
        current_time_str = now.strftime("%H:%M")
        if current_time_str == run_time:
            print(f"[*] Scheduled time ({run_time}) reached. Triggering pipeline...")
            run_pipeline(target_date=date.today())
            print("[*] Sleeping for 65 seconds to prevent re-triggering...")
            time.sleep(65)
        time.sleep(20)


def main():
    parser = argparse.ArgumentParser(description="PIB UPSC Daily Headline Automation")
    parser.add_argument("--today", action="store_true", help="Run extraction for today")
    parser.add_argument("--date", type=str, help="Run extraction for specific date (YYYY-MM-DD)")
    parser.add_argument("--open", action="store_true", help="Automatically open generated DOCX on completion")
    parser.add_argument("--install-task", type=str, nargs="?", const="21:00", help="Install Windows Scheduled Task (default 21:00)")
    parser.add_argument("--daemon", type=str, nargs="?", const="21:00", help="Run local scheduler loop at time HH:MM")
    parser.add_argument("--email", action="store_true", help="Dispatch generated files via email")
    parser.add_argument("--to", type=str, help="Comma-separated recipient email address(es)")

    args = parser.parse_args()

    if args.install_task:
        install_windows_task(args.install_task)
        return

    if args.daemon:
        run_scheduler_loop(args.daemon)
        return

    target = None
    if args.date:
        try:
            target = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            print("[!] Invalid date format. Use YYYY-MM-DD (e.g. 2026-10-04).")
            return
    
    excel_path, docx_path, result_date_code = run_pipeline(target_date=target, open_files=args.open)

    if (args.email or args.to) and (excel_path or docx_path):
        from send_email import send_daily_brief_email
        target_str = result_date_code if result_date_code else (target.strftime("%Y-%m-%d") if target else None)
        send_daily_brief_email(target_date=target_str, to_override=args.to)


if __name__ == "__main__":
    main()
