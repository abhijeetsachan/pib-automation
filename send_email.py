import os
import sys
import smtplib
import socket
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import date, timedelta

# Ensure safe console output encoding on Windows
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Force IPv4 socket resolution to prevent dropped connections on Azure/GitHub Actions runners
_orig_getaddrinfo = socket.getaddrinfo
def _getaddrinfo_ipv4(host, port, family=0, type=0, proto=0, flags=0):
    res = _orig_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)
    return res if res else _orig_getaddrinfo(host, port, family, type, proto, flags)
socket.getaddrinfo = _getaddrinfo_ipv4

# Ensure paths
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"


def load_env_file():
    """Loads environment variables from .env file if it exists."""
    env_file = BASE_DIR / ".env"
    if env_file.exists():
        with open(env_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    key = key.strip()
                    val = val.strip().strip("'\"")
                    if key not in os.environ:
                        os.environ[key] = val

def generate_email_html(releases: list, date_str: str) -> str:
    """Generates a rich, styled HTML email body listing scraped PIB articles with separate hyperlinked URLs."""
    relevant = [r for r in releases if r.get("is_relevant")]
    routine = [r for r in releases if not r.get("is_relevant")]

    paper_order = ["GS-1", "GS-2", "GS-3", "GS-4"]
    badge_colors = {
        "GS-1": ("#FEF3C7", "#92400E"),
        "GS-2": ("#DBEAFE", "#1E40AF"),
        "GS-3": ("#DCFCE7", "#166534"),
        "GS-4": ("#F3E8FF", "#6B21A8"),
    }

    relevant_html = ""
    if relevant:
        grouped = {p: [] for p in paper_order}
        for r in relevant:
            p = r.get("paper", "GS-3")
            if p not in grouped:
                grouped[p] = []
            grouped[p].append(r)

        for p in paper_order:
            items = grouped.get(p, [])
            if not items:
                continue
            bg_c, text_c = badge_colors.get(p, ("#E2E8F0", "#1E293B"))
            relevant_html += f"""
            <div style="margin-top: 18px; margin-bottom: 8px;">
              <span style="background-color: {bg_c}; color: {text_c}; font-weight: bold; font-size: 13px; padding: 4px 10px; border-radius: 4px; display: inline-block;">
                {p} Articles ({len(items)})
              </span>
            </div>
            """
            for idx, item in enumerate(items, 1):
                headline = str(item.get("headline", "")).replace("<", "&lt;").replace(">", "&gt;")
                ministry = str(item.get("ministry", "")).replace("<", "&lt;").replace(">", "&gt;")
                url = item.get("url", "")
                reason = str(item.get("relevance_reason", "")).replace("<", "&lt;").replace(">", "&gt;")
                tags = ", ".join(item.get("tags", []))

                relevant_html += f"""
                <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #2563EB; border-radius: 6px; padding: 12px 16px; margin-bottom: 12px;">
                  <div style="font-size: 14.5px; font-weight: bold; color: #0F172A; line-height: 1.4;">
                    {idx}. <a href="{url}" target="_blank" style="color: #1E3A8A; text-decoration: none;">{headline}</a>
                  </div>
                  <div style="font-size: 12px; color: #64748B; margin-top: 4px;">
                    <b>Ministry:</b> {ministry} {f'| <b>Keywords:</b> {tags}' if tags else ''}
                  </div>
                  {f'<div style="font-size: 12px; color: #1E40AF; margin-top: 3px; font-weight: 500;"><b>UPSC Alignment:</b> {reason}</div>' if reason else ''}
                  <div style="font-size: 12px; margin-top: 6px; padding-top: 6px; border-top: 1px dashed #E2E8F0; word-break: break-all;">
                    🔗 <b>Official PIB URL:</b> <a href="{url}" target="_blank" style="color: #0284C7; text-decoration: underline;">{url}</a>
                  </div>
                </div>
                """
    else:
        relevant_html = """
        <div style="background-color: #F8FAFC; border: 1px dashed #CBD5E1; border-radius: 6px; padding: 14px; text-align: center; color: #64748B; font-size: 13px; margin: 15px 0;">
          No releases published so far today match core UPSC CSE criteria. All routine releases are logged below.
        </div>
        """

    routine_html = ""
    if routine:
        routine_items_html = ""
        for idx, item in enumerate(routine, 1):
            headline = str(item.get("headline", "")).replace("<", "&lt;").replace(">", "&gt;")
            ministry = str(item.get("ministry", "")).replace("<", "&lt;").replace(">", "&gt;")
            url = item.get("url", "")
            routine_items_html += f"""
            <li style="margin-bottom: 10px; font-size: 13px; line-height: 1.4;">
              <a href="{url}" target="_blank" style="color: #0F172A; text-decoration: none; font-weight: 600;">{headline}</a><br/>
              <span style="color: #64748B; font-size: 12px;">Ministry: {ministry}</span><br/>
              <span style="font-size: 12px;">🔗 PIB URL: <a href="{url}" target="_blank" style="color: #0284C7; text-decoration: underline;">{url}</a></span>
            </li>
            """
        routine_html = f"""
        <div style="margin-top: 25px; background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 14px 18px;">
          <h4 style="margin: 0 0 10px 0; color: #475569; font-size: 14px;">
            📋 Other Scraped PIB Releases (Routine / Non-UPSC) ({len(routine)})
          </h4>
          <ol style="margin: 0; padding-left: 20px; color: #334155;">
            {routine_items_html}
          </ol>
        </div>
        """

    return f"""<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1E293B; line-height: 1.5; margin: 0; padding: 12px; background-color: #F1F5F9;">
    <div style="max-width: 680px; margin: 0 auto; background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
      
      <!-- Banner -->
      <div style="background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%); color: #FFFFFF; padding: 20px 24px;">
        <h2 style="margin: 0; font-size: 20px; font-weight: 700; color: #FFFFFF;">🇮🇳 PIB Daily UPSC Intelligence</h2>
        <p style="margin: 4px 0 0 0; font-size: 13px; color: #93C5FD;">Date: {date_str} | Civil Services Examination (GS 1-4)</p>
      </div>

      <!-- Main Body -->
      <div style="padding: 20px 24px;">
        <p style="margin: 0 0 12px 0; font-size: 14px;">Hello,</p>
        <p style="margin: 0 0 16px 0; font-size: 14px; color: #334155;">
          Your automated PIB briefing for <b>{date_str}</b> has been compiled. Below are the scraped articles with their official hyperlinked URLs, categorized for UPSC CSE preparation.
        </p>

        <!-- Stats Box -->
        <div style="background-color: #EFF6FF; border-left: 4px solid #2563EB; border-radius: 4px; padding: 10px 14px; margin-bottom: 20px; font-size: 13px; color: #1E40AF;">
          <b>Daily Summary:</b> {len(releases)} Total Releases Scraped | <b>{len(relevant)} UPSC Relevant</b> | {len(routine)} Filtered Routine Notices
        </div>

        <!-- Relevant Articles Section -->
        <h3 style="margin: 16px 0 8px 0; font-size: 16px; color: #1E3A8A; border-bottom: 2px solid #DBEAFE; padding-bottom: 6px;">
          📚 UPSC Relevant PIB Articles
        </h3>
        {relevant_html}

        <!-- Other Scraped Releases Section -->
        {routine_html}

        <!-- Attachments Reminder -->
        <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 12px 16px; margin-top: 24px;">
          <p style="margin: 0; font-weight: bold; font-size: 13px; color: #1E40AF;">Attached Documents:</p>
          <ul style="margin: 6px 0 0 0; padding-left: 20px; font-size: 12.5px; color: #475569;">
            <li><b>Word Brief (.docx)</b>: Full formatted executive summaries with syllabus mapping.</li>
            <li><b>Excel Tracker (.xlsx)</b>: Color-coded tracker sheet with audit logs and direct links.</li>
          </ul>
        </div>

        <p style="color: #94A3B8; font-size: 12px; margin-top: 20px; text-align: center;">
          Automated dispatch from PIB UPSC GitHub Automation runner by Abhijeet
        </p>
      </div>
    </div>
  </body>
</html>"""


def send_daily_brief_email(target_date: str = None, to_override: str = None, releases: list = None):
    """
    Sends today's generated Word (.docx) and Excel (.xlsx) files to one or multiple recipients.
    Recipients can be comma-separated: 'user1@gmail.com, user2@gmail.com'.
    """
    load_env_file()

    mail_user = os.environ.get("MAIL_USERNAME", "").strip()
    mail_pass = os.environ.get("MAIL_PASSWORD", "").strip()
    mail_to = (to_override or os.environ.get("MAIL_TO") or mail_user).strip()

    if not mail_user or not mail_pass:
        print("[!] Email credentials (MAIL_USERNAME / MAIL_PASSWORD) not configured.")
        print("    Configure them in your .env file or environment variables.")
        return False

    date_str = target_date if target_date else date.today().strftime("%Y-%m-%d")
    
    # Locate today's generated files
    docx_candidates = list(OUTPUT_DIR.glob(f"PIB_UPSC_Daily_{date_str}*.docx"))
    xlsx_candidates = list(OUTPUT_DIR.glob(f"PIB_UPSC_Daily_{date_str}*.xlsx"))

    # Smart Morning Fallback: if no files found for today and no specific date forced, use yesterday's
    if not docx_candidates and not xlsx_candidates and not target_date:
        yesterday_str = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        y_docx = list(OUTPUT_DIR.glob(f"PIB_UPSC_Daily_{yesterday_str}*.docx"))
        y_xlsx = list(OUTPUT_DIR.glob(f"PIB_UPSC_Daily_{yesterday_str}*.xlsx"))
        if y_docx or y_xlsx:
            print(f"[*] No documents found for today ({date_str}). Using yesterday's briefing ({yesterday_str}).")
            date_str = yesterday_str
            docx_candidates, xlsx_candidates = y_docx, y_xlsx

    # Pick the most recently modified files
    docx_file = max(docx_candidates, key=lambda f: f.stat().st_mtime) if docx_candidates else None
    xlsx_file = max(xlsx_candidates, key=lambda f: f.stat().st_mtime) if xlsx_candidates else None

    if not docx_file and not xlsx_file:
        print(f"[!] No output documents found for {date_str} to email.")
        return False

    # Create message
    msg = MIMEMultipart()
    msg["From"] = f"PIB UPSC Automation <{mail_user}>"
    msg["To"] = mail_to
    msg["Subject"] = f"🇮🇳 PIB UPSC Daily Compilation - {date_str}"

    # Determine HTML body (uses rich hyperlinked briefing if available)
    html_file = OUTPUT_DIR / f"PIB_UPSC_Daily_{date_str}_email.html"
    if releases:
        html_body = generate_email_html(releases, date_str)
    elif html_file.exists():
        try:
            with open(html_file, "r", encoding="utf-8") as f:
                html_body = f.read()
        except Exception:
            html_body = None
    else:
        html_body = None

    if not html_body:
        html_body = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #1E293B; line-height: 1.6; padding: 10px;">
        <div style="background-color: #1E3A8A; color: white; padding: 16px 20px; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0;">🇮🇳 PIB Daily UPSC Sorting</h2>
          <p style="margin: 4px 0 0 0; font-size: 14px; opacity: 0.9;">Date: {date_str} | Civil Services Examination (GS 1-4)</p>
        </div>
        <div style="border: 1px solid #E2E8F0; border-top: none; padding: 20px; border-radius: 0 0 8px 8px; background: #FFFFFF;">
          <p>Hello,</p>
          <p>Your automated PIB briefing for <b>{date_str}</b> has been compiled and aligned with the UPSC syllabus.</p>
          
          <div style="background: #F8FAFC; border-left: 4px solid #2563EB; padding: 12px 16px; margin: 15px 0;">
            <p style="margin: 0; font-weight: bold; color: #1E40AF;">Attached Documents:</p>
            <ul style="margin: 8px 0 0 0; padding-left: 20px;">
              <li><b>Word Brief (.docx)</b>: Executive summary with releases grouped by GS-1, GS-2, GS-3, GS-4, topic tags, and source links.</li>
              <li><b>Excel Tracker (.xlsx)</b>: Color-coded tracker sheet + full audit log of all releases.</li>
            </ul>
          </div>
          
          <p style="color: #64748B; font-size: 13px;">This is an automated dispatch from PIB UPSC GitHub Automation runner by Abhijeet</p>
        </div>
      </body>
    </html>
    """
    msg.attach(MIMEText(html_body, "html"))

    # Attach files
    for filepath in [docx_file, xlsx_file]:
        if filepath and filepath.exists():
            try:
                with open(filepath, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f'attachment; filename="{filepath.name}"')
                    msg.attach(part)
                    print(f"[+] Attached: {filepath.name}")
            except Exception as e:
                print(f"[!] Error attaching {filepath.name}: {e}")

    # Dispatch via Gmail SMTP (Port 587 STARTTLS is universally allowed on cloud runners)
    ports_to_try = [
        (587, False),  # Port 587 with STARTTLS (Primary for cloud VMs)
        (465, True)    # Port 465 with direct SSL (Fallback)
    ]
    
    sent_successfully = False
    last_error = None

    for port, use_ssl in ports_to_try:
        try:
            print(f"[*] Connecting to smtp.gmail.com on port {port} (SSL={use_ssl}) to send to {mail_to}...")
            if use_ssl:
                server = smtplib.SMTP_SSL("smtp.gmail.com", port, timeout=25)
            else:
                server = smtplib.SMTP("smtp.gmail.com", port, timeout=25)
                server.ehlo()
                server.starttls()
                server.ehlo()

            # Clean password of any spaces
            clean_pass = mail_pass.replace(" ", "")
            server.login(mail_user, clean_pass)
            recipients = [r.strip() for r in mail_to.split(",") if r.strip()]
            server.send_message(msg, to_addrs=recipients)
            server.quit()
            print(f"[OK] Email successfully delivered to {len(recipients)} recipient(s): {', '.join(recipients)} via port {port}!")
            sent_successfully = True
            break
        except Exception as err:
            last_error = err
            print(f"[!] Attempt on port {port} failed: {err}")

    if not sent_successfully:
        print(f"[X] Could not send email via any port. Last error: {last_error}")
        sys.exit(1)

    return True


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="PIB UPSC Email Dispatcher")
    parser.add_argument("--date", type=str, help="Target date YYYY-MM-DD")
    parser.add_argument("--to", type=str, help="Recipient email address(es), comma-separated")
    args = parser.parse_args()
    send_daily_brief_email(target_date=args.date, to_override=args.to)
