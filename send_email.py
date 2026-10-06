import os
import sys
import smtplib
import socket
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import date

# Force IPv4 socket resolution to prevent dropped connections on Azure/GitHub Actions runners
_orig_getaddrinfo = socket.getaddrinfo
def _getaddrinfo_ipv4(host, port, family=0, type=0, proto=0, flags=0):
    res = _orig_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)
    return res if res else _orig_getaddrinfo(host, port, family, type, proto, flags)
socket.getaddrinfo = _getaddrinfo_ipv4

# Ensure paths
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"


def send_daily_brief_email(target_date: str = None):
    """
    Sends today's generated Word (.docx) and Excel (.xlsx) files to the specified recipient.
    Reads credentials from environment variables:
      - MAIL_USERNAME: sender email (e.g. your_email@gmail.com)
      - MAIL_PASSWORD: Gmail App Password (16 characters)
      - MAIL_TO: recipient email (defaults to MAIL_USERNAME if not provided)
    """
    mail_user = os.environ.get("MAIL_USERNAME", "").strip()
    mail_pass = os.environ.get("MAIL_PASSWORD", "").strip()
    mail_to = os.environ.get("MAIL_TO", mail_user).strip()

    if not mail_user or not mail_pass:
        print("[!] Email credentials (MAIL_USERNAME / MAIL_PASSWORD) not configured. Skipping email.")
        return False

    date_str = target_date if target_date else date.today().strftime("%Y-%m-%d")
    
    # Locate today's generated files
    docx_candidates = list(OUTPUT_DIR.glob(f"PIB_UPSC_Daily_{date_str}*.docx"))
    xlsx_candidates = list(OUTPUT_DIR.glob(f"PIB_UPSC_Daily_{date_str}*.xlsx"))

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

    html_body = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #1E293B; line-height: 1.6; padding: 10px;">
        <div style="background-color: #1E3A8A; color: white; padding: 16px 20px; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0;">🇮🇳 PIB Daily UPSC Intelligence</h2>
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
          
          <p style="color: #64748B; font-size: 13px;">This is an automated dispatch from your PIB UPSC GitHub Automation runner.</p>
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
            server.send_message(msg)
            server.quit()
            print(f"[✓] Email successfully delivered to {mail_to} via port {port}!")
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
    target = sys.argv[1] if len(sys.argv) > 1 else None
    send_daily_brief_email(target_date=target)
