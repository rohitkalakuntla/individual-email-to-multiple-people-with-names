import os
import csv
import win32com.client
from datetime import datetime
import json

# Set paths
base_dir = os.getcwd()
csv_path = os.path.join(base_dir, "recipients.csv")
body_path = os.path.join(base_dir, "email_body.html")
log_path = os.path.join(base_dir, "email_log.txt")  # Single persistent log file


# Logging function
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")
    print(message)

# Log section starting point
with open(log_path, "a", encoding="utf-8") as log_file:
    log_file.write("\n\n")  # Leave two blank lines before this run
    
log_message(f"---- Starting the logging for sending the emails ---------")


# Load json config
config_path = os.path.join(base_dir, "Need_to_update_details.json")
try:
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
except Exception as e:
    log_message(f"Error reading config file: {e}")
    exit()

# Access json config values
subject = config.get("subject", "Default Subject")
cc_email = config.get("cc_email", "")
attachments = config.get("attachments", [])


# Read html body content
try:
    with open(body_path, "r", encoding="utf-8") as f:
        body_template = f.read()
except Exception as e:
    log_message(f"Error reading email body file: {e}")
    exit()

# Read recipient list
recipients = []
skipped_count = 0
total_rows_in_csv = 0

try:
    with open(csv_path, newline='') as csvfile:
        reader = csv.DictReader(filter(lambda row: not row.startswith('#'), csvfile))
        for row in reader:
            
            total_rows_in_csv += 1  # ✅ Count every row, even skipped
            
            #first_name = row.get("first_name", "").strip()
            #last_name = row.get("last_name", "").strip()
            #email = row.get("email", "").strip()
            
            first_name = (row.get("first_name") or "").strip()
            last_name = (row.get("last_name") or "").strip()
            email = (row.get("email") or "").strip()

            if not first_name or not last_name or not email:
                log_message(f"⚠️ Skipped row due to missing field(s): {row}")
                skipped_count += 1
                continue

            recipients.append({
                "first_name": first_name,
                "last_name": last_name,
                "email": email
            })
except Exception as e:
    log_message(f"Error reading recipients CSV: {e}")
    log_message(f"--------------------------------- DONE -----------------------------------")
    exit()

# Send emails

success_count = 0
failed_emails = []


try:
    log_message("Starting multiple email sending process...")

    outlook = win32com.client.Dispatch("Outlook.Application")

    for r in recipients:
        try:
            personalized_body = f"""
            Hi {r['first_name']} {r['last_name']},<br><br>
            {body_template}
            """
    
            mail = outlook.CreateItem(0)
            mail.To = r['email']
            mail.CC = cc_email
            mail.Subject = subject
            mail.HTMLBody = personalized_body
    
            for filename in attachments:
                full_path = os.path.join(base_dir, filename)
                if os.path.exists(full_path):
                    mail.Attachments.Add(full_path)
                else:
                    print(f"Attachment not found: {filename}")
    
            mail.Send()
            success_count += 1
            log_message(f"✅ Email sent to {r['first_name']} {r['last_name']} ({r['email']})")
            
        except Exception as e:
            log_message(f"❌ Failed to send email to {r['first_name']} {r['last_name']} ({r['email']}): {e}")
            failed_emails.append(f"{r['first_name']} {r['last_name']} ({r['email']})")

except Exception as e:
    log_message(f"❌ Global error occurred: {e}")
    
# Final summary
total_expected = len(recipients)

log_message(f"📊 Summary: Total rows in CSV: {total_rows_in_csv} | Emails Planned to send: {total_expected} | Successfully Sent: {success_count} | Skipped: {skipped_count}")


if success_count == total_expected:
    log_message("✅ All emails sent successfully.")
else:
    log_message("⚠️ Some emails failed. Please check the logs above.")

if failed_emails:
    log_message("❌ Failed email recipients:")
    for entry in failed_emails:
        log_message(f"   - {entry}")

log_message(f"--------------------------------- DONE -----------------------------------")