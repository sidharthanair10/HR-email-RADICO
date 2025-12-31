import os
import pandas as pd
import win32com.client as win32
from dotenv import load_dotenv

load_dotenv()

excel_path = "HR-email.xlsx"  # Update if needed

print("📂 Reading Excel file...")
df_email = pd.read_excel(excel_path, sheet_name="email id")
df_data = pd.read_excel(excel_path, sheet_name="emp-info")
print("✅ Excel sheets loaded.")

# === Clean Empl ID columns ===
df_email["Empl ID"] = df_email["Empl ID"].astype(str).str.strip().str.split('.').str[0]
df_data["Empl ID"] = df_data["Empl ID"].astype(str).str.strip().str.split('.').str[0]

# === Merge data ===
merged_df = pd.merge(df_data, df_email, on="Empl ID", how="inner")
merged_df.fillna('', inplace=True)

if merged_df.empty:
    print("❌ No matching Empl ID values found. Exiting.")
    exit()

outlook = win32.Dispatch('Outlook.Application')

# === Send Emails ===
for _, row in merged_df.iterrows():
    subject = str(row.get("Subject", "")).strip()
    email_to = str(row.get("Email", "")).strip()
    if not email_to:
        print("⚠️ Skipping row due to missing email address.")
        continue

    name = str(row.get("Name", "")).strip()
    title = str(row.get("Title", "")).strip()
    custom_body = str(row.get("Mail-body", "")).strip()
    info_value = str(row.get("Information /related to", "")).strip()
    custom_body = custom_body.replace("(info)", info_value)

    body = f"""Dear {title} {name},<br><br>
This is to inform you regarding the subject: <b>{subject}</b>.<br><br>
{custom_body}<br><br>
Regards,<br>
HR Department"""

    mail = outlook.CreateItem(0)
    mail.To = email_to
    mail.Subject = subject
    mail.HTMLBody = body

    # === Handle attachment ===
    attachment_path = str(row.get("Attachment", "")).strip()
    if attachment_path:
        print(f"📎 Trying to attach: {attachment_path}")
        if os.path.exists(attachment_path):
            try:
                mail.Attachments.Add(Source=attachment_path)
                print(f"✅ Attached: {attachment_path}")
            except Exception as e:
                print(f"⚠️ Could not attach file for {email_to}: {e}")
        else:
            print(f"⚠️ File not found: {attachment_path}. Sending without attachment.")
    else:
        print(f"ℹ️ No attachment specified for {email_to}. Sending without attachment.")

    # Send email
    try:
        mail.Send()
        print(f"✅ Email successfully sent to {email_to}")
    except Exception as e:
        print(f"❌ Failed to send email to {email_to}: {e}")
