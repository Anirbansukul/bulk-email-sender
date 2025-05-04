import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
import random
from getpass import getpass
from IPython.display import display, HTML
import os
import re
from google.colab import files
import traceback
print("📁 Upload the CSV file containing 'email' and 'name' columns...")
uploaded = files.upload()
csv_filename = next(iter(uploaded))
df = pd.read_csv(csv_filename)
required_cols = {'email', 'name'}
if not required_cols.issubset(df.columns):
    raise ValueError(f"CSV must contain the columns: {required_cols}")
email_regex = r'^\S+@\S+\.\S+$'
df = df[df['email'].apply(lambda x: re.match(email_regex, str(x)) is not None)]

display(HTML(df.to_html(index=False)))

EMAIL_ADDRESS = input("Enter your email: ")
EMAIL_PASSWORD = getpass("Enter your app password (not your real password): ")
SENDER_NAME = input("Enter your name to appear in the 'From' field: ")
subject = input("Enter the subject of your email: ")
print("\n📨 Enter your custom HTML message or type 'default' to use the built-in one.")
print("You can use placeholders like {name} and {sender}.")
user_input = input("Message: ").strip()

default_html = '''
<html>
  <body>
    <p>Hi {name},<br><br>
       This is a <b>personalized</b> message sent through our final year email sender project.<br><br>
       Best regards,<br><b>{sender}</b>
    </p>
  </body>
</html>
'''

html_message = default_html if user_input.lower() == "default" else user_input

sample = df.iloc[0]
preview = html_message.format(name=sample['name'], sender=SENDER_NAME)
print("\nPreview of the email (HTML format):")
display(HTML(preview))

send_test = input("Do you want to send a test email to yourself first? (y/n): ").strip().lower()
if send_test == 'y':
    test_msg = MIMEMultipart()
    test_msg['From'] = f"{SENDER_NAME} <{EMAIL_ADDRESS}>"
    test_msg['To'] = EMAIL_ADDRESS
    test_msg['Subject'] = f"[TEST] {subject}"
    test_msg.attach(MIMEText(preview, 'html'))

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    server.send_message(test_msg)
    server.quit()
    print("Test email sent to yourself.\n")

print("\n📎 Upload a file to attach (or skip if none)...")
uploaded_files = files.upload()
attachment_filename = next(iter(uploaded_files), None) if uploaded_files else None

try:
    delay_min = int(input("Enter minimum delay between emails (in seconds, e.g. 5): "))
    delay_max = int(input("Enter maximum delay between emails (in seconds, e.g. 10): "))
except:
    print("Invalid input, using default 5-10 seconds.")
    delay_min, delay_max = 5, 10

log = []
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()

try:
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    for index, row in df.iterrows():
        msg = MIMEMultipart()
        msg['From'] = f"{SENDER_NAME} <{EMAIL_ADDRESS}>"
        msg['To'] = row['email']
        msg['Subject'] = subject

        personalized_html = html_message.format(name=row['name'], sender=SENDER_NAME)
        plain_text = f"Hi {row['name']},\n\nThis is a personalized message.\n\nRegards,\n{SENDER_NAME}"
        msg.attach(MIMEText(plain_text, 'plain'))
        msg.attach(MIMEText(personalized_html, 'html'))

        if attachment_filename:
            try:
                with open(attachment_filename, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={attachment_filename}')
                    msg.attach(part)
            except Exception as e:
                print(f"Attachment error for {row['email']}: {e}")
                log.append((row['email'], f"Failed: Attachment error"))
                continue

        for attempt in range(2):
            try:
                server.send_message(msg)
                print(f"Sent to {row['name']} ({row['email']})")
                log.append((row['email'], "Success"))
                break
            except Exception as e:
                if attempt == 0:
                    print(f"Retrying for {row['email']} due to error: {e}")
                    time.sleep(3)
                else:
                    print(f"Failed to send to {row['email']}: {e}")
                    log.append((row['email'], f"Failed: {str(e)}"))

        time.sleep(random.randint(delay_min, delay_max))

finally:
    server.quit()

log_df = pd.DataFrame(log, columns=["Email", "Status"])
log_filename = "email_send_log.csv"
log_df.to_csv(log_filename, index=False)
print(f"\nLog saved as {log_filename}")
display(log_df)
files.download(log_filename)

failed_df = log_df[log_df["Status"] != "Success"]
if not failed_df.empty:
    failed_filename = "failed_emails.csv"
    failed_df.to_csv(failed_filename, index=False)
    print(f"Failed emails saved as {failed_filename}")
    files.download(failed_filename)

print("\nSummary:")
print(log_df["Status"].value_counts())
