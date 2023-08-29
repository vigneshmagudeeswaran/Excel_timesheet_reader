import imaplib
import email
import os
from datetime import datetime, timedelta
import schedule
import time
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# Email Configuration
email_user = os.getenv("EMAIL_USER")
email_pass = os.getenv("EMAIL_PASS")
imap_server = "imap.gmail.com"

# Root directory to store downloaded Excel files
root_download_dir = "downloaded_excel_files"

# Create a directory for downloaded files if it doesn't exist
if not os.path.exists(root_download_dir):
    os.mkdir(root_download_dir)

# List of employee email addresses and their corresponding sender names
employee_email_senders = {
    "bharanim.igs@gmail.com": "Bharani",
    "isaacrajselwyn1994@gmail.com": "Issac",
    "vijayanlokesh16@gmail.com": "Lokesh",
    "asvar.ahamed@intellecto.co.in": "Asvar",
    "adamofficial0145@gmail.com" :'Adam',
    "srinivasangogul.m@intellecto.co.in": "Gogul"
}

# Time period filter
today = datetime.today()
start_date = today - timedelta(days=7)  # Adjust the number of days as needed

# Connect to Gmail IMAP server
mail = imaplib.IMAP4_SSL(imap_server)
mail.login(email_user, email_pass)
mail.select("inbox")

# Function to download Excel attachments from emails
def download_excel_attachments():
    try:
        # Iterate through employees and download their emails
        for employee_email, sender_name in employee_email_senders.items():
            # Create a folder for the employee if it doesn't exist
            employee_download_dir = os.path.join(root_download_dir, sender_name)
            if not os.path.exists(employee_download_dir):
                os.mkdir(employee_download_dir)

            # Search for emails matching the filter for this employee and week
            search_query = f"FROM {employee_email} SINCE {start_date.strftime('%d-%b-%Y')}"
            result, data = mail.search(None, search_query)
            email_ids = data[0].split()

            for email_id in email_ids:
                # Fetch the email
                result, message_data = mail.fetch(email_id, "(RFC822)")
                message = email.message_from_bytes(message_data[0][1])

                # Iterate over email parts
                for part in message.walk():
                    if part.get_content_type() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                        # Extract the attachment
                        filename = part.get_filename()
                        if filename:
                            # Construct the full path for this week's download
                            filepath = os.path.join(employee_download_dir, f"{sender_name}_{start_date.strftime('%Y-%m-%d')}_{filename}")

                            # Save the Excel file
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))

                            print(f"Downloaded: {filename} (from {sender_name})")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Function to combine Excel files by date
def combine_excel_files():
    try:
        # Create a folder for combined files
        combined_dir = os.path.join(root_download_dir, "combined")
        if not os.path.exists(combined_dir):
            os.mkdir(combined_dir)

        # Get a list of Excel files from all employee folders
        excel_files = []
        for employee_name in employee_email_senders.values():
            employee_dir = os.path.join(root_download_dir, employee_name)
            employee_files = [os.path.join(employee_dir, file) for file in os.listdir(employee_dir) if file.endswith(".xlsx")]
            excel_files.extend(employee_files)

        if excel_files:
            # Combine Excel files while preserving all columns as strings
            combined_df = pd.concat([pd.read_excel(file, dtype=str) for file in excel_files], ignore_index=True)

            combined_filename = os.path.join(combined_dir, f"combined_{start_date.strftime('%Y-%m-%d')}.xlsx")
            combined_df.to_excel(combined_filename, index=False)
            print(f"Combined Excel created for date {start_date.strftime('%Y-%m-%d')}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")


        
download_excel_attachments()
combine_excel_files()
# Uncomment this section to schedule the tasks
# schedule.every().tuesday.at("19:09").do(download_excel_attachments)
# schedule.every().tuesday.at("19:09").do(combine_excel_files)

# # Run the scheduled jobs
# while True:
#     schedule.run_pending()
#     time.sleep(1)
