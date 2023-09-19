import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import csv
from app import app
# Step 1: Access Your Email Accountfrom app import app (Gmail in this example)
# Replace 'YOUR_EMAIL', 'YOUR_PASSWORD', and 'YOUR_EMAIL_SERVER' with your email credentials and server details.

email_user = 'senthurkumaran2004@gmail.com'
email_pass = 'llzs ydsg wrtk xsip'
email_server = 'imap.gmail.com'  # Update with your email server

# Connect to the email server
mail = imaplib.IMAP4_SSL(email_server)
mail.login(email_user, email_pass)

# Step 3: Select the Mailbox (e.g., "inbox")
mail.select("inbox")

# Step 2: Define the Time Frame
# Replace 'START_DATE' and 'END_DATE' with the desired date range.
# The code below selects emails received in the last 7 days.

END_DATE = datetime.now()
START_DATE = END_DATE - timedelta(days=7)

# Convert dates to the required format for IMAP search
end_date_str = END_DATE.strftime("%d-%b-%Y")
start_date_str = START_DATE.strftime("%d-%b-%Y")

# Step 4: Search for Emails Within the Time Frame
search_criteria = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'

# Search for emails matching the criteria
status, email_ids = mail.search(None, search_criteria)

# Initialize a list to store email subjects
email_subjects = []

# Loop through email IDs and fetch email subjects
for email_id in email_ids[0].split():
    status, email_data = mail.fetch(email_id, "(RFC822)")
    raw_email = email_data[0][1]
    email_message = email.message_from_bytes(raw_email)
    
    # Extract and decode the subject
    subject, charset = decode_header(email_message["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(charset if charset else 'utf-8')
    
    email_subjects.append(subject)

# Print or use the 'email_subjects' list as needed
print("Email Subjects for the Last 7 Days:")
for subject in email_subjects:
    print(subject)

# Save email subjects to a CSV file
csv_filename = 'email_subjects.csv'  # Specify the filename
with open(csv_filename, 'w', newline='', encoding='utf-8') as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(['Email Subject'])  # Write a header row
    writer.writerows([[subject] for subject in email_subjects])  # Write email subjects to the CSV file

# Close the email connection
mail.logout()
if __name__ == "__main__":
    app.run(debug=True)