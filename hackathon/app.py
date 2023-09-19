from flask import Flask, render_template, request, redirect, url_for
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import csv

app = Flask(__name__)

# Define your email configuration here
email_user = 'senthurkumaran2004@gmail.com'
email_pass = 'llzs ydsg wrtk xsip'
email_server = 'imap.gmail.com'

# Define the fetch_email_subjects function to fetch email subjects
def fetch_email_subjects(user_email):
    try:
        # Connect to the email server
        mail = imaplib.IMAP4_SSL(email_server)
        mail.login(email_user, email_pass)

        # Select the mailbox (inbox in this case)
        mail.select("inbox")

        # Define the time frame
        END_DATE = datetime.now()
        START_DATE = END_DATE - timedelta(days=7)
        end_date_str = END_DATE.strftime("%d-%b-%Y")
        start_date_str = START_DATE.strftime("%d-%b-%Y")

        # Search for emails within the time frame
        search_criteria = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'
        status, email_ids = mail.search(None, search_criteria)

        email_subjects = []
        for email_id in email_ids[0].split():
            status, email_data = mail.fetch(email_id, "(RFC822)")
            raw_email = email_data[0][1]
            email_message = email.message_from_bytes(raw_email)

            subject, charset = decode_header(email_message["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(charset if charset else 'utf-8')

            email_subjects.append(subject)

        mail.logout()
        return email_subjects
    except Exception as e:
        # Handle any exceptions that may occur during email fetching
        print(f"Error fetching emails: {e}")
        return []

# Define the save_to_csv function to save email subjects to a CSV file
def save_to_csv(data, filename):
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(['Email Subject'])
            writer.writerows([[subject] for subject in data])
    except Exception as e:
        # Handle any exceptions that may occur during CSV writing
        print(f"Error saving to CSV: {e}")

# Define the read_csv function to read email subjects from a CSV file
def read_csv(filename):
    try:
        email_subjects = []
        with open(filename, 'r', encoding='utf-8') as csv_file:
            reader = csv.reader(csv_file)
            next(reader)  # Skip the header row
            for row in reader:
                email_subjects.append(row[0])
        return email_subjects
    except Exception as e:
        # Handle any exceptions that may occur during CSV reading
        print(f"Error reading CSV: {e}")
        return []

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        user_email = request.form["user_email"]
        email_subjects = fetch_email_subjects(user_email)
        csv_filename = 'email_subjects.csv'
        save_to_csv(email_subjects, csv_filename)
        return redirect(url_for("success"))

    return render_template("index.html")  # Render the HTML page

@app.route("/success")
def success():
    email_subjects = read_csv('email_subjects.csv')
    return render_template("success.html", email_subjects=email_subjects)

if __name__ == "__main__":
    app.run(debug=True)
