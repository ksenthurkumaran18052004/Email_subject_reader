import openai
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import time  # Import the time module

# Define the get_completion function
def get_completion(prompt, model="gpt-3.5-turbo"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0,
    )
    return response.choices[0].text.strip()


# Step 1: Access Your Email Account (Gmail in this example)
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

# Get the latest email (assuming there's only one matching this criteria)
if email_ids[0]:
    email_ids = email_ids[0].split()
    latest_email_id = email_ids[-1]

    # Fetch the email content
    status, email_data = mail.fetch(latest_email_id, "(RFC822)")

    # Step 5: Copy the Email Text
    # Parse the email content and copy the text

    raw_email = email_data[0][1]
    email_message = email.message_from_bytes(raw_email)

    email_text = ""

    # Helper function to decode email text
    def decode_text(text, charset):
        try:
            decoded_text, _ = decode_header(text)[0]
            if isinstance(decoded_text, bytes):
                return decoded_text.decode(charset if charset else 'utf-8', errors='ignore')
            else:
                return decoded_text
        except Exception as e:
            print(f"Error decoding text: {e}")
            return text

    for part in email_message.walk():
        if part.get_content_type() == "text/plain":
            charset = part.get_content_charset()
            email_text += decode_text(part.get_payload(), charset)

    # Print or use the 'email_text' variable as needed
    print("Email Text:")
    print(email_text)

    # Step 6: Access ChatGPT (OpenAI API setup)
    api_key = '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'
    openai.api_key = api_key

    # Add a sleep of 2 seconds before making the API request
    time.sleep(2)

    # Step 7: Ask ChatGPT to Summarize
    summary_request = "Please summarize the content of this email: \n" + email_text

    # Step 8: Receive the Summary
    response = openai.Completion.create(
        engine="text-davinci-002",
        prompt=summary_request,
        max_tokens=100,  # Adjust max_tokens as needed for the desired summary length
    )

    summary = response.choices[0].text.strip()

    # Step 9: Edit and Refine (if necessary)
    # You can further interact with ChatGPT to refine the summary if needed.

    # Step 10: Save or Use the Summary
    print("Summary of the email:")
    print(summary)
else:
    print("No emails found within the specified time frame.")

# Close the email connection
mail.logout()
