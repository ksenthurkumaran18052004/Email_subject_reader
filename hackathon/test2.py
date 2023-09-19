import imaplib
import numpy as np
import email
from email.header import decode_header
from datetime import datetime, timedelta
import nltk
nltk.download('punkt')
nltk.download('stopwords')
from nltk import sent_tokenize
from nltk.corpus import stopwords
from nltk.cluster.util import cosine_distance  # Fix the import statement
import networkx as nx
import re
import html2text
import csv

# Define the function to calculate sentence similarity
def sentence_similarity(sent1, sent2, stopwords=None):
    if stopwords is None:
        stopwords = []
    
    sent1 = [word.lower() for word in sent1]
    sent2 = [word.lower() for word in sent2]
    
    all_words = list(set(sent1 + sent2))
    
    vector1 = [0] * len(all_words)
    vector2 = [0] * len(all_words)
    
    # Build the vector for the first sentence
    for w in sent1:
        if w not in stopwords:
            vector1[all_words.index(w)] += 1
            
    # Build the vector for the second sentence
    for w in sent2:
        if w not in stopwords:
            vector2[all_words.index(w)] += 1
            
    return 1 - cosine_distance(vector1, vector2)

# Define the function to generate summary
# Define the function to generate summary with enhanced text cleaning
def generate_summary(text, num_sentences=3):
    nltk.download('punkt')
    nltk.download('stopwords')
    
    # Tokenize the text into sentences
    sentences = sent_tokenize(text)
    
    # Add the following two lines to print the raw email content
    print("Raw Email Content:")
    print(text)
    
    # Create a similarity matrix as a NumPy array
    similarity_matrix = np.zeros((len(sentences), len(sentences)))

    # Create a list of English stopwords
    stop_words = stopwords.words('english')  # Use 'stopwords' instead of 'stop_words'
    
    # Define a regular expression to identify URLs
    url_pattern = re.compile(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+')
    
    # Remove URLs from the sentences
    cleaned_sentences = []
    for sentence in sentences:
        sentence_without_urls = re.sub(url_pattern, '', sentence)
        
        # Additional text cleaning (customize this based on your needs)
        # For example, removing lines starting with 'Meta Platforms, Inc.'
        sentence_without_urls = re.sub(r'^Meta Platforms, Inc\..*', '', sentence_without_urls)
        
        cleaned_sentences.append(sentence_without_urls)
    
    # Populate the similarity matrix
    for i in range(len(cleaned_sentences)):
        for j in range(len(cleaned_sentences)):
            if i != j:
                similarity_matrix[i][j] = sentence_similarity(cleaned_sentences[i], cleaned_sentences[j], stop_words)

    # Use PageRank algorithm to rank sentences
    nx_graph = nx.from_numpy_array(similarity_matrix)
    scores = nx.pagerank(nx_graph)
    
    # Sort sentences by their score
    ranked_sentences = sorted(((scores[i], s) for i, s in enumerate(cleaned_sentences)), reverse=True)
    
    # Extract the top 'num_sentences' sentences as the summary
    summary = " ".join([s[1] for s in ranked_sentences[:num_sentences]])
    
    return summary


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

    # Generate the summary of the email text
    email_summary = generate_summary(email_text, num_sentences=3)

    # Print or use the 'email_summary' variable as needed
    print("Summary of the email:")
    print(email_summary)
    csv_filename = 'email_summary.csv'  # Specify the filename
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['Email Summary'])  # Write a header row
        writer.writerow([email_summary])  # Write the email summary to the CSV file

else:
    print("No emails found within the specified time frame.")

# Close the email connection
mail.logout()
