import imaplib
import email
import pandas as pd
import yaml  

with open("credentials.yml") as f:
    content = f.read()
    
# from credentials.yml import user name and password
my_credentials = yaml.load(content, Loader=yaml.FullLoader)

#Load the user name and passwd from yaml file
user, password = my_credentials["user"], my_credentials["password"]

#URL for IMAP connection
imap_url = 'imap.gmail.com'

# Connection with GMAIL using SSL
my_mail = imaplib.IMAP4_SSL(imap_url)

# Log in using your credentials
my_mail.login(user, password)

# Select the Inbox to fetch messages
my_mail.select('Inbox')

#Define Key and Value for email search
key = 'FROM'
value = 'example@example.com'
#Search for emails with specific key and value
status, mail_id_lists = my_mail.search(None, key, value)  
#IDs of all emails that we want to fetch 
mail_id_lists = mail_id_lists[0].split()  

# Initialize lists to store email data
mail = []
senders = []
dates = []
subjects = []
bodies = []

# Iterate through email IDs and fetch email content
for num in mail_id_lists:
    status, msg_data = my_mail.fetch(num, "(RFC822)")
    msg = email.message_from_bytes(msg_data[0][1])

    sender = msg['from']
    date = msg['date']
    subject = msg['subject']
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode("utf-8")
                break
    else:
        body = msg.get_payload(decode=True).decode("utf-8")
    senders.append(sender)
    dates.append(date)
    subjects.append(subject)
    bodies.append(body)
    
# Create a DataFrame from the collected email data
print("Creating Dataframe...")
extracted_info = pd.DataFrame(columns=["Sender", "Date", "Subject", "Email Body"])
extracted_info["Sender"] = senders
extracted_info["Date"] = dates
extracted_info["Subject"] = subjects
extracted_info["Email Body"] = bodies
print("Dataframe created")

#Save the dataframe as an excel document. 
extracted_info.to_excel("exportfile.xlsx")
print("Dataframe exported to excel")
