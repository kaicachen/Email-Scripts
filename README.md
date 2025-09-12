# Basic Outlook Sending Script
## Requirements:
- Windows 10 or later
- Python3
- win32com.client
- Outlook
- OS User-tied Outlook Account
## Usage Guide:
1. Create local file in Email Scripts root folder called "emails.csv"
2. Write/Paste/Load emails into "emails.csv", with each email on a separate line and followed by a comma
3. Check that the email contents are as desired: (Subject and Attachments can be changed in main.send_emails(), Body can be changed in html_body.text)
4. Run main.py -> Console will output when emails successfully sent to Outlook "Outbox" folder
5. Check Outlook "Outbox" folder to ensure that emails have been sent
## Email Contents:
- Found in html_body.py class
- HTML format
## Input:
- CSV of email addresses
## Output:
- Sent emails
## TODO:
- Auto clean list of emails and compile to CSV
- GUI? -> easier for boss to use
- Read Outlook responses and compile CSV of response types
- Pretty-fy HTML body (more colors)
