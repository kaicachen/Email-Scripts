import csv
from html_body import HTML_BODY
import win32com.client as win32

# Str -> Str
def clean_address(address):
    address = address.replace('[', '').replace(']', '')
    address = address.replace("'", '')
    address = address.replace(",", '')
    address = address.replace(" ", '')
    return address

# CSV -> Void
def send_emails():
    body = HTML_BODY()
    olApp = win32.Dispatch('Outlook.Application')
    with open('test.csv') as csv_file:
        reader = csv.reader(csv_file)
        for line in reader:
            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'LEARNING2FLY FIELD TRIP OPPORTUNITIES'
            mailItem.HTMLbody = body.get_text()
            address = '<' + str(line) + '>'
            address = clean_address(address)
            # print(address)
            mailItem.To = address
            mailItem.Sensitivity  = 2
            mailItem.Attachments.Add(Source = "C:\\Users\\KaiAchen\\Oakwood Management\\Oakwood Management - Chakra Circus - Staff\\Resources\\Scripts\\Email Scripts\\Learning2Fly_Field_Trip_Poster.pdf")
            # mailItem.Display()
            mailItem.Send()

def main():
    send_emails()
    print("Emails Sent")

if __name__ == "__main__":
    main()