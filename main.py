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
def send_emails(email_list):
    body = HTML_BODY()
    olApp = win32.Dispatch('Outlook.Application')
    with open(email_list) as csv_file:
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
            mailItem.Attachments.Add(Source = "C:\\Users\\KaiAchen\\Scripts\\Email Scripts\\Learning2Fly_Field_Trip_Poster.pdf")
            # mailItem.Display()
            mailItem.Send()

# CSV -> CSV
def check_overlap(list_1, list_2):
    set_1 = set()
    set_2 = set()
    column_index = 0

    with open(list_1, 'r', newline='') as f1: # Read CSV 1 into Set
        reader1 = csv.reader(f1)
        for row in reader1:
            set_1.add(row[column_index])

    with open(list_2, 'r', newline='') as f2: # Read CSV 2 into Set
        reader2 = csv.reader(f2)
        for row in reader2:
           set_2.add(row[column_index])

    intersection_set = set_1.intersection(set_2) # Find Intersection of Sets (A ∩ B)
    difference_set = set_1.difference(set_2) # Find Difference of Sets (A ∩ B')

    percent_overlap_set_1 = round((((len(intersection_set) / len(set_1))) * 100), 2) # Calculate Overlap Percentages
    print(f"Percent Overlap Set 1: {percent_overlap_set_1}%")
    percent_overlap_set_2 = round((((len(intersection_set) / len(set_2))) * 100), 2)
    print(f"Percent Overlap Set 2: {percent_overlap_set_2}%")

    with open("intersection_set.csv", "w", newline='') as f3: # Write (A ∩ B) to CSV
        writer1 = csv.writer(f3)
        for item in intersection_set:
            writer1.writerow([item])
    
    with open("difference_set.csv", "w", newline='') as f4: # Write (A ∩ B') to CSV
        writer2 = csv.writer(f4)
        for item in difference_set:
            writer2.writerow([item])


def main():
    quit_program = False
    while(not(quit_program)):
        print("Make Selection:")
        print("1. Send Emails to emails.csv")
        print("2. Send Emails to test.csv")
        print("3. Check Overlap of Two CSV Lists")
        print("Q/q. Quit Program")
        selection = input("Enter Selection Number: ")
        if selection == "1":
            send_emails("emails.csv")
            print("Emails Sent")
        elif selection == "2":
            send_emails("test.csv")
            print("Emails Sent")
        elif selection == "3":
            check_overlap("overlap1.csv", "overlap2.csv")
            print("Overlap Calculated and Sent to intersection_set.csv and difference_set.csv")
        elif selection == "Q" or selection == "q":
            print("Quitting Program")
            quit_program = True
        else:
            print("Invalid Selection")



if __name__ == "__main__":
    main()