# version 2 with name
import win32com.client as client
import csv
import time
from datetime import datetime
outlook = client.Dispatch("Outlook.Application")
with open('Student links - Sheet3.csv', newline='') as f:
    reader = csv.reader(f)
    distro = [row for row in reader]
count=1
print("Now starting for loop")
for email, link in distro:
    if count % 20 == 0:
        print('Now waiting 61 seconds')
        time.sleep(61)
        count += 1
    else:
        message = outlook.CreateItem(0)
        message.To = email
        message.Subject = "Assignment Spreadsheet"
        bodyMessage = "Hello there, \n\nWelcome back to another year at Northwood High School! Attached below is a link to a spreadsheet that contains your current assignments and missing assignments. These will be updated daily so make sure to check to stay up to date with all your assignments. On the second tab you will see a column named '(Students) Finished' where you can check off what you have finished.\n" + link + "\n\nSincerely, \nJeremy "
        message.Body = bodyMessage
        message.Send()
        count += 1
        print("Successfully sent link to", email,"at",datetime.now(),"count at: ",count-1)


print("Finished sending emails")
