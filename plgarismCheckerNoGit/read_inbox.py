# Read normal outlook inbox messages

import win32com.client
import datetime as dt
import csv

# Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
DL_NAME = 'dummy_group'
HEADER = ['Subject', 'Body']
data = []
lastWeekDateTime = dt.datetime.now()
yesterday_date = lastWeekDateTime.date()

for message in messages:
    if message.ReceivedTime.date() >= yesterday_date:
        print(message.ReceivedTime.date())
        print(message.Subject)
        # data.append([ message.Subject, message.Body])

with open('inbox.csv', 'w', encoding='UTF8', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(HEADER)
    writer.writerow(data)

# message.ReceivedTime.date()
# message.SenderName
# message.To
# message.Subject
# message.Body