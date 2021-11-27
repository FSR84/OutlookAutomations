import win32com.client
from datetime import datetime, timedelta
import os


# define the Outlook App and Namespace
outlookApp = win32com.client.Dispatch("Outlook.Application")
outlookNS = outlookApp.GetNamespace("MAPI")


# define your address and folder
inbox = outlookNS.Folders['mail@gmail.com'].Folders['Inbox'] # (another .Folder will access a subfolder)
#inbox = outlookNS.GetDefaultFolder(6) # use this line if you have just one account in Outlook and want to use the Inbox folder
messages = inbox.Items
print('You have ' + str(len(messages)) + ' e-mails in ' + str(inbox) +'.')


# filter by date received
received_dt = datetime.now() - timedelta(days=30)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
print('After using Restrict, you have ' + str(len(messages)) + ' e-mails in ' + str(inbox) +'.')


# filter by range of dates
#start_time = datetime.today().replace(month=1, hour=0, minute=0, second=0).strftime('%Y-%m-%d %H:%M %p') # first day of the month
#end_time = datetime.today().replace(hour=0, minute=0, second=0).strftime('%Y-%m-%d %H:%M %p') # today 12am
#messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")


# filter by address
#messages = messages.Restrict("[SenderEmailAddress] = 'contact@codeforests.com'")
#messages = messages.Restrict("Not ([SenderEmailAddress] = 'abc@company.com')")


# filter by subject
#messages = messages.Restrict("[Subject] = 'Sample Report'")


# filter by address and subject
#messages = messages.Restrict("[Subject] = 'Sample Report'" + " And Not ([SenderEmailAddress] = 'abc@company.com')")


# print last 10 messages (sorting is necessary!)
messages.Sort("[ReceivedTime]", Descending=True)
for message in list(messages)[:10]:
	print(message.Subject, str(message.ReceivedTime), message.SenderEmailAddress)


# save attachments
save_folder = os.getcwd() + "\\OutlookAutomations\\Attachments\\"
for message in list(messages):
	try:
		for attachment in message.Attachments:
			attachment.SaveASFile(os.path.join(save_folder, attachment.FileName)) # overwrites without warning!
			print(f"Attachment {attachment.FileName} from {message.sender} saved.")
	except Exception as e:
		print("Error when saving the attachment:" + str(e))		
