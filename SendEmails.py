import win32com.client

outlookApp = win32com.client.Dispatch("Outlook.Application")
outlookNS = outlookApp.GetNamespace("MAPI")

mail = outlookApp.CreateItem(0)

# define send_from if you have multiple accounts in Outlook; if you just have one, you can comment out the two lines below
send_from = 'somebody@company.com'
mail._oleobj_.Invoke(*(64209, 0, 8, 0, outlookNS.Accounts.Item(send_from)))

mail.To = 'somebody@company.com'
#mail.CC = 'somebody@company.com'

mail.Subject = 'Sample Email'

mail.HTMLBody = '<h3>This is an HTML Body.</h3>'
#mail.Body = "This is a plain text Body."

#mail.Attachments.Add('c:\\sample.xlsx')
#mail.Attachments.Add('c:\\sample2.xlsx')

mail.Display()

#mail.Send() 
