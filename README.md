# Gear-1
This tool download the attachment from email, extract the IP address in the attachment and run it with HakiChekcer.py for reputation check and email the result to a recipient.

# Breakdown
1) Download the excel attachment in the email
2) Extract the IP addresses in the attachment
3) Delete duplicate IP addresses in the excel
4) Give the excel a IP reputation check with HakiChecker.py
5) Get the result and send email to CDOCOps

#Configuration
1) Inbox = Outlook.Folders **(" ")** .Folders.Item("Inbox") 
