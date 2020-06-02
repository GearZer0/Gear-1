# GearI
This tool download the attachment from email, extract the IP address in the attachment and run it with HakiChekcer.py for reputation check and email the result to a recipient.

# Breakdown
1. Download the excel attachment in the email
2. Extract the IP addresses in the attachment
3. Delete duplicate IP addresses in the excel
4. Give the excel a IP reputation check with HakiChecker.py
5. Get the result and send email to recipient

# Configuration
Fill in the following in Gear-1 :
````
1. Inbox = Outlook.Folders(" ").Folders.Item("Inbox")     *Input your mailbox name*
2. file_name = " {}".format(today)                        *Input the subject name of the email*
3. _file_name = "% {}".format(_today)
4. mail.To = ' '                                          *Input email address of the recipient*
5. mail.Subject = ' '                                     *Input email subject name*
6. mail.Body = ' '                                        *Input body of the email*
7. mail.SentOnBehalfOfName = ' '                          *Input email address of the sender*
````
# Insruction to run this tool
Place tool.py together in the same folder as Hakichecker

# Command to run this tool
python tool.py

# Create the Batch File
Create a Notepad file and input the following into the notepad :
```
cd "Path where your HakiChecker is"
python tool.py
pause
```
Save the Notepad with .bat
