# -*- coding: utf-8 -*-
import xlrd # pip install xlrd
from datetime import datetime
import re
import subprocess
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from time import sleep
import win32com.client

def downloadAttach():
    Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#    olNs = Outlook.GetNamespace("MAPI")
#    Inbox = olNs.GetDefaultFolder(6)
    Inbox = Outlook.Folders(" ").Folders.Item("Inbox")
    today = datetime.now().strftime("%d %B %Y")
    _today = str(datetime.now().day) + " " + datetime.now().strftime("%B %Y")
    file_name = "%Daily Summary Report {}".format(today)
    _file_name = "%Daily Summary Report {}".format(_today)
    Filter = ("@SQL=" + chr(34) + "urn:schemas:httpmail:subject" +
              chr(34) + " Like '" + file_name + "' AND " +
              chr(34) + "urn:schemas:httpmail:hasattachment" +
              chr(34) + "=1")

    Items = Inbox.Items.Restrict(Filter)
    for Item in Items:
        for attachment in Item.Attachments:
            print(attachment.FileName)
            attachment.SaveAsFile(os.getcwd() + "/Attachment/" + attachment.FileName)
            return "Attachment" + "/" + attachment.FileName
    # repeat again for other date format ...
    Filter = ("@SQL=" + chr(34) + "urn:schemas:httpmail:subject" +
              chr(34) + " Like '" + _file_name + "' AND " +
              chr(34) + "urn:schemas:httpmail:hasattachment" +
              chr(34) + "=1")

    Items = Inbox.Items.Restrict(Filter)
    for Item in Items:
        for attachment in Item.Attachments:
            print(attachment.FileName)
            attachment.SaveAsFile(os.getcwd() + "/Attachment/" + attachment.FileName)
            return "Attachment" + "/" + attachment.FileName

def sendEmail(filename):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ' '
    mail.Subject = ' '
    mail.Body = ' '
#   mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    attachment  = filename
    mail.Attachments.Add(attachment)

    mail.SentOnBehalfOfName = ' '
    mail.Send()
    print("Email sent ...")

if __name__ == "__main__":
    print("Downloading attachment")
    try:
        os.mkdir("Attachment")
    except:
        pass
    
    #today = datetime.now().strftime("%Y%m%d")
    file_name = downloadAttach()
    wb = xlrd.open_workbook(file_name)
    sheet = wb.sheet_by_index(0)
    all_ips = []
    for i in range(sheet.nrows):
        cell_data = sheet.cell_value(i,4)
        IP = re.findall('[\d]+.[\d]+.[\d]+.[\d]+', cell_data)
        if len(IP):
            IP = IP[0]
            all_ips.append(IP)
    all_ips = list(set(all_ips))
    if os.path.exists("tmp.txt"):
        os.remove("tmp.txt")
    with open('tmp.txt', 'a+') as ip_file:
        for ip in all_ips:
            ip_file.write(ip + "\n")
    print("Running command ... please wait for output to populate shortly ...")
    run_bot = subprocess.Popen('python HakiChecker.py -ip tmp.txt'.split(' ')).wait()
    while True:
        sleep(1)
        files = os.listdir("Results")
        if len(files) > 0:
            #files = sorted(filter(os.path.isfile, os.listdir('Results')), key=os.path.getmtime)
            sendEmail(os.getcwd() + "/Results/" + files[0])
            break
