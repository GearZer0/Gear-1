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

def sendEmail(filename):
    email_username = input("Enter outlook email: ") # set sender email
    email_pwd = input("Enter outlook password: ") # set sender password
    receiver_url = input("Enter receipient email: ") #set receiver email
    
    SourcePathName  = os.getcwd() + "/" + filename 

    msg = MIMEMultipart()
    msg['From'] = email_username
    msg['To'] = receiver_url
    msg['Subject'] = 'Report Update'
    body = 'Report File'
    msg.attach(MIMEText(body, 'plain'))

    ## ATTACHMENT PART OF THE CODE IS HERE
    attachment = open(SourcePathName, 'rb')
    part = MIMEBase('application', "octet-stream")
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)

    server = smtplib.SMTP('imap-mail.outlook.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(email_username, email_pwd)
    server.send_message(msg)
    server.quit()
    print("Email sent ...")

if __name__ == "__main__":
    today = datetime.now().strftime("%d%m%Y")
    file_name = "Daily Report {}.xlsx".format(today)
    wb = xlrd.open_workbook(file_name)
    sheet = wb.sheet_by_index(0)
    all_ips = []
    for i in range(sheet.nrows):
        cell_data = sheet.cell_value(i,4)
        IP = re.findall('[\d]+.[\d]+.[\d]+.[\d]', cell_data)
        if len(IP):
            IP = IP[0]
            all_ips.append(IP)
    all_ips = list(set(all_ips))
    if os.path.exists("tmp.txt"):
        os.remove("tmp.txt")
    if os.path.exists("Results.zip"):
        os.remove("Results.zip")
    with open('tmp.txt', 'a+') as ip_file:
        for ip in all_ips:
            ip_file.write(ip + "\n")
    print("Running command ... please wait for output to populate shortly ...")
    run_bot = subprocess.Popen('py Checker.py -ip tmp.txt'.split(' ')).wait()
    while True:
        sleep(1)
        files = os.listdir("Results")
        if len(files) > 0:
            #files = sorted(filter(os.path.isfile, os.listdir('Results')), key=os.path.getmtime)
            sendEmail("Results/" + files[0])
            break
