################# SENDING EMAIL WITH ATTACHMENT FROM SPECIFIC SENDER #################
#pip install imbox
import glob
import os
from imbox import Imbox
import traceback

host = "imap-mail.outlook.com"
username = "myemail@hotmail.com"
password = "mypassword"
download_folder = "mydirectory"

if not os.path.isdir(download_folder):
    os.makedirs(download_folder, exist_ok=True)

mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
messages = mail.messages(sent_from='specific_sender@email.com',unread = True)

for (uid, message) in messages:
    #mark latest email as read so it won't be considered unread in the next run of the script
    mail.mark_seen(uid) 
    for idx, attachment in enumerate(message.attachments):
        try:
            att_fn = attachment.get('filename')
            download_path = download_folder + "/" + att_fn
            print(download_path)
            with open(download_path, "wb") as fp:
                fp.write(attachment.get('content').read())
        except:
            print(traceback.print_exc())

file_pattern = "/your/pattern/of/choice/*.filetype"
latest_attachment = next(glob.iglob(file_pattern, recursive=True))

import smtplib
from os.path import basename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

msg = MIMEMultipart()
receiver = "receiver@email"
subject = "subject"
content = "content"

msg['From'] = username
msg['To'] = receiver
msg['Subject'] = subject
body = MIMEText(content, 'plain')
msg.attach(body)

part = MIMEApplication(attachment.get('content').read(), Name=basename(latest_attachment))
part['Content-Disposition'] = 'attachment; filename="{}"'.format(basename(latest_attachment))
msg.attach(part)

server = smtplib.SMTP("smtp.office365.com",587)

server.ehlo()
server.starttls()

server.login(username, password)

receivers = ["receiver1@email", "receiver2@email", "receiver3@email"]
for receiver in receivers :
    server.send_message(msg, from_addr = username, to_addrs = [receiver])

################# CONNECTING TO GOOGLE SHEETS #################
#pip install gspread
#pip install oauth2client

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
from datetime import date

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('json_file',scope)
client = gspread.authorize(creds)
spreadsheet = client.open('spreadsheet_name')
sheet = spreadsheet.worksheet('sheet_name')


for record in sheet.get_all_records():
    #matching record to a specific date
    if record['Date'] == (date.today() + datetime.timedelta(days = 0)).strftime("%d %b"): 
        var1 = record['var1']
        var2 = record['var2']
        var3 = record['var3']
        var4 = record['var4']
        var5 = record['var5']

################# SENDING THE WHATSAPP MESSAGE WITH INFO FROM GOOGLE SHEETS #################
import pywhatkit

template = """
Variable 1 is {}
Variable 2 is {}
Variable 3 is {}
Variable 4 is {}
Variable 5 is {}
""".format(var1, var2, var3, var4, var5)

#Specify message time here... ridiculous I can't send instantly lol
#maybe get current hour and minutes, send +1 minute
#for now, 12am is the scheduled sending time
pywhatkit.sendwhatmsg_to_group('Ch5pKpaMfodLDuAY88x80k',template,time_hour = 0,time_min = 0, wait_time = 7, tab_close = True, close_time = 2)