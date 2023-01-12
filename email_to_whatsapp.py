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
from email import message
from os.path
msg = message.Message()
receiver = "receiver@email.com"
subject = "Email Subject"
body = "Email Body"

msg.add_header('from',username)
msg.add_header('to',receiver)
msg.add_header('subject',subject)
msg.set_payload(body)

server = smtplib.SMTP("smtp.example.com",587) 

server.ehlo()
server.starttls()

server.login(username, password)

server.send_message(msg, from_addr = username, to_addrs = [receiver])
