#!/usr/bin/python

import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email import Encoders
import os
import sys

GMAIL_USER = 'username@gmail.com'
GMAIL_PASSWD = "password"
RECIPIENT = "backup@gmail.com"
SUBJECT = "Writer's Tools document backup"
MESSAGE = "A backup version of an OpenOffice.org document is attached."

def mail(to, subject, text, attach):
   msg = MIMEMultipart()

   msg['From'] = GMAIL_USER
   msg['To'] = to
   msg['Subject'] = subject

   msg.attach(MIMEText(text))

   part = MIMEBase('application', 'octet-stream')
   part.set_payload(open(attach, 'rb').read())
   Encoders.encode_base64(part)
   part.add_header('Content-Disposition',
           'attachment; filename="%s"' % os.path.basename(attach))
   msg.attach(part)

   mailServer = smtplib.SMTP("smtp.gmail.com", 587)
   mailServer.ehlo()
   mailServer.starttls()
   mailServer.ehlo()
   mailServer.login(GMAIL_USER, GMAIL_PASSWD)
   mailServer.sendmail(GMAIL_USER, to, msg.as_string())
   mailServer.close()

mail(RECIPIENT, SUBJECT, MESSAGE, sys.argv[1])
