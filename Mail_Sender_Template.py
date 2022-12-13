#edited version by @Ashok Vavare
import os
import cv2
import pandas as pd

import smtplib
import mimetypes
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage


def mime_init(from_addr, recipients_addr, subject, body):
    """
    :param str from_addr:           The email address you want to send mail from
    :param list recipients_addr:    The list of email addresses of recipients
    :param str subject:             Mail subject
    :param str body:                Mail body
    :return:                        MIMEMultipart object
    """

    msg = MIMEMultipart()

    msg['From'] = from_addr
    msg['To'] = ','.join(recipients_addr)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    return msg


def send_email(user, password, from_addr, recipients_addr, subject, body, files_path):
    """
    :param str user:                Sender's email userID
    :param str password:            sender's email password
    :param str from_addr:           The email address you want to send mail from
    :param list recipients_addr:    List of (or space separated string) email addresses of recipients
    :param str subject:             Mail subject
    :param str body:                Mail body
    :param list files_path:         List of paths of files you want to attach
    :param str server:              SMTP server (port is choosen 587)
    :return:                        None
    """

    #   assert isinstance(recipents_addr, list)
    FROM = from_addr
    TO = recipients_addr
    PASS = password
    SERVER = 'smtp.gmail.com'
    SUBJECT = subject
    BODY = body
    msg = mime_init(FROM, TO, SUBJECT, BODY)
    pdfname = file_path

# open the file in bynary
    binary_pdf = open(pdfname, 'rb')

    payload = MIMEBase('application', 'octate-stream', Name=pdfname)
    payload.set_payload((binary_pdf).read())

# enconding the binary into base64
    encoders.encode_base64(payload)

# add header with pdf name
    payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
    msg.attach(payload)
    server = smtplib.SMTP(SERVER, 587)
    server.starttls()
    # Enter login credentials for the email you want to sent mail from
    server.login(user, PASS)
    text = msg.as_string()

    server.sendmail(FROM, TO, text)
    print("Sent Successfully!!! " + recipients_addr)

    server.quit()


if __name__ == "__main__":
    user = '*************@gmail.com'         # Email userID
    password = '16_CHARACTRES_FOR_PASSWORD_FROM_GOOGLE_2FA'      # Email password
    from_addr = '**********@gmail.com' //same as above user mail.
    data = pd.read_excel("expert.xlsx")
    names = list(data.Name)
    emails = list(data.Email)

    for index in range(len(data)):
        reciever = emails[index]
        subject = 'Codvert Certificate'
        body = 'Thank you for participating in CODEVERT.\n Please find attached certificate of participation.\n Do participate in future events\n\n \n Regards, Team Codevert'
        file_path = "EXPERT" + "\\" + names[index] + ".pdf"


        send_email(user, password, from_addr,
                   reciever, subject, body, file_path)
