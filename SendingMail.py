import csv
from random import random
import smtplib
import ssl
import time
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def ReadCSV(FileName):
    data_list = []
    with open(f"{FileName}", 'r') as file:
        csvreader = csv.reader(file)
        for row in csvreader:
            data_list.append(row)
    return data_list

def main():
    FileName = input("Enter the File with Details: ")
    email = input("Enter your Email: ")
    password = 'tvbprmvgogbjwgmj'
    datas = ReadCSV(FileName)
    for data in datas:
        email_sender = f'{email}'
        email_password = f'{password}'
        email_receiver = data[1]
        subject = 'This is a Test Mail'
        body = """
            Hello !!
            """
        em = MIMEMultipart()
        em['From'] = email_sender
        em['To'] = email_receiver
        em['Subject'] = subject
        #em.set_content(body)
        context = ssl.create_default_context()
        attach_file_name = input("Enter the Attachment Path: ")
        if attach_file_name == "":
            em.attach(MIMEText(body, 'plain'))
            pass
        else:
            em.attach(MIMEText(body, 'plain'))
            attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
            payload = MIMEBase('application', 'octate-stream')
            payload.set_payload((attach_file).read())
            encoders.encode_base64(payload) #encode the attachment
            #add payload header with filename
            payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
            em.attach(payload)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(email_sender, email_password)
            smtp.sendmail(email_sender, email_receiver, em.as_string())
        time.sleep(random.randrange(1,10))
    print("The Task is complete")

if __name__=="__main__":
    main()