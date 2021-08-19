import pandas as pd 
import smtplib
from email.message import EmailMessage
import glob

e=pd.read_excel("")
emails= e['Emails'].values

print(emails)


for email in emails:
    msg=EmailMessage()
    msg['Subject']=""
    msg['From']=""
    msg["To"]=email
    with open("") as myfile:
        data=myfile.read()
        msg.set_content(data)

    for files in glob.glob("*.xlsx"):
        with open(files,"rb") as f:
            file_data=f.read()
            file_name=f.name
            msg.add_attachment(file_data,maintype="application",subtype="xlsx",filename=file_name)

    server=smtplib.SMTP("smtp.gmail.com",587)
    server.starttls()
    server.login()
    server.send_message(msg)
    server.quit()

print("Email Sent")