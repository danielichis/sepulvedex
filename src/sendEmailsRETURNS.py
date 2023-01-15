from Google import Create_Service
import base64
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from os import getcwd, path
from utilities import pathsManager
from googleSheet import get_data_from_gsheet
def sendEmailApi(sents):
    kardex=sents["kardex"]
    email=sents["correo"]
    current_parent_path = getcwd()
    CLIENT_SECRET_FILE = path.join(current_parent_path, 'src', 'BOT CONFRONT.json')
    API_NAME = 'gmail'
    API_VERSION = 'v1'
    SCOPES = ['https://mail.google.com/']
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    attachment_docx = path.join(current_parent_path,"KardexsOut", kardex+".docx")
    mimeMessage = MIMEMultipart()
    mimeMessage['to'] = email
    mimeMessage['subject'] = kardex+"-BOT"
    mimeMessage.attach(MIMEText(kardex+"-BOT", 'plain'))
    with open(attachment_docx, 'rb') as f:
        file_data = f.read()
        file_name = kardex+"-BOT"+".docx"
        attach_file=MIMEApplication(file_data)
        attach_file.add_header('Content-Disposition', 'attachment', filename=file_name)
        mimeMessage.attach(attach_file)

    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
    message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
    print(f"correo enviado a {email} con el kardex {kardex}")

def senEmailsApi(listSents):
    for sent in listSents:
        sendEmailApi(sent)
        #print(sent)
if __name__=='__main__':
    listSents=get_data_from_gsheet()
    #print(listSents)
    senEmailsApi(listSents)