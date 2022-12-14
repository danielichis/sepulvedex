import smtplib
from email.message import EmailMessage
import ssl
from pathlib import Path
import os
import sys

class sendEmails:
    def __init__(self):
        self.msg = EmailMessage()
        self.msg['Subject'] = 'Kardexs confrontados'
        self.msg['From'] = 'Bot Sepúlveda'
        self.msg['To'] = 'dchaconb@uni.pe', 
        self.currentPathFolder = getCurrentPath()


    def send(self, docx):
        docxPath = os.path.join(self.currentPathFolder, docx)
        with open(docxPath, 'rb') as f:
            file_data = f.read()
            file_name = f.name
            self.msg.add_attachment(file_data, maintype = 'application', subtype = 'docx', filename = file_name)

        context = ssl.create_default_context()

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context= context) as smtp:
            smtp.login("sepulvedabot0@gmail.com", "fwmjcpowfrjiautv")
            smtp.send_message(self.msg)
            smtp.quit()
       
def getCurrentPath():   
    config_name = 'myapp.cfg'
    # determine if application is a script file or frozen exe
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    application_path2 = Path(application_path)
    return application_path2.absolute()


if __name__=='__main__':
    x = sendEmails()
    x.send('K42218-2.docx')