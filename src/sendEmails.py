import smtplib
from email.message import EmailMessage
import ssl
from pathlib import Path
import os
import sys
from googleSheet import get_data_from_gsheet

class sendEmails:
    def __init__(self):
        self.msg = EmailMessage()
        # send boy message
        self.msg.set_content('This is a plain text email')   
        self.currentPathFolder = getCurrentPath()

    def sendEmails(self,listOfSents):
        for sent in listOfSents:
            self.msg['Subject'] = 'Kardexs confrontados'
            self.msg['From'] = 'Bot Sepúlveda'
            self.msg['To'] = sent["correo"]
            docxPath = os.path.join(self.currentPathFolder, "kardexsOut", sent["kardex"])
            # check if the file exists
            if not os.path.isfile(docxPath):
                messageEmail="El kardex {} no pudo ser procesado".format(sent["kardex"])
                with open(docxPath, 'rb') as f:
                    file_data = f.read()
                    file_name = f.name
                    self.msg.add_attachment(file_data, maintype = 'application', subtype = 'docx', filename = sent["kardex"])
            else:
                messageEmail="El kardex {} fue procesado".format(sent["kardex"])
            self.msg.set_content(messageEmail)
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


#if __name__=='__main__':
x = sendEmails()
ks=get_data_from_gsheet()
x.sendEmails(ks)