import smtplib
from email.message import EmailMessage
import ssl

class sendEmails:
    def __init__(self):
        self.msg = EmailMessage()
        self.msg['Subject'] = 'Kardexs confrontados'
        self.msg['From'] = 'Bot Sepúlveda'
        self.msg['To'] = 'dchaconb@uni.pe'

    def send(self):
        with open(r'C:\Users\crist\Documents\K42218-2.docx', 'rb') as f:
            file_data = f.read()
            file_name = f.name
            self.msg.add_attachment(file_data, maintype = 'application', subtype = 'docx', filename = file_name)

        context = ssl.create_default_context()

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context= context) as smtp:
            smtp.login("cristhiamllerena@gmail.com", "eqgodczcyagtbiqz")
            smtp.send_message(self.msg)
            smtp.quit()
       
