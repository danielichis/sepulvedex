from Google import Create_Service
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from os import getcwd, path

current_parent_path = getcwd()
CLIENT_SECRET_FILE = path.join(current_parent_path, 'src', 'BOT CONFRONT.json')
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

emailMsg = 'Esto está fucionando de ptm'
attachment_docx = path.join(current_parent_path, 'src', 'daniel.docx')
mimeMessage = MIMEMultipart()
mimeMessage['to'] = 'cllerenam@uni.pe'
mimeMessage['subject'] = 'Prueba bot confront'
with open(attachment_docx, 'rb') as f:
    file_data = f.read()
    file_name = 'lets go.docx'
    mimeMessage.attach(MIMEApplication(file_data, maintype = 'application', subtype = 'docx', filename = file_name))
mimeMessage.attach(MIMEText(emailMsg, 'plain'))
raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
print(message)
