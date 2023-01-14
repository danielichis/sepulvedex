from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from utilities import pathsManager as pm
from os.path import join

# Establecer credenciales
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
credentials_json_path = join(pm().currentFolderPath, 'src', 'BOT CONFRONT.json')
credentials = InstalledAppFlow.from_client_secrets_file(credentials_json_path, SCOPES).run_local_server()

# Crear servicio de Gmail
service = build('gmail', 'v1', credentials=credentials)

# Crear el mensaje
message = MIMEText('Este es el contenido del mensaje')
message['to'] = 'cllerenam@uni.pe'
message['from'] = 'sepulvedabot0@gmail.com'
message['subject'] = 'Probando gmail api'

# Enviar el mensaje
message = {'raw': message.as_string()}
service.users().messages().send(userId='me', body=message).execute()
print('hola')