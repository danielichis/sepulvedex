cyear = cfdate[6:10] 
cmonth = cfdate[3:5]
cday = cfdate[0:2]
cdatemayor=cday+"."+cmonth+"."+cyear
des_month = read_excel.month_sel(cyear,cmonth,cday)
filename = PATH_FILEBANK+"Conciliacion Occidente "+des_month+" "+cyear+".xlsx"
print(filename)
remitente = "bloomcker@grupovenado.com"
destinatario = EMAILDESTI #"jaimemarston@gmail.com"
#destinatario = "eunicequisbert@grupovenado.com;felipevillarroel@grupovenado.com;rolandocespedes@grupovenado.com;ximenabautista@grupovenado.com"
mensaje = "Adjunto se envía conciliaciones bancarias al "+cfdate
message = MIMEMultipart()
message["From"] = remitente
message["To"] = destinatario
message["Subject"] = "Conciliaciones Bancarias"
#email.set_content(mensaje)
message.attach(MIMEText(mensaje, 'plain'))


with open(filename, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
encoders.encode_base64(part)

# Add header as key/value pair to attachment part
part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
)

# Add attachment to message and convert message to string
message.attach(part)
text = message.as_string()
smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
smtp.starttls()
smtp.login(remitente, "Boot123*")
smtp.sendmail(remitente, destinatario, text)
smtp.quit()