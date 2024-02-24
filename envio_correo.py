import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os

def enviar_correo():
    # Configuración del servidor y credenciales
    smtp_server = 'smtp.gmail.com'  # Cambia esto al servidor SMTP que estés utilizando
    smtp_port = 587  # Cambia esto al puerto adecuado
    sender_email = 'jeam.mendoza.melo@gmail.com'  # Cambia esto a tu dirección de correo electrónico
    sender_password = 'nwzi udhw bpbj kjzu'  # Cambia esto a tu contraseña

    # Detalles del correo electrónico
    receiver_email = 'anthony.mcg24@gmail.com'  # Cambia esto al destinatario deseado
    subject = 'Reportes Reactiva: Top 5 Costo de Inversión para Junin y Apurimac'
    body = 'Adjunto los reportes de Reactiva'

    # Crear el objeto MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Adjuntar archivos
    file_paths = ['top5_inversion_JUNIN.xlsx', 'top5_inversion_APURIMAC.xlsx']
    for file_path in file_paths:
        with open(file_path, 'rb') as file:
            attachment = MIMEApplication(file.read(), _subtype="xlsx")
            attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
            msg.attach(attachment)

    # Iniciar la conexión con el servidor SMTP
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()  # Iniciar el modo seguro
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())

    print('Correo enviado exitosamente')

enviar_correo()
