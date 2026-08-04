import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

load_dotenv()

SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_FROM = os.getenv("SMTP_FROM")

def enviar_correo(destinatario: str, asunto: str, contenido_html: str):
    """
    Función síncrona para enviar correos electrónicos mediante SMTP.
    Diseñada para ejecutarse en segundo plano con FastAPI BackgroundTasks.
    """
    if not SMTP_USER or not SMTP_PASSWORD:
        print("⚠️ [EMAIL] Las credenciales SMTP no están configuradas en el archivo .env")
        return

    try:
        # Configurar mensaje
        mensaje = MIMEMultipart("alternative")
        mensaje["Subject"] = asunto
        mensaje["From"] = SMTP_FROM
        mensaje["To"] = destinatario

        # Adjuntar contenido HTML
        parte_html = MIMEText(contenido_html, "html", "utf-8")
        mensaje.attach(parte_html)

        # Conectar al servidor SMTP y enviar
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as servidor:
            servidor.starttls()
            servidor.login(SMTP_USER, SMTP_PASSWORD)
            servidor.sendmail(SMTP_FROM, destinatario, mensaje.as_string())
            
        print(f"📧 [EMAIL] Correo enviado exitosamente a: {destinatario}")

    except Exception as e:
        print(f"❌ [EMAIL] Error al enviar el correo a {destinatario}: {str(e)}")