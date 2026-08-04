import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication  # Importación necesaria para adjuntos
from dotenv import load_dotenv

load_dotenv()

SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_FROM = os.getenv("SMTP_FROM")

def enviar_correo(destinatario: str, asunto: str, contenido_html: str, ruta_adjunto: str = None):
    """
    Función síncrona para enviar correos electrónicos mediante SMTP.
    Soporta el envío de archivos adjuntos (ej. .docx).
    Diseñada para ejecutarse en segundo plano con FastAPI BackgroundTasks.
    """
    if not SMTP_USER or not SMTP_PASSWORD:
        print("⚠️ [EMAIL] Las credenciales SMTP no están configuradas en el archivo .env")
        return

    try:
        # Configurar mensaje (usamos "mixed" cuando hay adjuntos)
        mensaje = MIMEMultipart("mixed")
        mensaje["Subject"] = asunto
        mensaje["From"] = SMTP_FROM
        mensaje["To"] = destinatario

        # Adjuntar contenido HTML
        parte_html = MIMEText(contenido_html, "html", "utf-8")
        mensaje.attach(parte_html)

        # Lógica para adjuntar archivo si se proporciona la ruta
        if ruta_adjunto and os.path.exists(ruta_adjunto):
            with open(ruta_adjunto, "rb") as f:
                adjunto = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.wordprocessingml.document")
                nombre_archivo = os.path.basename(ruta_adjunto)
                adjunto.add_header("Content-Disposition", "attachment", filename=nombre_archivo)
                mensaje.attach(adjunto)
        elif ruta_adjunto:
            print(f"⚠️ [EMAIL] No se encontró el archivo adjunto en la ruta: {ruta_adjunto}")

        # Conectar al servidor SMTP y enviar
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as servidor:
            servidor.starttls()
            servidor.login(SMTP_USER, SMTP_PASSWORD)
            servidor.sendmail(SMTP_FROM, destinatario, mensaje.as_string())
            
        print(f"📧 [EMAIL] Correo enviado exitosamente a: {destinatario}")

    except Exception as e:
        print(f"❌ [EMAIL] Error al enviar el correo a {destinatario}: {str(e)}")