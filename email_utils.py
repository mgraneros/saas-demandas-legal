import os
import resend
from dotenv import load_dotenv

load_dotenv(override=True)

# Configurar la clave API de Resend
resend.api_key = os.getenv("RESEND_API_KEY")

def enviar_correo(destinatario: str, asunto: str, contenido_html: str, ruta_adjunto: str = None):
    """
    Función síncrona para enviar correos electrónicos mediante la API HTTP de Resend.
    Soporta el envío de archivos adjuntos (ej. .docx).
    Diseñada para saltar los bloqueos SMTP y ejecutarse en segundo plano.
    """
    if not resend.api_key:
        print("⚠️ [EMAIL] La API Key de Resend (RESEND_API_KEY) no está configurada en el archivo .env")
        return

    try:
        # Construir la estructura principal del correo
        parametros_correo = {
            "from": "SaaS Legal <onboarding@resend.dev>",
            "to": [destinatario],
            "subject": asunto,
            "html": contenido_html
        }

        # Lógica para procesar e inyectar el archivo adjunto si existe
        if ruta_adjunto and os.path.exists(ruta_adjunto):
            nombre_archivo = os.path.basename(ruta_adjunto)
            with open(ruta_adjunto, "rb") as f:
                # El SDK de Resend en Python requiere transformar los bytes en una lista de enteros
                contenido_bytes = list(f.read())
            
            parametros_correo["attachments"] = [
                {
                    "filename": nombre_archivo,
                    "content": contenido_bytes
                }
            ]
        elif ruta_adjunto:
            print(f"⚠️ [EMAIL] No se encontró el archivo adjunto en la ruta local: {ruta_adjunto}")

        # Ejecutar el disparo mediante la API
        respuesta = resend.Emails.send(parametros_correo)
        print(f"📧 [EMAIL API] Correo enviado exitosamente a: {destinatario} | ID Resend: {respuesta.get('id', 'Desconocido')}")

    except Exception as e:
        print(f"❌ [EMAIL API] Error al enviar el correo a {destinatario}: {str(e)}")