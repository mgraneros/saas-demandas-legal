import os
import datetime
from google.cloud import storage
from dotenv import load_dotenv

# Cargamos las variables del .env
load_dotenv()

# Obtenemos el nombre de tu bucket desde el .env
BUCKET_NAME = os.getenv("GCP_BUCKET_NAME")

def get_storage_client():
    """
    Inicializa y retorna el cliente de Google Cloud Storage.
    Al tener GOOGLE_APPLICATION_CREDENTIALS en el .env,
    esta librería encuentra tu llave .json automáticamente.
    """
    return storage.Client()

def upload_to_gcp(local_file_path: str, destination_blob_name: str) -> str:
    """
    Sube un archivo local a Google Cloud Storage.
    Retorna el nombre del archivo tal como quedó guardado en la nube.
    """
    client = get_storage_client()
    bucket = client.bucket(BUCKET_NAME)
    
    # Creamos un "blob" (un espacio para el archivo) en la nube
    blob = bucket.blob(destination_blob_name)

    # Subimos el archivo físico
    blob.upload_from_filename(local_file_path)
    
    return destination_blob_name

def generate_signed_url(blob_name: str, expiration_minutes: int = 15) -> str:
    """
    Genera una URL temporal y ultra-segura para descargar el archivo.
    Expira automáticamente protegiendo la privacidad de la demanda.
    """
    client = get_storage_client()
    bucket = client.bucket(BUCKET_NAME)
    blob = bucket.blob(blob_name)

    # Genera la URL firmada (V4 es el estándar más seguro)
    url = blob.generate_signed_url(
        version="v4",
        expiration=datetime.timedelta(minutes=expiration_minutes),
        method="GET",
    )
    
    return url

def download_from_gcp(blob_name: str, local_destination: str) -> str:
    """
    Descarga un archivo desde Google Cloud Storage a una ruta temporal local.
    Ideal para traer las plantillas base (.docx) a la memoria del servidor 
    justo antes de inyectarles los datos del formulario.
    """
    client = get_storage_client()
    bucket = client.bucket(BUCKET_NAME)
    blob = bucket.blob(blob_name)

    # Descarga el archivo de la nube a la ruta local especificada
    blob.download_to_filename(local_destination)
    
    return local_destination