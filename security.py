import os
from dotenv import load_dotenv
from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from sqlalchemy.orm import Session
from jose import JWTError, jwt
from passlib.context import CryptContext
from itsdangerous import URLSafeTimedSerializer
from fastapi_mail import ConnectionConfig

import models
from database import get_db

# Forzamos la lectura del archivo .env
load_dotenv(override=True)

# --- VARIABLES DE ENTORNO Y SEGURIDAD ---
SECRET_KEY = os.getenv("SECRET_KEY", "admin123")
ALGORITHM = os.getenv("ALGORITHM", "HS256")
print(f"DEBUG - SECRET_KEY ACTIVA EN SECURITY.PY: {SECRET_KEY}")

# --- CONFIGURACIÓN DE HASHEO Y TOKENS ---
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

# Firmante para generar y validar tokens con expiración en recuperación de contraseña
serializer = URLSafeTimedSerializer(SECRET_KEY)

# --- CONFIGURACIÓN DE CORREO SMTP (FASTAPI-MAIL) ---
mail_config = ConnectionConfig(
    MAIL_USERNAME=os.getenv("SMTP_USER", "tu_correo@gmail.com"),
    MAIL_PASSWORD=os.getenv("SMTP_PASSWORD", "tu_contraseña_de_aplicacion"),
    MAIL_FROM=os.getenv("SMTP_FROM", "tu_correo@gmail.com"),
    MAIL_PORT=int(os.getenv("SMTP_PORT", 587)),
    MAIL_SERVER=os.getenv("SMTP_HOST", "smtp.gmail.com"),
    MAIL_STARTTLS=True,
    MAIL_SSL_TLS=False,
    USE_CREDENTIALS=True,
    VALIDATE_CERTS=True
)

# --- OAUTH2 SCHEME ---
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")


# --- FUNCIONES DE HASHEO DE CONTRASEÑA ---
def hash_password(password: str) -> str:
    return pwd_context.hash(password)


def verify_password(plain_password: str, hashed_password: str) -> bool:
    return pwd_context.verify(plain_password, hashed_password)


# --- DEPENDENCIAS DE AUTENTICACIÓN Y ROLES ---
def get_current_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)):
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="No se pudieron validar las credenciales de autenticación.",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        email: str = payload.get("sub")
        if email is None:
            raise credentials_exception
    except JWTError as e:
        print("❌ Error JWTError al decodificar el token:", str(e))
        raise credentials_exception

    usuario = db.query(models.Usuario).filter(models.Usuario.email == email).first()
    if usuario is None:
        raise credentials_exception
    return usuario


def get_current_admin_user(current_user: models.Usuario = Depends(get_current_user)):
    """
    Verifica que el usuario autenticado tenga permisos de administrador.
    Soporta tanto 'is_admin' como 'es_admin' según cómo se haya creado en la BD.
    """
    es_administrador = getattr(current_user, "is_admin", False) or getattr(current_user, "es_admin", False)

    if not es_administrador:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Acceso denegado. Esta sección es exclusiva para administradores."
        )
    return current_user


def verificar_suscripcion_activa(
    current_user: models.Usuario = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Verifica que el usuario tenga suscripción activa en la base de datos. 
    Los administradores tienen acceso libre por defecto.
    """
    if getattr(current_user, "es_admin", False) or getattr(current_user, "is_admin", False):
        return current_user

    # Buscamos la suscripción real en la tabla Suscripcion
    suscripcion = db.query(models.Suscripcion).filter(
        models.Suscripcion.usuario_id == current_user.id,
        models.Suscripcion.activa == True
    ).first()

    if not suscripcion:
        raise HTTPException(
            status_code=status.HTTP_402_PAYMENT_REQUIRED,
            detail="Acceso restringido. Debes abonar la suscripción para utilizar esta funcionalidad."
        )
    return current_user