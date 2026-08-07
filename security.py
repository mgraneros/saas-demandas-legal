import os
from dotenv import load_dotenv
from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from sqlalchemy.orm import Session
from jose import JWTError, jwt
import models
from database import get_db

# Forzamos la lectura del archivo .env y sobrescribimos cualquier variable previa del sistema
load_dotenv(override=True)

SECRET_KEY = os.getenv("SECRET_KEY", "admin123")
print(f"DEBUG - SECRET_KEY ACTIVA EN SECURITY.PY: {SECRET_KEY}")
ALGORITHM = "HS256"

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")

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

def verificar_suscripcion_activa(current_user: models.Usuario = Depends(get_current_user)):
    """
    Verifica que el usuario tenga suscripción activa. 
    Los administradores tienen acceso libre por defecto.
    """
    # Si tu campo de administrador se llama 'is_admin' o similar, ajustalo acá:
    if getattr(current_user, "es_admin", False):
        return current_user

    suscripcion_activa = getattr(current_user, "suscripcion_activa", False)

    if not suscripcion_activa:
        raise HTTPException(
            status_code=status.HTTP_402_PAYMENT_REQUIRED,
            detail="Acceso restringido. Debes abonar la suscripción para utilizar esta funcionalidad."
        )
    return current_user