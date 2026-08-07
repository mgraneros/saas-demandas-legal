from database import SessionLocal
from models import Usuario
from passlib.context import CryptContext

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

db = SessionLocal()

email_admin = "admin@demandas.com"
admin_existente = db.query(Usuario).filter(Usuario.email == email_admin).first()

hashed_password = pwd_context.hash("admin123")

if admin_existente:
    # Si ya existe, actualizamos sus datos sin romper el historial/demandas asociadas
    admin_existente.hashed_password = hashed_password
    admin_existente.activo = True
    admin_existente.es_admin = True
    print("🔄 El usuario administrador ya existía, se han actualizado sus credenciales y permisos.")
else:
    # Si no existe, lo creamos de cero
    nuevo_admin = Usuario(
        email=email_admin,
        hashed_password=hashed_password,
        activo=True,
        es_admin=True
    )
    db.add(nuevo_admin)
    print("✨ ¡Nuevo administrador creado con éxito!")

db.commit()
db.close()

print("📧 Email: admin@demandas.com")
print("🔑 Contraseña: admin123")