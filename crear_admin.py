from database import SessionLocal
from models import Usuario
from passlib.context import CryptContext

# Usamos el mismo contexto de encriptación que el resto de tu app
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

db = SessionLocal()

email_admin = "admin@demandas.com"
admin_existente = db.query(Usuario).filter(Usuario.email == email_admin).first()

# Si ya existía pero con la contraseña vieja/incorrecta, lo borramos para recrearlo bien
if admin_existente:
    db.delete(admin_existente)
    db.commit()

# Creamos el hash seguro de la contraseña
hashed_password = pwd_context.hash("admin123")

nuevo_admin = Usuario(
    email=email_admin,
    hashed_password=hashed_password,
    activo=True,
    es_admin=True  # 📌 Con permisos de administrador
)

db.add(nuevo_admin)
db.commit()
db.close()

print("✅ ¡Administrador recreado y hasheado correctamente!")
print("📧 Email: admin@demandas.com")
print("🔑 Contraseña: admin123")