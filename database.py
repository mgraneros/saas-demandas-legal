import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base
from dotenv import load_dotenv

# El override=True OBLIGA a Python a leer tu archivo .env actualizado, ignorando cachés
load_dotenv(override=True)

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./historial.db")

# Normalizar 'postgres://' a 'postgresql://'
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# Print de diagnóstico para la consola
print(f"🚀 [DEBUG DB] CONECTANDO A: {DATABASE_URL}")

# Separamos la configuración según el motor de base de datos
if "sqlite" in DATABASE_URL:
    engine = create_engine(
        DATABASE_URL,
        connect_args={"check_same_thread": False}
    )
else:
    engine = create_engine(
        DATABASE_URL,
        pool_pre_ping=True,  # Verifica que la conexión con Neon siga viva
        pool_recycle=300     # Renueva conexiones inactivas cada 5 minutos
    )

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()