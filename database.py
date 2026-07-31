import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

# URL de la base de datos (SQLite local)
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./historial.db")

engine = create_engine(
    DATABASE_URL, 
    connect_args={"check_same_thread": False} if "sqlite" in DATABASE_URL else {}
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

# 👈 Esta es la función que te faltaba:
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def inicializar_db():
    """Crea todas las tablas definidas en los modelos."""
    import models  # Carga los modelos de SQLAlchemy
    Base.metadata.create_all(bind=engine)