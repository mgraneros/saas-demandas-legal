from database import engine, SessionLocal
from models import Base, Plantilla

# Crea las tablas en la base de datos si aún no existen
Base.metadata.create_all(bind=engine)

db = SessionLocal()

plantilla_existente = db.query(Plantilla).filter(Plantilla.nombre == "Demanda Tránsito Auto/Moto").first()

if not plantilla_existente:
    nueva_plantilla = Plantilla(
        nombre="Demanda Tránsito Auto/Moto",
        categoria="Daños y Perjuicios",
        descripcion="Plantilla automatizada para accidentes de tránsito entre auto y moto.",
        ruta_archivo="templates/Borrador_Demanda_Auto_Moto.docx",
        activa=True
    )
    db.add(nueva_plantilla)
    db.commit()
    print("✅ Plantilla inicial cargada correctamente.")
else:
    print("ℹ️ La plantilla ya existe en la base de datos.")

db.close()