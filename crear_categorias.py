from database import SessionLocal, engine
from models import Base, Categoria, Plantilla
from sqlalchemy import text

def poblar_base_datos():
    print("⏳ Verificando/Creando tablas faltantes en PostgreSQL...")
    Base.metadata.create_all(bind=engine)
    
    db = SessionLocal()
    try:
        # 1. Inyectar la columna faltante en la tabla existente mediante SQL puro
        print("⏳ Agregando columna categoria_id a la tabla plantillas (si no existe)...")
        db.execute(text("ALTER TABLE plantillas ADD COLUMN IF NOT EXISTS categoria_id INTEGER REFERENCES categorias(id)"))
        db.commit()

        # 2. Crear las categorías
        modulos = ["Accidentes de Tránsito", "ART"]
        
        for nombre in modulos:
            categoria_existente = db.query(Categoria).filter(Categoria.nombre == nombre).first()
            if not categoria_existente:
                nueva_categoria = Categoria(nombre=nombre, activa=True)
                db.add(nueva_categoria)
                print(f"✅ Módulo creado: {nombre}")
            else:
                print(f"✅ El módulo '{nombre}' ya existe en la base de datos.")
        
        db.commit()

        # 3. Vincular las plantillas existentes al módulo de Tránsito
        cat_transito = db.query(Categoria).filter(Categoria.nombre == "Accidentes de Tránsito").first()
        
        if cat_transito:
            plantillas_huerfanas = db.query(Plantilla).filter(Plantilla.categoria_id == None).all()
            for plantilla in plantillas_huerfanas:
                plantilla.categoria_id = cat_transito.id
                print(f"🔗 Plantilla '{plantilla.nombre}' vinculada al módulo 'Accidentes de Tránsito'.")
            
            db.commit()
            
        print("🚀 ¡Base de datos actualizada y poblada con éxito!")

    except Exception as e:
        print(f"❌ Ocurrió un error: {e}")
        db.rollback()
    finally:
        db.close()

if __name__ == "__main__":
    poblar_base_datos()