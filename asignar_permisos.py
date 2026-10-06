from database import SessionLocal
from models import Usuario, Categoria

def configurar_accesos():
    db = SessionLocal()
    try:
        # 1. Configurar Administradores
        admins = ["martin_graneros@hotmail.com", "pablodgargiulo.laboral@gmail.com"]
        for email in admins:
            user = db.query(Usuario).filter(Usuario.email == email).first()
            if user:
                user.es_admin = True
                print(f"✅ Permisos de administrador otorgados a: {email}")
            else:
                print(f"⚠️ No se encontró al usuario admin: {email}")

        # 2. Configurar Estudio de Prueba (Asignación al Titular)
        titular_email = "drmaximilianosegon@gmail.com"
        modulo_nombre = "Accidentes de Tránsito"

        titular = db.query(Usuario).filter(Usuario.email == titular_email).first()
        modulo_transito = db.query(Categoria).filter(Categoria.nombre == modulo_nombre).first()

        if titular and modulo_transito:
            if modulo_transito not in titular.modulos_activos:
                titular.modulos_activos.append(modulo_transito)
                print(f"✅ Módulo '{modulo_nombre}' asignado al titular: {titular_email}")
            else:
                print(f"✅ El titular {titular_email} ya tenía el módulo asignado.")
        else:
            print(f"⚠️ No se encontró al titular {titular_email} o el módulo '{modulo_nombre}'.")

        db.commit()
        print("🚀 ¡Configuración de prueba finalizada con éxito!")
    except Exception as e:
        print(f"❌ Ocurrió un error: {e}")
        db.rollback()
    finally:
        db.close()

if __name__ == "__main__":
    configurar_accesos()