import os
from datetime import datetime, timedelta
from typing import Optional, List
from docx import Document
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage, RichText
from fastapi import BackgroundTasks, Depends, FastAPI, HTTPException, Request, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from jose import JWTError, jwt
from num2words import num2words
from sqlalchemy.orm import Session
import mercadopago
import bcrypt
from dotenv import load_dotenv
import httpx
from models import Usuario as User
from fastapi import FastAPI, Request, HTTPException
from fastapi import BackgroundTasks
from email_utils import enviar_correo
from security import get_current_admin_user, get_current_user
import schemas

# Módulos propios del proyecto
from database import get_db, engine
import models
import schemas  # 👈 Esquemas centralizados en schemas.py

load_dotenv()

# 1. Inicialización de la aplicación FastAPI
app = FastAPI(title="SaaS Demandas Legal API", version="0.3.0")
sdk = mercadopago.SDK(os.getenv("MP_ACCESS_TOKEN"))
# ⚠️ ESTA LÍNEA CREA LAS TABLAS AUTOMÁTICAMENTE EN LA BASE DE DATOS
models.Base.metadata.create_all(bind=engine)
# Agregar el middleware de CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 2. Configuración JWT
SECRET_KEY = "admin123"
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24  # 24 horas

# 3. Esquema OAuth2 para Swagger UI
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")


# 4. Funciones de encriptación y utilidades
def get_password_hash(password: str) -> str:
    pwd_bytes = password.encode('utf-8')
    salt = bcrypt.gensalt()
    hashed = bcrypt.hashpw(pwd_bytes, salt)
    return hashed.decode('utf-8')

def verify_password(plain_password: str, hashed_password: str) -> bool:
    pwd_bytes = plain_password.encode('utf-8')
    hashed_bytes = hashed_password.encode('utf-8')
    return bcrypt.checkpw(pwd_bytes, hashed_bytes)

def create_access_token(data: dict, expires_delta: timedelta | None = None):
    to_encode = data.copy()
    expire = datetime.utcnow() + (expires_delta or timedelta(minutes=15))
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

def formatear_moneda(valor: float) -> str:
    return f"$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def monto_a_letras_legal(monto: float) -> str:
    entero = int(monto)
    decimales = int(round((monto - entero) * 100))
    texto_entero = num2words(entero, lang='es')
    return f"{texto_entero.upper()} CON {decimales:02d}/100"


# 5. Diccionario Modular de Competencias
PARRAFOS_COMPETENCIA = {
    1: (
        "En los supuestos de citación en garantía, la ley de seguros 17.418, en su artículo 118, 2do párrafo indica que la demanda podrá ser interpuesta ante el juez del lugar del hecho o del domicilio del asegurador. Por lo expuesto, V.S. es competente en la materia para entender en estos actuados, ya que el hecho ocurrió en su jurisdicción.\n"
        "En tal sentido debemos señalar que el presente juicio tiene por objeto reclamar los daños y perjuicios derivados de un delito, siendo que el art. 5 inc. 4º del CPCCN establece que “en las acciones derivadas de delitos o cuasidelitos, el del lugar del hecho o el del domicilio del demandado a elección del actor”. Así se ha sostenido que si bien las leyes procesales establecen que en las acciones personales derivadas de delitos o cuasidelitos será juez competente el del lugar del hecho o el del domicilio del demandado, corresponde conocer de la causa al juez del lugar donde se domicilia el asegurador citado en garantía, conforme con la opción que acuerda el art. 118 de la ley 17.418, que legisla sobre seguros para toda la Nación. (Jara Zúñiga, Romiglio. 01/01/74 T. 290, p. 387).-"
    ),
    2: (
        "En los supuestos de citación en garantía, la ley de seguros 17.418, en su artículo 118, 2do párrafo indica que la demanda podrá ser interpuesta ante el juez del lugar del hecho o del domicilio del asegurador. Por lo expuesto, V.S. es competente en la materia para entender en estos actuados, ya que el domicilio del demandado se encuentra en su jurisdiccion.\n"
        "En tal sentido debemos señalar que el presente juicio tiene por objeto reclamar los daños y perjuicios derivados de un delito, siendo que el art. 5 inc. 4º del CPCCN establece que “en las acciones derivadas de delitos o cuasidelitos, el del lugar del hecho o el del domicilio del demandado a elección del actor”. Así se ha sostenido que si bien las leyes procesales establecen que en las acciones personales derivadas de delitos o cuasidelitos será juez competente el del lugar del hecho o el del domicilio del demandado, corresponde conocer de la causa al juez del lugar donde se domicilia el asegurador citado en garantía, conforme con la opción que acuerda el art. 118 de la ley 17.418, que legisla sobre seguros para toda la Nación. (Jara Zúñiga, Romiglio. 01/01/74 T. 290, p. 387).-"
    ),
    3: (
        "En los supuestos de citación en garantía, la ley de seguros 17.418, en su artículo 118, 2do párrafo indica que la demanda podrá ser interpuesta ante el juez del lugar del hecho o del domicilio del asegurador. Por lo expuesto, V.S. es competente en la materia para entender en estos actuados, ya que el domicilio de la citada en garantía se encuentra en su jurisdiccion.\n"
        "En tal sentido debemos señalar que el presente juicio tiene por objeto reclamar los daños y perjuicios derivados de un delito, siendo que el art. 5 inc. 4º del CPCCN establece que “en las acciones derivadas de delitos o cuasidelitos, el del lugar del hecho o el del domicilio del demandado a elección del actor”. Así se ha sostenido que si bien las leyes procesales establecen que en las acciones personales derivadas de delitos o cuasidelitos será juez competente el del lugar del hecho o el del domicilio del demandado, corresponde conocer de la causa al juez del lugar donde se domicilia el asegurador citado en garantía, conforme con la opción que acuerda el art. 118 de la ley 17.418, que legisla sobre seguros para toda la Nación. (Jara Zúñiga, Romiglio. 01/01/74 T. 290, p. 387).-"
    )
}




# ==========================================
# RUTAS DE USUARIOS Y AUTENTICACIÓN
# ==========================================

@app.post("/register", response_model=schemas.UsuarioResponse, summary="Registrar nuevo usuario")
def registrar_usuario(usuario: schemas.UsuarioCreate, db: Session = Depends(get_db)):
    db_usuario = db.query(models.Usuario).filter(models.Usuario.email == usuario.email).first()
    if db_usuario:
        raise HTTPException(status_code=400, detail="El email ya se encuentra registrado.")
    
    hashed_pwd = get_password_hash(usuario.password)
    nuevo_usuario = models.Usuario(
        email=usuario.email,
        hashed_password=hashed_pwd,
        nombre_estudio=usuario.nombre_estudio
    )
    db.add(nuevo_usuario)
    db.commit()
    db.refresh(nuevo_usuario)
    return nuevo_usuario


@app.post("/token", response_model=schemas.Token, summary="Iniciar sesión y obtener JWT")
def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    usuario = db.query(models.Usuario).filter(models.Usuario.email == form_data.username).first()
    if not usuario or not verify_password(form_data.password, usuario.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Email o contraseña incorrectos.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    
    access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = create_access_token(
        data={"sub": usuario.email, "id": usuario.id},
        expires_delta=access_token_expires
    )
    return {"access_token": access_token, "token_type": "bearer"}


@app.get("/users/me", response_model=schemas.UsuarioOut, summary="Obtener información del usuario autenticado")
def obtener_perfil_usuario(current_user: models.Usuario = Depends(get_current_user)):
    return current_user


# ==========================================
# RUTAS DE DEMANDAS E HISTORIAL
# ==========================================


@app.post("/generar-demanda/", summary="Generar documento Word y registrar en la BD", operation_id="crear_nueva_demanda")
def generar_demanda(
    datos: schemas.DatosDemanda,
    request: Request,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == current_user.id).first()

    if not suscripcion:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="No posees una suscripción activa. Este servicio requiere un plan mensual pago para su uso."
        )

    if not suscripcion.activa:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tu suscripción se encuentra inactiva. Por favor, renová tu plan para continuar."
        )

    if suscripcion.fecha_expiracion and suscripcion.fecha_expiracion < datetime.utcnow():
        suscripcion.activa = False
        db.commit()
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tu suscripción mensual ha expirado. Por favor, actualizá tu pago para recuperar el acceso ilimitado."
        )

    plantilla = db.query(models.Plantilla).filter(
        models.Plantilla.id == datos.plantilla_id,
        models.Plantilla.activa == True
    ).first()

    if not plantilla or not os.path.exists(plantilla.ruta_archivo):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"La plantilla especificada (ID: {datos.plantilla_id}) no existe o el archivo base no está disponible."
        )

    carpeta_salida = "demandas_generadas"
    os.makedirs(carpeta_salida, exist_ok=True)

    nombre_limpio = datos.NombreActor.replace(' ', '_')
    ruta_salida = os.path.join(carpeta_salida, f"temp_{nombre_limpio}.docx")

    valor_punto = 2000000.0
    incapacidad_fisica = datos.PuntosdeIncapacidad * valor_punto
    dano_moral = incapacidad_fisica * 0.33
    dano_psicologico = incapacidad_fisica * 0.15
    gastos_farmacia = 1500000.0
    gastos_medicos = 2000000.0
    
    liquidacion_total = (
        datos.LiquiDanoMaterialNum + 
        incapacidad_fisica + 
        dano_moral + 
        dano_psicologico + 
        gastos_farmacia + 
        gastos_medicos
    )

    try:
        opcion_comp_int = int(datos.OpcionCompetencia)
    except ValueError:
        opcion_comp_int = 1

    texto_competencia = PARRAFOS_COMPETENCIA.get(
        opcion_comp_int, 
        PARRAFOS_COMPETENCIA[1]
    )

 # 5. MAPEO DE VARIABLES E INYECCIÓN
    
    if datos.ListaDocumental:
        lista_doc_limpia = [doc.strip() for doc in datos.ListaDocumental.split(",")]
    else:
        lista_doc_limpia = []

    datos_procesados = {
        "NombreActor": datos.NombreActor,
        "DniActor": f"{datos.DniActor:,}".replace(",", "."),
        "ParrafoCompetencia": texto_competencia,
        "PuntosdeIncapacidad": str(datos.PuntosdeIncapacidad),
        "IncapacidadFisicaPorcentaje": f"{datos.PuntosdeIncapacidad}%",
        
        "LiquiDanoMaterialNum": formatear_moneda(datos.LiquiDanoMaterialNum),
        "LiquiDanoMaterialLetras": monto_a_letras_legal(datos.LiquiDanoMaterialNum),
        "LiquiIncapacidadFisicaNum": formatear_moneda(incapacidad_fisica),
        "LiquiIncapacidadFisicaLetras": monto_a_letras_legal(incapacidad_fisica),
        "LiquiDanoMoralNum": formatear_moneda(dano_moral),
        "LiquiDanoMoralLetras": monto_a_letras_legal(dano_moral),
        "LiquiDanoPsicologicoNum": formatear_moneda(dano_psicologico),
        "LiquiDanoPsicologicoLetras": monto_a_letras_legal(dano_psicologico),
        "LiquiGastosFarmaciaNum": formatear_moneda(gastos_farmacia),
        "LiquiGastosFarmaciaLetras": monto_a_letras_legal(gastos_farmacia),
        "LiquiGastosMedicosNum": formatear_moneda(gastos_medicos),
        "LiquiGastosMedicosLetras": monto_a_letras_legal(gastos_medicos),
        "LiquiTotalNum": formatear_moneda(liquidacion_total),
        "LiquiTotalLetras": monto_a_letras_legal(liquidacion_total),
        
        "DomicilioActor": datos.DomicilioActor,
        "NombreDemandado": datos.NombreDemandado,
        "DomicilioDemandado": datos.DomicilioDemandado,
        "AutoDemandado": datos.AutoDemandado,
        "FechaHecho": datos.FechaHecho,
        "NombreAseguradora": datos.NombreAseguradora,
        "CuitAseguradora": datos.CuitAseguradora,
        "DomicilioAseguradora": datos.DomicilioAseguradora,
        "DescripcionHechos": datos.DescripcionHechos,
        "LesionesDetalles": datos.LesionesDetalles,
        "ListadoSecuelas": datos.ListadoSecuelas,
        "VehiculoActor": datos.VehiculoActor,
        "TallerNombre": datos.TallerNombre,
        "DirecciónTaller": datos.DirecciónTaller,
        
        "ListaDocumental": lista_doc_limpia,
        
        "CentroMedico": datos.CentroMedico,
        "CentroMedicoDireccion": datos.CentroMedicoDireccion,
        "LugarHecho": datos.LugarHecho,
        "FechaPresupuesto": datos.FechaPresupuesto
    }

    try:
        doc = DocxTemplate(plantilla.ruta_archivo)

        ruta_logo = "assets/logo_defecto.png"
        logo_imagen = InlineImage(doc, ruta_logo, width=Mm(40)) if os.path.exists(ruta_logo) else ""
        datos_procesados["logo_estudio"] = logo_imagen

        doc.render(datos_procesados)
        doc.save(ruta_salida)

        ip_cliente = request.client.host if request.client else "Desconocida"
        user_agent_cliente = request.headers.get("user-agent", "Desconocido")

        try:
            dni_val = int(datos.DniActor) if datos.DniActor else None
        except (ValueError, TypeError):
            dni_val = None

        nueva_demanda = models.DemandaGenerada(
            usuario_id=current_user.id,
            plantilla_id=plantilla.id,
            dni_actor=dni_val,
            nombre_actor=datos.NombreActor,
            estado_operativo="Generada",
            ip_origen=ip_cliente,
            user_agent=user_agent_cliente,
            archivo_generado=ruta_salida
        )
        db.add(nueva_demanda)

        if hasattr(suscripcion, 'demandas_restantes') and suscripcion.demandas_restantes is not None:
            suscripcion.demandas_restantes -= 1

        nuevo_log = models.AuditoriaLog(
            usuario_id=current_user.id,
            accion="GENERAR_DEMANDA",
            ip_origen=ip_cliente,
            detalles=f"Demanda para {datos.NombreActor} generada con plantilla ID {plantilla.id}."
        )
        db.add(nuevo_log)

        db.commit()
        db.refresh(nueva_demanda)

    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"Error interno al procesar la demanda: {str(e)}")

    return FileResponse(
        path=ruta_salida,
        filename=f"demanda_{nombre_limpio}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.get("/mis-demandas", summary="Listar todas las demandas generadas por el usuario actual")
def listar_mis_demandas(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # Consultamos las demandas filtradas por el ID del usuario logueado
    demandas = db.query(models.DemandaGenerada).filter(
        models.DemandaGenerada.usuario_id == current_user.id
    ).all()
    
    return {
        "cantidad": len(demandas),
        "demandas": [
            {
                "id": d.id,
                "fecha_creacion": d.fecha_creacion if hasattr(d, "fecha_creacion") else "N/A",
                "ruta_archivo": d.ruta_archivo,
                # Agregá los campos extra que tenga tu modelo DemandaGenerada
            } 
            for d in demandas
        ]
    }
        
    return resultado

@app.get("/historial", response_model=List[schemas.DemandaHistorialOut], summary="Obtener historial de demandas del usuario con filtros y paginación")
def obtener_historial(
    limit: int = 10,
    skip: int = 0,
    nombre_actor: Optional[str] = None,
    dni_actor: Optional[int] = None,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    """
    Devuelve las demandas del usuario autenticado de forma paginada y filtrable.
    
    - **limit**: Cantidad de registros por página (por defecto 10).
    - **skip**: Cantidad de registros a saltar/omitir (para avanzar de página).
    - **nombre_actor**: Búsqueda parcial por nombre del actor.
    - **dni_actor**: Búsqueda por DNI exacto.
    """
    query = db.query(models.DemandaGenerada).filter(models.DemandaGenerada.usuario_id == current_user.id)

    if nombre_actor:
        query = query.filter(models.DemandaGenerada.nombre_actor.ilike(f"%{nombre_actor}%"))

    if dni_actor:
        query = query.filter(models.DemandaGenerada.dni_actor == dni_actor)

    demandas = query.order_by(models.DemandaGenerada.fecha_creacion.desc()).offset(skip).limit(limit).all()

    return demandas

@app.post("/generar-demanda/", summary="Generar documento Word y registrar en la BD")
def generar_demanda(
    datos: schemas.DatosDemanda, 
    request: Request, 
    db: Session = Depends(get_db), 
    current_user: models.Usuario = Depends(get_current_user)
):
    # 1. VERIFICAR Y CONSULTAR LA SUSCRIPCIÓN DEL USUARIO
    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == current_user.id).first()

    # Si por alguna razón el usuario no tiene registro de suscripción, le creamos la "Free" por defecto
    if not suscripcion:
        suscripcion = models.Suscripcion(usuario_id=current_user.id, plan="Free", demandas_restantes=3)
        db.add(suscripcion)
        db.commit()
        db.refresh(suscripcion)

    # Validar que le queden demandas
    if suscripcion.demandas_restantes <= 0:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Has alcanzado el límite de demandas de tu plan. Actualizá tu suscripción para continuar."
        )

    # 2. BUSCAR LA PLANTILLA EN LA BASE DE DATOS
    plantilla = db.query(models.Plantilla).filter(
        models.Plantilla.id == datos.plantilla_id,
        models.Plantilla.activa == True
    ).first()

    if not plantilla or not os.path.exists(plantilla.ruta_archivo):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"La plantilla especificada (ID: {datos.plantilla_id}) no existe o el archivo base no está disponible."
        )

    # 3. PREPARACIÓN DE CARPETAS Y RUTAS
    carpeta_salida = "demandas_generadas"
    os.makedirs(carpeta_salida, exist_ok=True)

    nombre_limpio = datos.NombreActor.replace(' ', '_')
    ruta_salida = os.path.join(carpeta_salida, f"temp_{nombre_limpio}.docx")

    # 4. LÓGICA MATEMÁTICA Y DERECHO DUREZA
    valor_punto = 2000000.0
    incapacidad_fisica = datos.PuntosdeIncapacidad * valor_punto
    dano_moral = incapacidad_fisica * 0.33
    dano_psicologico = incapacidad_fisica * 0.15
    gastos_farmacia = 1500000.0
    gastos_medicos = 2000000.0
    
    liquidacion_total = (
        datos.LiquiDanoMaterialNum + 
        incapacidad_fisica + 
        dano_moral + 
        dano_psicologico + 
        gastos_farmacia + 
        gastos_medicos
    )

    texto_competencia = PARRAFOS_COMPETENCIA.get(
        datos.OpcionCompetencia, 
        PARRAFOS_COMPETENCIA[1]
    )

    # 5. MAPEO DE VARIABLES E INYECCIÓN
    # Convertimos el texto separado por comas en una lista de Python
    if datos.ListaDocumental:
        lista_doc_limpia = [doc.strip() for doc in datos.ListaDocumental.split(",")]
    else:
        lista_doc_limpia = []
    datos_procesados = {
        "NombreActor": datos.NombreActor,
        "DniActor": f"{datos.DniActor:,}".replace(",", "."),
        "ParrafoCompetencia": texto_competencia,
        "PuntosdeIncapacidad": str(datos.PuntosdeIncapacidad),
        "IncapacidadFisicaPorcentaje": f"{datos.PuntosdeIncapacidad}%",
        
        "LiquiDanoMaterialNum": formatear_moneda(datos.LiquiDanoMaterialNum),
        "LiquiDanoMaterialLetras": monto_a_letras_legal(datos.LiquiDanoMaterialNum),
        "LiquiIncapacidadFisicaNum": formatear_moneda(incapacidad_fisica),
        "LiquiIncapacidadFisicaLetras": monto_a_letras_legal(incapacidad_fisica),
        "LiquiDanoMoralNum": formatear_moneda(dano_moral),
        "LiquiDanoMoralLetras": monto_a_letras_legal(dano_moral),
        "LiquiDanoPsicologicoNum": formatear_moneda(dano_psicologico),
        "LiquiDanoPsicologicoLetras": monto_a_letras_legal(dano_psicologico),
        "LiquiGastosFarmaciaNum": formatear_moneda(gastos_farmacia),
        "LiquiGastosFarmaciaLetras": monto_a_letras_legal(gastos_farmacia),
        "LiquiGastosMedicosNum": formatear_moneda(gastos_medicos),
        "LiquiGastosMedicosLetras": monto_a_letras_legal(gastos_medicos),
        "LiquiTotalNum": formatear_moneda(liquidacion_total),
        "LiquiTotalLetras": monto_a_letras_legal(liquidacion_total),
        
        "DomicilioActor": datos.DomicilioActor,
        "NombreDemandado": datos.NombreDemandado,
        "DomicilioDemandado": datos.DomicilioDemandado,
        "AutoDemandado": datos.AutoDemandado,
        "FechaHecho": datos.FechaHecho,
        "NombreAseguradora": datos.NombreAseguradora,
        "CuitAseguradora": datos.CuitAseguradora,
        "DomicilioAseguradora": datos.DomicilioAseguradora,
        "DescripcionHechos": datos.DescripcionHechos,
        "LesionesDetalles": datos.LesionesDetalles,
        "ListadoSecuelas": datos.ListadoSecuelas,
        "VehiculoActor": datos.VehiculoActor,
        "TallerNombre": datos.TallerNombre,
        "DirecciónTaller": datos.DirecciónTaller,
        "ListaDocumental": lista_doc_limpia,
        
        "CentroMedico": datos.CentroMedico,
        "CentroMedicoDireccion": datos.CentroMedicoDireccion,
        "LugarHecho": datos.LugarHecho,
        "FechaPresupuesto": datos.FechaPresupuesto
    }
    try:
        # Usamos la ruta recuperada desde la Base de Datos
        doc = DocxTemplate(plantilla.ruta_archivo)

        ruta_logo = "assets/logo_defecto.png"
        logo_imagen = InlineImage(doc, ruta_logo, width=Mm(40)) if os.path.exists(ruta_logo) else ""
        datos_procesados["logo_estudio"] = logo_imagen

        doc.render(datos_procesados)
        doc.save(ruta_salida)

        # 6. PERSISTENCIA EN BD, DESCUENTO DE CRÉDITOS Y AUDITORÍA
        ip_cliente = request.client.host if request.client else "Desconocida"
        user_agent_cliente = request.headers.get("user-agent", "Desconocido")

        try:
            dni_val = int(datos.DniActor) if hasattr(datos, "DniActor") and datos.DniActor else None
        except (ValueError, TypeError):
            dni_val = None

        # a) Guardar la demanda vinculada
        nueva_demanda = models.DemandaGenerada(
            usuario_id=current_user.id,
            plantilla_id=plantilla.id, # 👈 Guardamos el ID de la plantilla usada
            dni_actor=dni_val,
            nombre_actor=datos.NombreActor,
            estado_operativo="Generada",
            ip_origen=ip_cliente,
            user_agent=user_agent_cliente
        )
        db.add(nueva_demanda)

        # b) Descontar 1 demanda de la suscripción
        suscripcion.demandas_restantes -= 1

        # c) Registrar log de auditoría
        nuevo_log = models.AuditoriaLog(
            usuario_id=current_user.id,
            accion="GENERAR_DEMANDA",
            ip_origen=ip_cliente,
            detalles=f"Demanda para {datos.NombreActor} generada con plantilla ID {plantilla.id}. Restantes: {suscripcion.demandas_restantes}"
        )
        db.add(nuevo_log)

        # Guardar todas las operaciones juntas (Transacción Atómica)
        db.commit()
        db.refresh(nueva_demanda)

        print(f"🔒 [SISTEMA] Demanda #{nueva_demanda.id} generada. Créditos restantes de Usuario #{current_user.id}: {suscripcion.demandas_restantes}")

    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"Error interno al procesar la demanda: {str(e)}")

    return FileResponse(
        path=ruta_salida,
        filename=f"demanda_{nombre_limpio}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# 🔄 WORKFLOW: Endpoint para actualizar el estado operativo de una demanda
@app.patch("/demanda/{demanda_id}/estado", summary="Actualizar estado operativo de una demanda")
def actualizar_estado_demanda(demanda_id: int, nuevo_estado: str, db: Session = Depends(get_db)):
    """
    Estados permitidos recomendados: 'Generada', 'Presentada', 'En Notificacion', 'Archivada'
    """
    registro = db.query(models.DemandaGenerada).filter(models.DemandaGenerada.id == demanda_id).first()
    
    if not registro:
        raise HTTPException(status_code=404, detail="No se encontró la demanda especificada.")
        
    registro.estado_operativo = nuevo_estado
    db.commit()
    db.refresh(registro)
    
    return {
        "mensaje": f"Estado de la demanda #{demanda_id} actualizado a '{nuevo_estado}' con éxito.",
        "demanda": registro
    }


from fastapi.responses import FileResponse

@app.get("/descargar-demanda/{demanda_id}", summary="Descargar documento Word generado", operation_id="descargar_demanda_por_id")
def descargar_demanda(
    demanda_id: int,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # 1. Buscar el registro de la demanda en la base de datos
    demanda = db.query(models.DemandaGenerada).filter(models.DemandaGenerada.id == demanda_id).first()
    
    if not demanda:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="La demanda especificada no existe en la base de datos."
        )
    
    # 2. Control de seguridad: Verificar que la demanda pertenezca al usuario autenticado
    if demanda.usuario_id != current_user.id:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="No tenés autorización para descargar este documento."
        )
    
    # 3. Verificar que el archivo físico exista en el servidor
    if not os.path.exists(demanda.ruta_archivo):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="El archivo físico ya no se encuentra disponible en el servidor."
        )
    
    # 4. Devolver el archivo como respuesta descargable
    nombre_archivo = os.path.basename(demanda.ruta_archivo)
    return FileResponse(
        path=demanda.ruta_archivo,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=nombre_archivo
    )
@app.get("/plantillas", response_model=List[schemas.PlantillaOut], summary="Listar plantillas de demandas disponibles")
def obtener_plantillas(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    """
    Devuelve la lista de todas las plantillas activas disponibles en el sistema 
    para que el usuario elija cuál utilizar al generar su demanda.
    """
    plantillas = db.query(models.Plantilla).filter(models.Plantilla.activa == True).all()
    return plantillas
    
    # PREVISUALIZACIÓN DE AUDITORÍA
@app.post("/preview-demanda/")
def preview_demanda(datos: schemas.DatosDemanda):
    # 1. Cálculos matemáticos idénticos al generador final
    valor_punto = 2000000.0
    incapacidad_fisica = datos.PuntosdeIncapacidad * valor_punto
    dano_moral = incapacidad_fisica * 0.33
    dano_psicologico = incapacidad_fisica * 0.15
    gastos_farmacia = 1500000.0
    gastos_medicos = 2000000.0
    
    liquidacion_total = (
        datos.LiquiDanoMaterialNum + 
        incapacidad_fisica + 
        dano_moral + 
        dano_psicologico + 
        gastos_farmacia + 
        gastos_medicos
    )

    # 2. Selección del texto de competencia
    texto_competencia = PARRAFOS_COMPETENCIA.get(
        datos.OpcionCompetencia, 
        PARRAFOS_COMPETENCIA[1]
    )

    # 3. Retorno del 100% de los datos mapeados (Los 24 campos expuestos)
    return {
        "estado": "Éxito",
        "mensaje": "Auditoría generada. Verifique todos los campos ingresados.",
        "datos_para_revision": {
            "1_DATOS_ACTOR": {
                "NombreActor": datos.NombreActor,
                "DniActor": datos.DniActor,
                "DomicilioActor": datos.DomicilioActor,
                "VehiculoActor": datos.VehiculoActor
            },
            "2_DATOS_DEMANDADO_Y_SEGURO": {
                "NombreDemandado": datos.NombreDemandado,
                "DomicilioDemandado": datos.DomicilioDemandado,
                "AutoDemandado": datos.AutoDemandado,
                "NombreAseguradora": datos.NombreAseguradora,
                "CuitAseguradora": datos.CuitAseguradora,
                "DomicilioAseguradora": datos.DomicilioAseguradora
            },
            "3_HECHOS_Y_LESIONES": {
                "FechaHecho": datos.FechaHecho,
                "LugarHecho": datos.LugarHecho,
                "DescripcionHechos": datos.DescripcionHechos,
                "LesionesDetalles": datos.ListadoSecuelas,
                "ListadoSecuelas": datos.ListadoSecuelas
            },
            "4_PRUEBA_Y_ATENCION": {
                "CentroMedico": datos.CentroMedico,
                "CentroMedicoDireccion": datos.CentroMedicoDireccion,
                "TallerNombre": datos.TallerNombre,
                "DirecciónTaller": datos.DirecciónTaller,
                "FechaPresupuesto": datos.FechaPresupuesto,
                "ListaDocumental": datos.ListaDocumental
            },
            "5_COMPETENCIA_Y_LIQUIDACION": {
                "OpcionCompetencia_Elegida": datos.OpcionCompetencia,
                "Texto_Competencia_Asignado": texto_competencia,
                "PuntosdeIncapacidad_Ingresado": datos.PuntosdeIncapacidad,
                "Daño_Material_Ingresado": formatear_moneda(datos.LiquiDanoMaterialNum),
                "Incapacidad_Fisica_Calculada": formatear_moneda(incapacidad_fisica),
                "Daño_Moral_Calculado": formatear_moneda(dano_moral),
                "Daño_Psicologico_Calculado": formatear_moneda(dano_psicologico),
                "Gastos_Farmacia_Fijos": formatear_moneda(gastos_farmacia),
                "Gastos_Medicos_Fijos": formatear_moneda(gastos_medicos),
                "LIQUIDACION_TOTAL_NUM": formatear_moneda(liquidacion_total),
                "LIQUIDACION_TOTAL_LETRAS": monto_a_letras_legal(liquidacion_total)
            }
        }
    }

@app.post("/simular-pago/", summary="Simular pago exitoso y renovar suscripción por 30 días")
def simular_pago(
    plan: str = "Pro",
    demandas: int = 50,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    """
    Simula una pasarela de pago exitosa (como Mercado Pago o Stripe). 
    Actualiza la suscripción del usuario actual, activándola, otorgando 
    nuevos créditos y extendiendo la fecha de expiración 30 días a partir de hoy.
    """
    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == current_user.id).first()

    if not suscripcion:
        suscripcion = models.Suscripcion(usuario_id=current_user.id)
        db.add(suscripcion)

    suscripcion.plan = plan
    suscripcion.demandas_restantes = demandas
    suscripcion.activa = True
    suscripcion.fecha_inicio = datetime.utcnow()
    suscripcion.fecha_expiracion = datetime.utcnow() + timedelta(days=30)
    
    db.commit()
    db.refresh(suscripcion)

    return {
        "mensaje": "¡Pago simulado con éxito!",
        "usuario": current_user.email,
        "plan": suscripcion.plan,
        "demandas_restantes": suscripcion.demandas_restantes,
        "fecha_inicio": suscripcion.fecha_inicio,
        "fecha_expiracion": suscripcion.fecha_expiracion
    }
@app.get("/descargar-demanda/{demanda_id}", summary="Descargar un documento generado específico por su ID")
def descargar_demanda_por_id(
    demanda_id: int,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # Buscar la demanda asegurando que pertenezca al usuario autenticado
    demanda = db.query(models.DemandaGenerada).filter(
        models.DemandaGenerada.id == demanda_id,
        models.DemandaGenerada.usuario_id == current_user.id
    ).first()

    if not demanda:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="La demanda especificada no existe o no tenés permisos para acceder a ella."
        )

    if not demanda.archivo_generado or not os.path.exists(demanda.archivo_generado):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="El archivo físico asociado a esta demanda ya no se encuentra disponible en el servidor."
        )

    nombre_archivo = os.path.basename(demanda.archivo_generado)

    return FileResponse(
        path=demanda.archivo_generado,
        filename=nombre_archivo,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
@app.post("/crear-preferencia-suscripcion/", summary="Crear preferencia de pago en Mercado Pago")
def crear_preferencia_suscripcion(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    sdk = mercadopago.SDK(os.getenv("MP_ACCESS_TOKEN"))

    preference_data = {
        "items": [
            {
                "title": "Suscripción Mensual - SaaS Demandas Legales",
                "quantity": 1,
                "currency_id": "ARS",
                "unit_price": 15000.0  # Precio mensual de ejemplo
            }
        ],
        "payer": {
            "email": current_user.email
        },
        "back_urls": {
            "success": "https://snide-uranium-hungrily.ngrok-free.dev/pago-exitoso",
            "failure": "https://snide-uranium-hungrily.ngrok-free.dev/pago-fallido",
            "pending": "https://snide-uranium-hungrily.ngrok-free.dev/pago-pendiente"
        },
        "auto_return": "approved",
        "external_reference": str(current_user.id)  # Guardamos el ID del usuario para identificarlo después
    }

    try:
        preference_response = sdk.preference().create(preference_data)
        print("RESPUESTA DE MERCADO PAGO:", preference_response)

        if preference_response.get("status") not in [200, 201]:
            raise HTTPException(
                status_code=400, 
                detail=f"Error de Mercado Pago: {preference_response}"
            )

        preference = preference_response["response"]

        return {
            "preference_id": preference["id"],
            "init_point": preference["init_point"]  # Link al que redirigís al usuario para pagar
        }

    except Exception as e:
        print("EXCEPCIÓN CAPTURADA:", str(e))
        raise HTTPException(
            status_code=500,
            detail=f"Error al conectar con la pasarela de pagos: {str(e)}"
        )
    

@app.post("/webhook/mercadopago")
async def mercadopago_webhook(request: Request, db: Session = Depends(get_db)):
    try:
        body = await request.json()
        print("-> Webhook recibido de Mercado Pago:", body)

        # Identificar el tipo de notificación (soporta formato moderno y clásico)
        topic = body.get("type") or body.get("topic")
        action = body.get("action")
        
        payment_id = None

        # Formato moderno (action: payment.created / payment.updated)
        if action and "payment" in action:
            data = body.get("data", {})
            payment_id = data.get("id")
        
        # Formato clásico (type: payment o topic: payment)
        elif topic == "payment":
            payment_id = body.get("id") or body.get("data", {}).get("id")
        
        # Formato IPN clásico con resource URL
        elif "resource" in body:
            resource_url = body.get("resource")
            if "payments" in resource_url:
                payment_id = resource_url.split("/")[-1]

        if not payment_id:
            return {"status": "ignored", "message": "No se encontró el ID de pago"}

        # Token de acceso de Mercado Pago
        mp_access_token = os.getenv("MP_ACCESS_TOKEN", "TU_ACCESS_TOKEN_DE_MERCADO_PAGO")

        # Consultar los detalles reales del pago a la API de Mercado Pago
        async with httpx.AsyncClient() as client:
            headers = {"Authorization": f"Bearer {mp_access_token}"}
            response = await client.get(f"https://api.mercadopago.com/v1/payments/{payment_id}", headers=headers)
            
            if response.status_code != 200:
                return {"status": "error", "message": "No se pudo consultar el pago en Mercado Pago"}
            
            payment_data = response.json()

        status = payment_data.get("status") # ej: "approved", "rejected", "pending"
        status_detail = payment_data.get("status_detail") # ej: "cc_rejected_insufficient_amount", "bad_security_code"
        external_reference = payment_data.get("external_reference") # ID del usuario enviado en la preferencia

        print(f"💰 Pago {payment_id} | Estado: {status} | Detalle: {status_detail} | Ref Usuario: {external_reference}")

        if not external_reference:
            return {"status": "ignored", "message": "El pago no tiene un external_reference asociado"}

        # Buscamos al usuario en la base de datos usando el external_reference
        user = db.query(User).filter(User.id == int(external_reference)).first()

        if not user:
            print(f"⚠️ Usuario con ID {external_reference} no encontrado en la base de datos.")
            return {"status": "error", "message": "Usuario no encontrado"}

        # 1. Pago Aprobado -> Activamos la suscripción
        if status == "approved":
            user.suscripcion_activa = True
            user.estado_pago = "approved"
            db.commit()
            print(f"✅ [DB] Suscripción activada con éxito para el usuario ID: {user.id}")

        # 2. Pago Pendiente o En Proceso -> Marcamos como pendiente
        elif status in ["pending", "in_process"]:
            user.suscripcion_activa = False
            user.estado_pago = status
            db.commit()
            print(f"⏳ [DB] Pago pendiente registrado para el usuario ID: {user.id}")

        # 3. Pago Rechazado -> Mantenemos inactivo y guardamos el estado
        elif status == "rejected":
            user.suscripcion_activa = False
            user.estado_pago = status
            db.commit()
            print(f"❌ [DB] Pago rechazado ({status_detail}) registrado para el usuario ID: {user.id}")

        return {"status": "success"}

    except Exception as e:
        db.rollback() # Revertir cambios si algo falla
        print(f"❌ Error crítico en webhook: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
@app.get("/admin/estadisticas", summary="Estadísticas globales para el panel de administración")
def obtener_estadisticas_admin(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # Verificamos si el usuario actual es administrador
    if not getattr(current_user, "es_admin", False):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Acceso denegado. Se requieren permisos de administrador."
        )

    total_usuarios = db.query(models.Usuario).count()
    total_demandas = db.query(models.DemandaGenerada).count()
    suscripciones_activas = db.query(models.Suscripcion).filter(models.Suscripcion.activa == True).count()

    return {
        "total_usuarios": total_usuarios,
        "total_demandas_generadas": total_demandas,
        "suscripciones_activas": suscripciones_activas
    }

@app.get("/descargar-demanda/{demanda_id}", summary="Descargar documento Word generado", operation_id="descargar_demanda_por_id")
def descargar_demanda(
    demanda_id: int,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # 1. Buscar la demanda en la base de datos
    demanda = db.query(models.DemandaGenerada).filter(models.DemandaGenerada.id == demanda_id).first()
    
    if not demanda:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="La demanda especificada no existe en la base de datos."
        )
    
    # 2. Control de seguridad: Verificar que la demanda pertenezca al usuario autenticado
    if demanda.usuario_id != current_user.id:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="No tenés autorización para descargar este documento."
        )
    
    # 3. Verificar que el archivo físico exista en el servidor
    if not os.path.exists(demanda.ruta_archivo):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="El archivo físico ya no se encuentra disponible en el servidor."
        )
    
    # 4. Devolver el archivo como respuesta descargable
    nombre_archivo = os.path.basename(demanda.ruta_archivo)
    return FileResponse(
        path=demanda.ruta_archivo,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=nombre_archivo
    )
@app.get("/suscripcion/estado", summary="Verificar el estado de la suscripción actual")
def verificar_estado_suscripcion(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # Buscamos la suscripción asociada al usuario autenticado[cite: 1]
    suscripcion = db.query(models.Suscripcion).filter(
        models.Suscripcion.usuario_id == current_user.id
    ).first()

    if not suscripcion:
        return {
            "tiene_suscripcion": False,
            "activa": False,
            "mensaje": "No posees ningún plan o suscripción registrada."
        }

    # Verificamos si la suscripción expiró por fecha (actualizando el estado automáticamente)
    if suscripcion.fecha_expiracion and suscripcion.fecha_expiracion < datetime.utcnow():
        if suscripcion.activa:
            suscripcion.activa = False
            db.commit()

    return {
        "tiene_suscripcion": True,
        "activa": suscripcion.activa,
        "fecha_expiracion": suscripcion.fecha_expiracion,
        "usuario_email": current_user.email
    }

@app.get("/admin/panel-control", summary="Panel exclusivo para administradores")
def panel_control_admin(
    db: Session = Depends(get_db),
    admin_user: models.Usuario = Depends(get_current_admin_user)
):
    return {
        "status": "success",
        "mensaje": f"Bienvenido al panel de administración, {admin_user.email}. Acceso concedido."
    }

@app.post("/admin/plantillas", response_model=schemas.PlantillaResponse, status_code=status.HTTP_201_CREATED, summary="Registrar nueva plantilla")
def registrar_plantilla(
    plantilla: schemas.PlantillaCreate,
    db: Session = Depends(get_db),
    admin_user: models.Usuario = Depends(get_current_admin_user)
):
    """
    Registra una nueva plantilla legal en el sistema. 
    Exclusivo para administradores.
    """
    nueva_plantilla = models.Plantilla(
        nombre=plantilla.nombre,
        categoria=plantilla.categoria,
        descripcion=plantilla.descripcion,
        ruta_archivo=plantilla.ruta_archivo,
        activa=plantilla.activa
    )
    
    db.add(nueva_plantilla)
    db.commit()
    db.refresh(nueva_plantilla)
    
    return nueva_plantilla

@app.patch("/admin/plantillas/{plantilla_id}/estado", summary="Habilitar o deshabilitar una plantilla existente")
def cambiar_estado_plantilla(
    plantilla_id: int,
    estado_data: schemas.PlantillaEstadoUpdate,
    db: Session = Depends(get_db),
    admin_user: models.Usuario = Depends(get_current_admin_user)
):
    """
    Permite activar o desactivar una plantilla mediante un JSON en el body.
    """
    plantilla = db.query(models.Plantilla).filter(models.Plantilla.id == plantilla_id).first()
    if not plantilla:
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    
    plantilla.activa = estado_data.activa
    db.commit()
    
    estado_texto = "habilitada" if estado_data.activa else "deshabilitada"
    return {"status": "success", "mensaje": f"La plantilla ha sido {estado_texto} correctamente."}