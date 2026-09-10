import os
import json
import uuid
from datetime import datetime, timedelta, timezone
from typing import Optional, List, Any
from pydantic import BaseModel
from pydantic import EmailStr

# 1. CARGAR LAS VARIABLES DE ENTORNO ANTES DE CUALQUIER OTRA COSA
from dotenv import load_dotenv
load_dotenv(override=True)  # El override=True obliga a leer siempre del .env

# Librerías de terceros
import bcrypt
import httpx
import mercadopago
import resend  # <-- NUEVA INTEGRACIÓN DE CORREOS
from docx import Document
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage, RichText
from num2words import num2words
from jose import JWTError, jwt
from itsdangerous import SignatureExpired, BadSignature
import google.generativeai as genai

# FastAPI y utilidades de Web/API
from fastapi import (
    FastAPI,
    Request,
    HTTPException,
    Depends,
    status,
    Form,
    File,
    UploadFile,
    Query,
    BackgroundTasks,
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, RedirectResponse, HTMLResponse
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from sqlalchemy.orm import Session

# Módulos propios del proyecto
from database import get_db, engine
import models
import schemas
from models import Usuario as User
from email_utils import enviar_correo
from storage_utils import upload_to_gcp, generate_signed_url
import security
from security import (
    get_current_user,
    get_current_admin_user,
    verificar_suscripcion_activa,
    serializer,
    pwd_context,
)

# Configuración de integraciones externas
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
resend.api_key = os.getenv("RESEND_API_KEY")  # <-- CLAVE API DE RESEND
sdk = mercadopago.SDK(os.getenv("MP_ACCESS_TOKEN"))

# Inicialización de la aplicación FastAPI
app = FastAPI(title="SaaS Demandas Legal API", version="0.3.0")

# Creación automática de tablas en DB
models.Base.metadata.create_all(bind=engine)

# Configuración Dinámica de CORS para Producción y Desarrollo
FRONTEND_URL = os.getenv("FRONTEND_URL", "http://127.0.0.1:5500/frontend_demandas")

origins = [
    "http://127.0.0.1:5500",
    "http://localhost:5500",
    "https://saas-demandas-legal.onrender.com",
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "https://saas-demandas-legal.vercel.app",
    "https://autodemandas.com.ar",            # Tu nuevo dominio raíz
    "https://www.autodemandas.com.ar",
    FRONTEND_URL,
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 2. Configuración JWT
SECRET_KEY = "4ut0D3m4nd452027#"
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 10  # La sesión expira 10 minutos después del login

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


def create_access_token(data: dict, expires_delta: Optional[timedelta] = None):
    to_encode = data.copy()
    if expires_delta:
        expire = datetime.utcnow() + expires_delta
    else:
        expire = datetime.utcnow() + timedelta(minutes=15)
    to_encode.update({"exp": expire})
    
    # Forzamos a usar la SECRET_KEY que vive dentro de security.py
    print(f"DEBUG - CLAVE USADA PARA FIRMAR EL TOKEN: {security.SECRET_KEY}")
    
    encoded_jwt = jwt.encode(to_encode, security.SECRET_KEY, algorithm=security.ALGORITHM)
    return encoded_jwt


@app.post("/token", response_model=schemas.Token, summary="Iniciar sesión y obtener JWT")
def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    usuario = db.query(models.Usuario).filter(models.Usuario.email == form_data.username).first()
    
    # 1. Verificamos que el usuario exista y la contraseña sea correcta
    if not usuario or not verify_password(form_data.password, usuario.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Email o contraseña incorrectos.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    
    # 2. BLOQUEO DE SEGURIDAD: Verificamos si la cuenta fue deshabilitada por un admin
    if not usuario.activo:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tu cuenta ha sido inhabilitada. Por favor, contactá a soporte."
        )
    
    # 3. Si todo está bien, emitimos el token de acceso
    access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = create_access_token(
        data={"sub": usuario.email, "id": usuario.id},
        expires_delta=access_token_expires
    )
    return {"access_token": access_token, "token_type": "bearer"}

@app.get("/usuarios/me", response_model=schemas.PerfilOut, summary="Obtener información ampliada del usuario autenticado")
def obtener_perfil_usuario(current_user: models.Usuario = Depends(get_current_user), db: Session = Depends(get_db)):
    # Buscamos la suscripción vinculada al usuario
    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == current_user.id).first()
    creditos = suscripcion.demandas_restantes if suscripcion else 0
    plan = suscripcion.plan if suscripcion else "Sin Plan"
    
    return {
        "id": current_user.id,
        "email": current_user.email,
        "nombre_estudio": current_user.nombre_estudio,
        "es_admin": current_user.es_admin,
        "activo": current_user.activo,
        "rol_estudio": current_user.rol_estudio,
        "cuenta_madre_id": current_user.cuenta_madre_id,
        "creditos_disponibles": creditos,
        "plan_actual": plan
    }
    
@app.put("/usuarios/password", summary="Actualizar contraseña desde el perfil")
def cambiar_password(
    datos: schemas.PasswordUpdate,
    current_user: models.Usuario = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    # 1. Verificar que el usuario conoce su contraseña actual
    if not verify_password(datos.password_actual, current_user.hashed_password):
        raise HTTPException(
            status_code=400, 
            detail="La contraseña actual ingresada es incorrecta."
        )
    
    # 2. Encriptar la nueva contraseña
    nuevo_hash = get_password_hash(datos.password_nueva)
    
    # 3. Guardar en la base de datos
    current_user.hashed_password = nuevo_hash
    db.commit()
    
    return {"mensaje": "Contraseña actualizada exitosamente."}

@app.post("/admin/usuarios/{usuario_id}/creditos", summary="Asignar o descontar créditos manualmente")
def actualizar_creditos(
    usuario_id: int, 
    datos: schemas.CreditosUpdate, 
    background_tasks: BackgroundTasks, 
    current_user: models.Usuario = Depends(get_current_user), 
    db: Session = Depends(get_db)
):
    # 1. Bloqueo de seguridad: Solo administradores
    if not current_user.es_admin:
        raise HTTPException(status_code=403, detail="Acceso denegado: Permisos de administrador requeridos")
    
    # 2. Buscar al usuario
    usuario_destino = db.query(models.Usuario).filter(models.Usuario.id == usuario_id).first()
    if not usuario_destino:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
        
    # 3. Buscar la billetera (Suscripcion)
    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == usuario_id).first()
    
    # --- NUEVA LÓGICA: Si no existe, la inicializamos automáticamente ---
    if not suscripcion:
        suscripcion = models.Suscripcion(
            usuario_id=usuario_id,
            plan="Free",           # Inicializamos con el plan base
            activa=True,
            demandas_restantes=0   # Arranca en 0, luego le suma el monto
        )
        db.add(suscripcion)
        db.flush() # Sincronizamos con la base de datos antes del commit final
    # ---------------------------------------------------------------------
    
    # 4. Aplicar el ajuste de saldo (permite números negativos para descontar)
    suscripcion.demandas_restantes += datos.monto
    
    # 5. Grabar el comprobante en la auditoría
    nuevo_movimiento = models.HistorialCreditos(
        usuario_id=usuario_id,
        monto=datos.monto,
        motivo=datos.motivo
    )
    db.add(nuevo_movimiento)
    db.commit()
    
    # 6. Notificación por correo vía Resend
    def notificar_recarga():
        try:
            resend.Emails.send({
                "from": "SaaS Legal <soporte@autodemandas.com.ar>",
                "to": [usuario_destino.email],
                "subject": "Actualización de saldo - SaaS Legal",
                "html": f"""
                <h3>¡Tu saldo ha sido actualizado!</h3>
                <p>Se ha registrado un movimiento de <strong>{datos.monto}</strong> créditos en tu cuenta.</p>
                <p><strong>Motivo:</strong> {datos.motivo}</p>
                <p><strong>Saldo actual disponible:</strong> {suscripcion.demandas_restantes}</p>
                """
            })
            print(f"✅ Notificación de saldo enviada a {usuario_destino.email}")
        except Exception as e:
            print(f"⚠️ Error enviando notificación: {e}")

    background_tasks.add_task(notificar_recarga)

    return {
        "mensaje": "Saldo actualizado exitosamente y correo enviado", 
        "nuevo_saldo": suscripcion.demandas_restantes
    }

# ==========================================
# RUTAS DE DEMANDAS E HISTORIAL
# ==========================================

@app.post("/generar-demanda/", summary="Generar documento Word y registrar en la BD")
def generar_demanda(
    datos: schemas.DatosDemanda, 
    request: Request, 
    background_tasks: BackgroundTasks, 
    db: Session = Depends(get_db), 
    current_user: models.Usuario = Depends(get_current_user) # Cambiado a get_current_user para no bloquear asistentes prematuramente
):
    # 1. DETERMINAR LA CUENTA TITULAR (MADRE O PROPIA)
    cuenta_titular = current_user
    es_asistente = False
    
    # Si el current_user tiene configurada una cuenta madre, redirigimos las validaciones a ese ID
    if getattr(current_user, 'cuenta_madre_id', None):
        cuenta_titular = db.query(models.Usuario).filter(models.Usuario.id == current_user.cuenta_madre_id).first()
        es_asistente = True
        
        if not cuenta_titular:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Error de jerarquía: La cuenta principal (estudio jurídico) asociada no existe."
            )

    # 2. VERIFICAR LA SUSCRIPCIÓN DE LA CUENTA TITULAR (En vez del current_user directo)
    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == cuenta_titular.id).first()

    if not suscripcion:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="La cuenta principal no posee una suscripción activa. Este servicio requiere un plan mensual pago."
        )

    if not suscripcion.activa:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="La suscripción de la cuenta principal se encuentra inactiva."
        )

    if suscripcion.fecha_expiracion and suscripcion.fecha_expiracion < datetime.utcnow():
        suscripcion.activa = False
        db.commit()
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="La suscripción de la cuenta principal ha expirado."
        )
        
    # Verificar si le quedan demandas a la cuenta titular
    if hasattr(suscripcion, 'demandas_restantes') and suscripcion.demandas_restantes is not None:
        if suscripcion.demandas_restantes < 1:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="La cuenta principal ya no tiene demandas disponibles en su plan actual."
            )

    # 3. SELECCIÓN DINÁMICA DE LA PLANTILLA SEGÚN EL FORMULARIO
    diccionario_plantillas = {
        "auto_moto": 1,
        "auto_auto": 2
    }
    plantilla_seleccionada = diccionario_plantillas.get(datos.TipoDemanda, 1)

    plantilla = db.query(models.Plantilla).filter(
        models.Plantilla.id == plantilla_seleccionada,
        models.Plantilla.activa == True
    ).first()

    if not plantilla or not os.path.exists(plantilla.ruta_archivo):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"La plantilla especificada (ID: {plantilla_seleccionada}) no existe o el archivo base no está disponible."
        )

    carpeta_salida = "demandas_generadas"
    os.makedirs(carpeta_salida, exist_ok=True)

    nombre_limpio = datos.NombreActor.replace(' ', '_')
    ruta_salida = os.path.join(carpeta_salida, f"temp_{nombre_limpio}.docx")

    # 4. CÁLCULOS MATEMÁTICOS (Sin alterar tu lógica)
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
    
    # Formatear Fechas
    fecha_medica_formateada = datos.FechaMedica
    if datos.FechaMedica and "-" in datos.FechaMedica:
        anio, mes, dia = datos.FechaMedica.split("-")
        fecha_medica_formateada = f"{dia}/{mes}/{anio}"
        
    fecha_presupuesto_formateada = datos.FechaPresupuesto
    if datos.FechaPresupuesto and "-" in datos.FechaPresupuesto:
        anio_p, mes_p, dia_p = datos.FechaPresupuesto.split("-")
        fecha_presupuesto_formateada = f"{dia_p}/{mes_p}/{anio_p}"
        
    fecha_hecho_formateada = datos.FechaHecho
    if datos.FechaHecho and "-" in datos.FechaHecho:
        anio_h, mes_h, dia_h = datos.FechaHecho.split("-")
        fecha_hecho_formateada = f"{dia_h}/{mes_h}/{anio_h}"

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
        "DniDemandado": datos.DniDemandado,
        "DomicilioDemandado": datos.DomicilioDemandado,
        "AutoDemandado": datos.AutoDemandado,
        "FechaHecho": fecha_hecho_formateada,
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
        "FechaPresupuesto": fecha_presupuesto_formateada,
        "FechaMedica": fecha_medica_formateada,
        "PorcentajeDanoPsicologico": datos.PorcentajeDanoPsicologico
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

        # --- INICIO MAGIA GOOGLE CLOUD ---
        nombre_unico = f"demandas/{current_user.id}_{uuid.uuid4().hex[:6]}_{nombre_limpio}.docx"
        upload_to_gcp(ruta_salida, nombre_unico)
        # --- FIN MAGIA GOOGLE CLOUD ---

        # Registramos la demanda a nombre del usuario actual (el asistente) para que pueda verla en "Mis Demandas"
        nueva_demanda = models.DemandaGenerada(
            usuario_id=current_user.id,
            plantilla_id=plantilla.id,
            dni_actor=dni_val,
            nombre_actor=datos.NombreActor,
            estado_operativo="Generada",
            ip_origen=ip_cliente,
            user_agent=user_agent_cliente,
            archivo_generado=nombre_unico
        )
        db.add(nueva_demanda)

        # 6. IMPACTAR LOS CRÉDITOS Y LA AUDITORÍA A NOMBRE DE LA CUENTA TITULAR
        if hasattr(suscripcion, 'demandas_restantes') and suscripcion.demandas_restantes is not None:
            suscripcion.demandas_restantes -= 1

        detalles_auditoria = f"Demanda para {datos.NombreActor} generada y respaldada en GCP."
        if es_asistente:
            detalles_auditoria += f" (Ejecutado por asistente: {current_user.email})"

        nuevo_log = models.AuditoriaLog(
            usuario_id=cuenta_titular.id, # Asignamos el log a la cuenta principal para su control
            accion="GENERAR_DEMANDA",
            ip_origen=ip_cliente,
            detalles=detalles_auditoria
        )
        db.add(nuevo_log)

        db.commit()
        db.refresh(nueva_demanda)

        print(f"🔒 [SISTEMA] Demanda #{nueva_demanda.id} guardada en GCP como: {nombre_unico}")

        # ENVÍO DE CORREO
        try:
            enviar_correo(
                destinatario=current_user.email,
                asunto="Tu demanda legal ha sido generada",
                contenido_html=f"<h2>¡Éxito {datos.NombreActor}!</h2><p>Adjunto documento respaldado en la nube.</p>",
                ruta_adjunto=ruta_salida
            )
            print("📧 Correo ejecutado síncronamente con éxito.")
        except Exception as mail_err:
            print(f"❌ Error al intentar disparar el correo: {mail_err}")

    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"Error interno al procesar la demanda: {str(e)}")

    return FileResponse(
        path=ruta_salida,
        filename=f"demanda_{nombre_limpio}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

from pydantic import BaseModel

# Definimos la estructura para recibir el dato de forma segura
class EstadoUpdate(BaseModel):
    nuevo_estado: str

@app.patch("/demanda/{demanda_id}/estado", summary="Actualizar estado operativo de una demanda")
def actualizar_estado_demanda(
    demanda_id: int, 
    datos: EstadoUpdate, 
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user) # <- Bloqueo de seguridad agregado
):
    """
    Estados permitidos recomendados: 'Generada', 'En Revisión', 'Lista para Presentar', 'Presentada', 'En Mediación', 'Archivada'
    """
    registro = db.query(models.DemandaGenerada).filter(models.DemandaGenerada.id == demanda_id).first()
    
    if not registro:
        raise HTTPException(status_code=404, detail="No se encontró la demanda especificada.")

    # 1. Validación B2B: Verificar que la demanda pertenezca al mismo paraguas (estudio jurídico)
    cuenta_madre_actual = getattr(current_user, 'cuenta_madre_id', None) or current_user.id
    
    creador = db.query(models.Usuario).filter(models.Usuario.id == registro.usuario_id).first()
    cuenta_madre_creador = getattr(creador, 'cuenta_madre_id', None) or creador.id

    if cuenta_madre_actual != cuenta_madre_creador:
        raise HTTPException(status_code=403, detail="No tienes permisos para editar el estado de esta demanda.")

    # 2. Aplicar el cambio
    registro.estado_operativo = datos.nuevo_estado
    db.commit()
    db.refresh(registro)
    
    return {
        "mensaje": f"Estado de la demanda #{demanda_id} actualizado a '{datos.nuevo_estado}' con éxito.",
        "demanda": registro
    }


@app.get("/descargar-demanda/{demanda_id}", summary="Descargar documento Word seguro desde la nube")
def descargar_demanda_nube(
    demanda_id: int,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(verificar_suscripcion_activa)
):
    demanda = db.query(models.DemandaGenerada).filter(models.DemandaGenerada.id == demanda_id).first()
    
    if not demanda:
        raise HTTPException(status_code=404, detail="La demanda no existe.")
    
    # CONTROL DE SEGURIDAD VIP: Permite descargar si es el dueño O si es Administrador
    es_administrador = getattr(current_user, "es_admin", False)
    if demanda.usuario_id != current_user.id and not es_administrador:
        raise HTTPException(status_code=403, detail="No tenés autorización para descargar este documento.")
    
    referencia_archivo = getattr(demanda, 'archivo_generado', None) or getattr(demanda, 'ruta_archivo', None)
    
    if not referencia_archivo:
        raise HTTPException(status_code=404, detail="Referencia de archivo no encontrada.")

    # SISTEMA DE RETROCOMPATIBILIDAD
    if "demandas_generadas" in referencia_archivo:
        if os.path.exists(referencia_archivo):
            return FileResponse(
                path=referencia_archivo,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename=os.path.basename(referencia_archivo)
            )
        else:
            raise HTTPException(status_code=404, detail="El archivo antiguo no se encuentra en el servidor local.")

    # --- MAGIA DE GOOGLE CLOUD PARA NUEVAS DEMANDAS ---
    try:
        url_segura = generate_signed_url(referencia_archivo, expiration_minutes=5)
        return RedirectResponse(url=url_segura)
    except Exception as e:
        print(f"Error de GCP: {e}")
        raise HTTPException(status_code=500, detail="Error al conectar con la bóveda de seguridad en la nube.")
    

@app.get("/plantillas", response_model=List[schemas.PlantillaOut], summary="Listar plantillas de demandas disponibles")
def obtener_plantillas(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(verificar_suscripcion_activa)
):
    """
    Devuelve la lista de todas las plantillas activas disponibles en el sistema 
    para que el usuario elija cuál utilizar al generar su demanda.
    """
    plantillas = db.query(models.Plantilla).filter(models.Plantilla.activa == True).all()
    return plantillas
    

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

    # 3. Retorno del 100% de los datos mapeados
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
                "FechaMedica": datos.FechaMedica,
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
                "Porcentaje_Psicologico_Ingresado": datos.PorcentajeDanoPsicologico,
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
    ahora = datetime.now(timezone.utc)

    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == current_user.id).first()

    if not suscripcion:
        suscripcion = models.Suscripcion(usuario_id=current_user.id)
        db.add(suscripcion)

    suscripcion.plan = plan
    suscripcion.demandas_restantes = (suscripcion.demandas_restantes or 0) + demandas
    suscripcion.activa = True
    suscripcion.fecha_inicio = ahora
    suscripcion.fecha_expiracion = ahora + timedelta(days=30)

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


@app.post("/crear-preferencia-suscripcion/", summary="Crear preferencia de pago en Mercado Pago")
def crear_preferencia_suscripcion(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    try:
        access_token = os.getenv("MP_ACCESS_TOKEN")
        if not access_token:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="El Token de Mercado Pago no está configurado en las variables de entorno."
            )

        sdk = mercadopago.SDK(access_token)
        base_url = os.getenv("BASE_URL", "https://snide-uranium-hungrily.ngrok-free.dev")

        preference_data = {
            "items": [
                {
                    "title": "Suscripción Mensual - SaaS Demandas Legales",
                    "quantity": 1,
                    "currency_id": "ARS",
                    "unit_price": 500000.0  # Ajustar este monto al precio real final
                }
            ],
            "payer": {
                "email": current_user.email
            },
            "back_urls": {
                "success": f"{base_url}/pago-exitoso",
                "failure": f"{base_url}/pago-fallido",
                "pending": f"{base_url}/pago-pendiente"
            },
            "auto_return": "approved",
            "notification_url": f"{base_url}/webhook-mercadopago/",
            "external_reference": str(current_user.id)
        }

        preference_response = sdk.preference().create(preference_data)
        preference = preference_response.get("response")

        if not preference or "init_point" not in preference:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Error al crear la preferencia en Mercado Pago: {preference_response}"
            )

        return {
            "init_point": preference["init_point"],
            "sandbox_init_point": preference.get("sandbox_init_point"),
            "preference_id": preference.get("id")
        }

    except HTTPException as http_ex:
        raise http_ex
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error interno del servidor: {str(e)}"
        )


@app.post("/webhook-mercadopago/", summary="Webhook de notificaciones para Mercado Pago")
async def mercadopago_webhook(request: Request, db: Session = Depends(get_db)):
    try:
        body = await request.json()
        print("-> Webhook recibido de Mercado Pago:", body)

        topic = body.get("type") or body.get("topic")
        action = body.get("action")
        payment_id = None

        if action and "payment" in action:
            data = body.get("data", {})
            payment_id = data.get("id")
        elif topic == "payment":
            payment_id = body.get("id") or body.get("data", {}).get("id")
        elif "resource" in body:
            resource_url = body.get("resource")
            if "payments" in resource_url:
                payment_id = resource_url.split("/")[-1]

        if not payment_id:
            return {"status": "ignored", "message": "No se encontró el ID de pago"}

        mp_access_token = os.getenv("MP_ACCESS_TOKEN")
        if not mp_access_token:
            print("⚠️ Token de Mercado Pago no configurado.")
            return {"status": "error", "message": "Configuración de credenciales incompleta"}

        async with httpx.AsyncClient() as client:
            headers = {"Authorization": f"Bearer {mp_access_token}"}
            response = await client.get(f"https://api.mercadopago.com/v1/payments/{payment_id}", headers=headers)
            
            if response.status_code != 200:
                return {"status": "error", "message": "No se pudo consultar el pago en Mercado Pago"}
            
            payment_data = response.json()

        status_pago = payment_data.get("status")
        status_detail = payment_data.get("status_detail")
        external_reference = payment_data.get("external_reference")

        print(f"💰 Pago {payment_id} | Estado: {status_pago} | Detalle: {status_detail} | Ref Usuario: {external_reference}")

        if not external_reference:
            return {"status": "ignored", "message": "El pago no tiene un external_reference asociado"}

        user = db.query(models.Usuario).filter(models.Usuario.id == int(external_reference)).first()
        if not user:
            print(f"⚠️ Usuario ID {external_reference} no encontrado.")
            return {"status": "error", "message": "Usuario no encontrado"}

        suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == user.id).first()
        if not suscripcion:
            suscripcion = models.Suscripcion(usuario_id=user.id)
            db.add(suscripcion)

        if status_pago == "approved":
            ahora = datetime.now(timezone.utc)
            suscripcion.plan = "Pro"
            suscripcion.activa = True
            suscripcion.demandas_restantes = (suscripcion.demandas_restantes or 0) + 50
            suscripcion.fecha_inicio = ahora
            suscripcion.fecha_expiracion = ahora + timedelta(days=30)
            
            db.commit()
            print(f"✅ [DB] Suscripción activada y renovada para el usuario ID: {user.id}")

        elif status_pago in ["pending", "in_process", "rejected"]:
            suscripcion.activa = False
            db.commit()
            print(f"❌ [DB] Suscripción inactiva/fallida (Estado: {status_pago}) para usuario ID: {user.id}")

        return {"status": "success"}

    except Exception as e:
        db.rollback()
        print(f"❌ Error crítico en webhook: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
    

@app.get("/admin/estadisticas", summary="Estadísticas globales para el panel de administración")
def obtener_estadisticas_admin(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    if not getattr(current_user, "es_admin", False):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Acceso denegado. Se requieren permisos de administrador."
        )

    total_usuarios = db.query(models.Usuario).count()
    total_demandas = db.query(models.DemandaGenerada).count()
    
    # 1. Total para la tarjeta (incluye las pruebas gratuitas)
    suscripciones_activas = db.query(models.Suscripcion).filter(
        models.Suscripcion.activa == True
    ).count()

    # 2. Filtro de pagas para el cálculo (excluimos el plan "Free" o "Prueba")
    suscripciones_pagas = db.query(models.Suscripcion).filter(
        models.Suscripcion.activa == True,
        models.Suscripcion.plan != "Free" 
    ).count()

    # 3. Cálculo de ingresos (Modifica el 25000 por tu tarifa real)
    tarifa_mensual = 25000 
    ingresos_estimados = suscripciones_pagas * tarifa_mensual

    return {
        "total_usuarios": total_usuarios,
        "suscripciones_activas": suscripciones_activas,
        "demandas_generadas": total_demandas,
        "ingresos_mensuales": ingresos_estimados
    }


@app.get("/suscripcion/estado", summary="Verificar el estado de la suscripción actual")
def verificar_estado_suscripcion(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    suscripcion = db.query(models.Suscripcion).filter(
        models.Suscripcion.usuario_id == current_user.id
    ).first()

    if not suscripcion:
        return {
            "tiene_suscripcion": False,
            "activa": False,
            "mensaje": "No posees ningún plan o suscripción registrada."
        }

    if suscripcion.fecha_expiracion and suscripcion.fecha_expiracion < datetime.utcnow():
        if suscripcion.activa:
            suscripcion.activa = False
            db.commit()
            db.refresh(suscripcion)

    return {
        "tiene_suscripcion": True,
        "activa": suscripcion.activa,
        "fecha_expiracion": suscripcion.fecha_expiracion,
        "usuario_email": current_user.email
    }


@app.get("/pago-exitoso", summary="Maneja el retorno de un pago exitoso")
def pago_exitoso(
    external_reference: str = Query(None),
    collection_status: str = Query(None),
    payment_id: str = Query(None),
    db: Session = Depends(get_db)
):
    if external_reference and external_reference != "None":
        try:
            user_id = int(external_reference)
            ahora = datetime.utcnow()
            expiracion = ahora + timedelta(days=30)

            suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == user_id).first()
            
            if suscripcion:
                suscripcion.activa = True
                suscripcion.plan = "Premium"
                suscripcion.demandas_restantes = 50
                suscripcion.fecha_inicio = ahora
                suscripcion.fecha_expiracion = expiracion
                db.commit()
            else:
                nueva_suscripcion = models.Suscripcion(
                    usuario_id=user_id,
                    plan="Premium",
                    demandas_restantes=50,
                    activa=True,
                    fecha_inicio=ahora,
                    fecha_expiracion=expiracion
                )
                db.add(nueva_suscripcion)
                db.commit()
                
            print(f"Suscripción actualizada con éxito para el usuario {user_id} hasta {expiracion}")

        except Exception as e:
            print(f"Error al actualizar la suscripción en la base de datos: {e}")
            db.rollback()

    url_retorno = "http://127.0.0.1:5500/frontend_demandas/index.html"
    
    response = HTMLResponse(content=f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Pago Exitoso - SaaS Legal</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <meta http-equiv="refresh" content="3;url={url_retorno}" />
    </head>
    <body class="bg-light d-flex align-items-center justify-content-center vh-100">
        <div class="card p-4 text-center shadow-sm" style="max-width: 450px;">
            <div class="text-success mb-3" style="font-size: 3rem;">✓</div>
            <h3 class="fw-bold text-dark">¡Pago Confirmado!</h3>
            <p class="text-muted small">Tu suscripción ha sido activada correctamente.</p>
            <p class="text-secondary small">Redirigiendo al panel de control...</p>
            <a href="{url_retorno}" class="btn btn-primary btn-sm mt-2">Volver Manualmente</a>
        </div>
    </body>
    </html>
    """)
    
    response.headers["ngrok-skip-browser-warning"] = "true"
    return response


@app.post("/admin/plantillas", response_model=schemas.PlantillaResponse, status_code=status.HTTP_201_CREATED, summary="Registrar nueva plantilla")
def registrar_plantilla(
    plantilla: schemas.PlantillaCreate,
    db: Session = Depends(get_db),
    admin_user: models.Usuario = Depends(get_current_admin_user)
):
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
    plantilla = db.query(models.Plantilla).filter(models.Plantilla.id == plantilla_id).first()
    if not plantilla:
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    
    plantilla.activa = estado_data.activa
    db.commit()
    
    estado_texto = "habilitada" if estado_data.activa else "deshabilitada"
    return {"status": "success", "mensaje": f"La plantilla ha sido {estado_texto} correctamente."}


@app.post("/extraer-datos-acta/", summary="Extraer datos del acta de mediación con IA")
async def extraer_datos_acta(
    archivo: UploadFile = File(...),
    current_user: models.Usuario = Depends(verificar_suscripcion_activa)
):
    print(f"\n▶️ [DEBUG IA] 1. Recibiendo archivo: {archivo.filename}")
    try:
        contenido_archivo = await archivo.read()
        MAX_FILE_SIZE = 5 * 1024 * 1024 # 5 MB
        
        if len(contenido_archivo) > MAX_FILE_SIZE:
            print("❌ [DEBUG IA] Archivo supera los 5MB.")
            raise HTTPException(status_code=413, detail="El archivo es demasiado grande (Máximo 5 MB).")

        mime_type = archivo.content_type
        print(f"▶️ [DEBUG IA] 2. Tipo de archivo detectado: {mime_type}")
        
        formatos_soportados = ["application/pdf", "image/jpeg", "image/png", "image/webp"]
        if mime_type not in formatos_soportados:
            print("❌ [DEBUG IA] Formato no soportado.")
            raise HTTPException(status_code=400, detail="Formato no soportado. Sube un PDF o imagen.")

        print("▶️ [DEBUG IA] 3. Conectando con Gemini (modelo gemini-1.5-flash)...")
        modelo = genai.GenerativeModel('gemini-flash-latest')
        
        prompt = """
        Eres un asistente legal experto en analizar actas de mediación.
        Lee el documento adjunto y extrae EXCLUSIVAMENTE los siguientes datos.
        
        Devuelve la respuesta ESTRICTAMENTE en formato JSON válido.
        NO uses bloques de código markdown (como ```json). 
        NO agregues ningún texto antes ni después del JSON.
        
        Utiliza exactamente estas claves:
        {
            "DniActor": "",
            "NombreActor": "",
            "DomicilioActor": "",
            "DniDemandado": "",
            "NombreDemandado": "",
            "DomicilioDemandado": "",
            "NombreAseguradora": "",
            "CuitAseguradora": "",
            "DomicilioAseguradora": ""
        }
        
        Si no encuentras un dato específico en el documento, deja el valor como un string vacío "".
        Asegúrate de limpiar los números de DNI y CUIT quitando puntos si los tuvieran.
        """

        respuesta = await modelo.generate_content_async([
            {"mime_type": mime_type, "data": contenido_archivo}, 
            prompt
        ])
        
        print("✅ [DEBUG IA] 4. ¡Respuesta de Gemini recibida con éxito!")

        texto_limpio = respuesta.text.replace("```json", "").replace("```", "").strip()
        datos_extraidos = json.loads(texto_limpio)
        print("✅ [DEBUG IA] 5. JSON procesado correctamente.")

        return {
            "status": "success",
            "mensaje": "Datos extraídos correctamente.",
            "datos": datos_extraidos
        }

    except json.JSONDecodeError:
        print("❌ [DEBUG IA] Error: La respuesta de la IA no era un JSON válido.")
        raise HTTPException(status_code=500, detail="La IA no devolvió un formato JSON válido. Intenta nuevamente.")
    except Exception as e:
        print(f"❌ [DEBUG IA] Error crítico: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error interno al procesar el acta con IA: {str(e)}")
        

@app.get("/modelos-ia", summary="Listar modelos permitidos por mi API Key")
def listar_modelos():
    try:
        modelos_permitidos = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                modelos_permitidos.append(m.name)
        
        return {
            "status": "success", 
            "cantidad": len(modelos_permitidos),
            "modelos": modelos_permitidos
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al conectar con Google: {str(e)}")


def require_admin(current_user: models.Usuario = Depends(get_current_user)):
    if not current_user.es_admin:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Acceso denegado: Se requieren permisos de administrador."
        )
    return current_user


@app.get("/admin/metricas", summary="Obtener estadísticas generales para el Admin")
def obtener_metricas_admin(
    db: Session = Depends(get_db),
    admin: models.Usuario = Depends(require_admin)
):
    total_usuarios = db.query(models.Usuario).count()
    suscripciones_activas = db.query(models.Suscripcion).filter(models.Suscripcion.activa == True).count()
    
    total_demandas = db.query(models.DemandaGenerada).count() if hasattr(models, 'DemandaGenerada') else 0

    return {
        "total_usuarios": total_usuarios,
        "suscripciones_activas": suscripciones_activas,
        "total_demandas": total_demandas
    }


@app.get("/admin/usuarios", summary="Obtener lista de usuarios y sus estados")
def listar_usuarios_admin(
    db: Session = Depends(get_db),
    admin: models.Usuario = Depends(require_admin)
):
    usuarios = db.query(models.Usuario).all()
    resultado = []

    for u in usuarios:
        suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == u.id).first()
        resultado.append({
            "id": u.id,
            "email": u.email,
            "es_admin": u.es_admin,
            "activo": u.activo,
            "suscripcion_activa": suscripcion.activa if suscripcion else False,
            "plan": suscripcion.plan if suscripcion else "Sin Plan",
            "demandas_restantes": suscripcion.demandas_restantes if suscripcion else 0
        })

    return resultado


@app.put("/admin/usuarios/{usuario_id}/toggle-suscripcion")
def toggle_suscripcion_usuario(
    usuario_id: int,
    db: Session = Depends(get_db),
    admin: models.Usuario = Depends(require_admin)
):
    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == usuario_id).first()
    if not suscripcion:
        suscripcion = models.Suscripcion(
            usuario_id=usuario_id,
            plan="Premium (Manual)",
            demandas_restantes=50,
            activa=True,
            fecha_inicio=datetime.utcnow(),
            fecha_expiracion=datetime.utcnow() + timedelta(days=30)
        )
        db.add(suscripcion)
    else:
        suscripcion.activa = not suscripcion.activa
        if suscripcion.activa:
            suscripcion.demandas_restantes = 50
            suscripcion.fecha_expiracion = datetime.utcnow() + timedelta(days=30)

    db.commit()
    return {"mensaje": f"Estado de la suscripción actualizado a {suscripcion.activa}"}

@app.post("/auth/olvide-password", summary="Solicitar restablecimiento de contraseña")
async def solicitar_recuperacion(
    email: str = Form(...),
    background_tasks: BackgroundTasks = None,
    db: Session = Depends(get_db)
):
    usuario = db.query(models.Usuario).filter(models.Usuario.email == email).first()
    
    if usuario:
        token = serializer.dumps(usuario.email, salt="reset-password-salt")
        
        frontend_url = "https://autodemandas.com.ar"
        link_recuperacion = f"{frontend_url}/reset-password.html?token={token}"
        
        # Función interna que Resend ejecutará en segundo plano
        def enviar_correo_resend():
            try:
                resend.Emails.send({
                    "from": "SaaS Legal <soporte@autodemandas.com.ar>", 
                    "to": [email],
                    "subject": "Restablecimiento de Contraseña - SaaS Legal",
                    "html": f"""
                    <h3>Restablecimiento de Contraseña</h3>
                    <p>Haz clic en el siguiente enlace para continuar:</p>
                    <p><a href="{link_recuperacion}">Restablecer mi contraseña</a></p>
                    """
                })
                print("✅ [API HTTP] Correo de recuperación enviado exitosamente vía Resend.")
            except Exception as e:
                print(f"⚠️ [ERROR RESEND]: {e}")

        # Ejecutamos la tarea sin bloquear la respuesta al usuario
        if background_tasks:
            background_tasks.add_task(enviar_correo_resend)
        else:
            enviar_correo_resend()

    return {"mensaje": "Si el correo está registrado, recibirás un enlace de recuperación a la brevedad."}

@app.post("/auth/reset-password", summary="Cambiar la contraseña usando el token")
def resetear_password(
    token: str = Form(...),
    nueva_password: str = Form(...),
    db: Session = Depends(get_db)
):
    try:
        email = serializer.loads(token, salt="reset-password-salt", max_age=900)
    except SignatureExpired:
        raise HTTPException(status_code=400, detail="El enlace ha expirado. Solicita uno nuevo.")
    except BadSignature:
        raise HTTPException(status_code=400, detail="El enlace de recuperación es inválido.")

    usuario = db.query(models.Usuario).filter(models.Usuario.email == email).first()
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado.")

    usuario.hashed_password = pwd_context.hash(nueva_password)
    db.commit()

    return {"mensaje": "Contraseña actualizada exitosamente. Ya puedes iniciar sesión."}


@app.get("/mis-demandas", summary="Listar todas las demandas del equipo (Titular) o propias (Asistente)")
def listar_mis_demandas(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(verificar_suscripcion_activa)
):
    # 1. Lógica Jerárquica: Determinar qué demandas puede ver
    if getattr(current_user, 'cuenta_madre_id', None) is None:
        # Es Titular: obtenemos su ID y los IDs de sus asistentes
        asistentes = db.query(models.Usuario.id).filter(models.Usuario.cuenta_madre_id == current_user.id).all()
        ids_permitidos = [current_user.id] + [a.id for a in asistentes]
    else:
        # Es Asistente: solo ve las suyas
        ids_permitidos = [current_user.id]

    # 2. Búsqueda con JOIN para traer el email del creador
    demandas_con_creador = db.query(models.DemandaGenerada, models.Usuario.email).join(
        models.Usuario, models.DemandaGenerada.usuario_id == models.Usuario.id
    ).filter(
        models.DemandaGenerada.usuario_id.in_(ids_permitidos)
    ).all()

    lista_demandas = []
    for d, email_creador in demandas_con_creador:
        lista_demandas.append({
            "id": d.id,
            "nombre_actor": getattr(d, "nombre_actor", "Sin nombre"),
            "dni_actor": getattr(d, "dni_actor", "-"),
            "estado_operativo": getattr(d, "estado_operativo", "Generada"),
            "fecha_creacion": d.fecha_creacion if hasattr(d, "fecha_creacion") else "N/A",
            "creado_por": email_creador, # <-- Dato clave para el Titular
            "download_url": f"https://saas-demandas-legal.onrender.com/descargar-demanda/{d.id}"
        })

    return {
        "cantidad": len(lista_demandas),
        "demandas": lista_demandas
    }
@app.get("/historial", response_model=List[schemas.DemandaHistorialOut], summary="Obtener historial de demandas jerárquico")
def obtener_historial(
    limit: int = 10,
    skip: int = 0,
    nombre_actor: Optional[str] = None,
    dni_actor: Optional[int] = None,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(verificar_suscripcion_activa)
):
    # Lógica Jerárquica B2B
    if getattr(current_user, 'cuenta_madre_id', None) is None:
        asistentes = db.query(models.Usuario.id).filter(models.Usuario.cuenta_madre_id == current_user.id).all()
        ids_permitidos = [current_user.id] + [a.id for a in asistentes]
    else:
        ids_permitidos = [current_user.id]

    query = db.query(models.DemandaGenerada).filter(models.DemandaGenerada.usuario_id.in_(ids_permitidos))

    if nombre_actor:
        query = query.filter(models.DemandaGenerada.nombre_actor.ilike(f"%{nombre_actor}%"))
    if dni_actor:
        query = query.filter(models.DemandaGenerada.dni_actor == dni_actor)

    demandas = query.order_by(models.DemandaGenerada.fecha_creacion.desc()).offset(skip).limit(limit).all()
    return demandas


@app.get("/registrar-plantillas")
def registrar_plantillas(db: Session = Depends(get_db)):
    p1 = db.query(models.Plantilla).filter(models.Plantilla.id == 1).first()
    if not p1:
        nueva_p1 = models.Plantilla(
            id=1, 
            nombre="Auto vs Moto", 
            categoria="Accidentes de Tránsito", 
            descripcion="Demanda por accidente entre auto y moto",
            ruta_archivo="templates/Borrador_Demanda_Auto_Moto.docx", 
            activa=True
        )
        db.add(nueva_p1)
    else:
        p1.ruta_archivo = "templates/Borrador_Demanda_Auto_Moto.docx"
    
    p2 = db.query(models.Plantilla).filter(models.Plantilla.id == 2).first()
    if not p2:
        nueva_p2 = models.Plantilla(
            id=2, 
            nombre="Auto vs Auto", 
            categoria="Accidentes de Tránsito", 
            descripcion="Demanda por accidente entre dos autos",
            ruta_archivo="templates/Borrador_Demanda_Auto_Auto.docx", 
            activa=True
        )
        db.add(nueva_p2)
    else:
        p2.ruta_archivo = "templates/Borrador_Demanda_Auto_Auto.docx"
    
    db.commit()
    return {"mensaje": "¡Las plantillas se registraron y actualizaron correctamente en la base de datos con la ruta 'templates/'!"}


@app.get("/activar-prueba/{email}")
def activar_prueba(email: str, db: Session = Depends(get_db)):
    usuario = db.query(models.Usuario).filter(models.Usuario.email == email).first()
    
    if not usuario:
        return {"error": f"Usuario con correo {email} no encontrado."}
        
    suscripcion = db.query(models.Suscripcion).filter(models.Suscripcion.usuario_id == usuario.id).first()
    
    if not suscripcion:
        nueva_suscripcion = models.Suscripcion(
            usuario_id=usuario.id,
            activa=True,
            demandas_restantes=10, 
            fecha_expiracion=datetime.utcnow() + timedelta(days=30)
        )
        db.add(nueva_suscripcion)
    else:
        suscripcion.activa = True
        suscripcion.demandas_restantes = 10
        suscripcion.fecha_expiracion = datetime.utcnow() + timedelta(days=30)
        
    if hasattr(usuario, 'suscripcion_activa'):
        usuario.suscripcion_activa = True
        
    db.commit()
    return {"mensaje": f"¡Éxito! Se le otorgó una suscripción Premium con 10 demandas de prueba a {email}."}


@app.get("/admin/usuarios/{usuario_id}/demandas", summary="Ver historial de demandas de un usuario")
def ver_demandas_usuario_admin(
    usuario_id: int,
    db: Session = Depends(get_db),
    admin: models.Usuario = Depends(require_admin)
):
    demandas = db.query(models.DemandaGenerada).filter(
        models.DemandaGenerada.usuario_id == usuario_id
    ).order_by(models.DemandaGenerada.fecha_creacion.desc()).all()

    lista_demandas = []
    for d in demandas:
        lista_demandas.append({
            "id": d.id,
            "nombre_actor": getattr(d, "nombre_actor", "Sin nombre"),
            "dni_actor": getattr(d, "dni_actor", "-"),
            "estado_operativo": getattr(d, "estado_operativo", "Generada"),
            "fecha_creacion": d.fecha_creacion if hasattr(d, "fecha_creacion") else "N/A",
            "download_url": f"[https://saas-demandas-legal.onrender.com/descargar-demanda/](https://saas-demandas-legal.onrender.com/descargar-demanda/){d.id}"
        })

    return lista_demandas


@app.put("/admin/usuarios/{usuario_id}/toggle-rol", summary="Cambiar rol de un usuario (Admin/Usuario)")
def toggle_rol_usuario(
    usuario_id: int,
    db: Session = Depends(get_db),
    admin: models.Usuario = Depends(require_admin)
):
    usuario = db.query(models.Usuario).filter(models.Usuario.id == usuario_id).first()
    
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    
    if usuario.id == admin.id:
        raise HTTPException(
            status_code=400, 
            detail="No puedes quitarte el rol de Administrador a tu propia cuenta activa."
        )

    usuario.es_admin = not usuario.es_admin
    db.commit()
    
    return {"status": "success", "mensaje": "Rol actualizado correctamente", "es_admin": usuario.es_admin}

@app.put("/admin/usuarios/{usuario_id}/toggle-activo", summary="Habilitar o inhabilitar a un usuario")
def toggle_activo_usuario(
    usuario_id: int,
    db: Session = Depends(get_db),
    admin: models.Usuario = Depends(require_admin)
):
    usuario = db.query(models.Usuario).filter(models.Usuario.id == usuario_id).first()
    
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    
    # Evitamos que el administrador principal se bloquee a sí mismo por error
    if usuario.id == admin.id:
        raise HTTPException(
            status_code=400, 
            detail="Por seguridad, no puedes inhabilitar tu propia cuenta."
        )

    # Invertimos el estado (si estaba activo pasa a inactivo, y viceversa)
    usuario.activo = not usuario.activo
    db.commit()
    
    estado_texto = "habilitado" if usuario.activo else "inhabilitado"
    return {
        "status": "success", 
        "mensaje": f"El usuario ha sido {estado_texto} correctamente.", 
        "activo": usuario.activo
    }

@app.get("/equipo", summary="Obtener lista de asistentes del titular")
def ver_mi_equipo(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # Buscamos a todos los usuarios cuya "cuenta madre" sea el ID del usuario logueado
    asistentes = db.query(models.Usuario).filter(
        models.Usuario.cuenta_madre_id == current_user.id
    ).all()
    
    lista_equipo = []
    for asistente in asistentes:
        lista_equipo.append({
            "id": asistente.id,
            "email": asistente.email,
            "activo": asistente.activo,
            "rol_estudio": getattr(asistente, 'rol_estudio', 'asistente'),
            "fecha_creacion": asistente.fecha_creacion
        })
        
    return lista_equipo

# 1. Definimos la estructura de datos que enviará el frontend
class AsistenteCreate(BaseModel):
    email: EmailStr
    password: str

# 2. Endpoint para registrar al asistente y enlazarlo automáticamente
@app.post("/equipo", summary="Crear y agregar un nuevo asistente al equipo")
def agregar_asistente(
    datos: AsistenteCreate,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # Verificamos que el correo no exista ya en la base de datos
    usuario_existente = db.query(models.Usuario).filter(models.Usuario.email == datos.email).first()
    if usuario_existente:
        raise HTTPException(status_code=400, detail="Este correo ya está registrado en el sistema.")

    # Hasheamos la contraseña (Asegúrate de que la función de hash se llame así en tu código, 
    # a veces suele estar en utils.get_password_hash o security.get_password_hash)
    from security import get_password_hash # Ajustá la importación si la tuya se llama distinto
    password_hasheada = get_password_hash(datos.password)

    # Creamos el usuario y lo atamos al Titular
    nuevo_asistente = models.Usuario(
        email=datos.email,
        hashed_password=password_hasheada,
        cuenta_madre_id=current_user.id, # <-- Acá ocurre la magia del enlace B2B
        rol_estudio="asistente",
        activo=True,
        es_admin=False
    )
    
    db.add(nuevo_asistente)
    db.commit()
    db.refresh(nuevo_asistente)

    return {
        "status": "success",
        "mensaje": f"Asistente {nuevo_asistente.email} agregado exitosamente.",
        "asistente_id": nuevo_asistente.id
    }
@app.delete("/equipo/{asistente_id}", summary="Desvincular a un asistente del equipo")
def desvincular_asistente(
    asistente_id: int,
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    # Buscamos al asistente asegurándonos de que su jefe sea el usuario logueado
    asistente = db.query(models.Usuario).filter(
        models.Usuario.id == asistente_id,
        models.Usuario.cuenta_madre_id == current_user.id
    ).first()

    if not asistente:
        raise HTTPException(status_code=404, detail="Asistente no encontrado o no pertenece a tu equipo.")

    # Lo desvinculamos del titular y por seguridad inhabilitamos la cuenta
    asistente.cuenta_madre_id = None
    asistente.activo = False
    
    db.commit()
    
    return {"status": "success", "mensaje": "Asistente desvinculado e inhabilitado correctamente."}
# ==========================================
# RUTAS DE CONTACTO (LANDING PAGE)
# ==========================================

class ContactoRequest(BaseModel):
    nombre: str
    email: str
    mensaje: str

@app.post("/contacto", summary="Procesar formulario de contacto desde Landing Page")
def procesar_contacto(datos: ContactoRequest, background_tasks: BackgroundTasks):
    correos_destino = ["martin_graneros@hotmail.com"]

    cuerpo_mensaje = f"""
    <div style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #0d6efd;">Nueva consulta desde la Landing Page</h2>
        <p><strong>👤 Nombre:</strong> {datos.nombre}</p>
        <p><strong>📧 Email de contacto:</strong> {datos.email}</p>
        <hr>
        <p><strong>💬 Mensaje:</strong></p>
        <blockquote style="background: #f8f9fa; padding: 15px; border-left: 5px solid #0d6efd; border-radius: 5px;">
            {datos.mensaje}
        </blockquote>
        <br>
        <p style="font-size: 12px; color: #6c757d;">Este es un mensaje automático de SaaS Demandas Legal.</p>
    </div>
    """

    def enviar_contacto_resend():
        try:
            resend.Emails.send({
                "from": "SaaS Legal <onboarding@resend.dev>",
                "to": correos_destino,
                "subject": f"NUEVO CONTACTO - SaaS Legal - {datos.nombre}",
                "html": cuerpo_mensaje
            })
            print("✅ [API HTTP] Correo de contacto enviado exitosamente vía Resend.")
        except Exception as e:
            print(f"⚠️ [ERROR RESEND]: {e}")

    if background_tasks:
        background_tasks.add_task(enviar_contacto_resend)
    else:
        enviar_contacto_resend()

    return {"status": "success", "mensaje": "Tu mensaje ha sido recibido. Te contactaremos pronto."}