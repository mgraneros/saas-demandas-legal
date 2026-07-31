import os
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage, RichText
from fastapi import BackgroundTasks, Depends, FastAPI, HTTPException, Request, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from jose import JWTError, jwt
from num2words import num2words
from pydantic import BaseModel, EmailStr
from sqlalchemy.orm import Session
import bcrypt
import models
from models import DemandaGenerada

# Módulos propios del proyecto
from database import get_db
import models

# 1. Inicialización de la aplicación FastAPI
app = FastAPI(title="SaaS Demandas Legal API", version="0.3.0")

# 2. Configuración JWT
SECRET_KEY = "tu_clave_secreta_super_segura_aqui_cambiar_en_produccion"
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24  # El token durará 24 horas

# 3. Esquema OAuth2 para Swagger UI
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")

# 4. Funciones de encriptación con bcrypt
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

# 5. Modelos Pydantic para Autenticación
class UsuarioCreate(BaseModel):
    email: EmailStr
    password: str
    nombre_estudio: str | None = None

class UsuarioResponse(BaseModel):
    id: int
    email: str
    nombre_estudio: str | None = None

    class Config:
        from_attributes = True

class Token(BaseModel):
    access_token: str
    token_type: str

# 6. Dependencia para obtener el usuario autenticado
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
    except JWTError:
        raise credentials_exception
        
    usuario = db.query(models.Usuario).filter(models.Usuario.email == email).first()
    if usuario is None:
        raise credentials_exception
    return usuario

# Agregar el middleware de CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Permitir todas las origenes (ajustar según sea necesario)
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 1. Diccionario Modular de Competencias
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

# 2. El modelo ahora solo pide lo estrictamente necesario para calcular
class DatosDemanda(BaseModel):
    NombreActor: str
    DniActor: int
    OpcionCompetencia: int  # 1, 2, o 3 desde el checkbox del frontend
    PuntosdeIncapacidad: float
    LiquiDanoMaterialNum: float
    DomicilioActor: str
    NombreDemandado: str
    DomicilioDemandado: str
    AutoDemandado: str
    FechaHecho: str
    NombreAseguradora: str
    CuitAseguradora: str
    DomicilioAseguradora: str
    DescripcionHechos: str
    LesionesDetalles: str
    ListadoSecuelas: str
    VehiculoActor: str
    TallerNombre: str
    DirecciónTaller: str
    ListaDocumental: str
    CentroMedico: str
    CentroMedicoDireccion: str
    LugarHecho: str
    FechaPresupuesto: str

def formatear_moneda(valor: float) -> str:
    """Convierte un float al formato argentino $ X.XXX,XX"""
    return f"$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def monto_a_letras_legal(monto: float) -> str:
    """
    Convierte un monto numérico a formato de texto legal en mayúsculas.
    """
    # Separamos la parte entera de los decimales
    entero = int(monto)
    decimales = int(round((monto - entero) * 100))
    
    # Convertimos el número a palabras en español
    texto_entero = num2words(entero, lang='es')
    
    # Armamos la estructura legal final
    texto_final = f"{texto_entero.upper()} CON {decimales:02d}/100"
    
    return texto_final

@app.post("/register", response_model=UsuarioResponse, summary="Registrar nuevo usuario")
def registrar_usuario(usuario: UsuarioCreate, db: Session = Depends(get_db)):
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


@app.post("/token", response_model=Token, summary="Iniciar sesión y obtener JWT")
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

@app.post("/generar-demanda/")
def generar_demanda(datos: DatosDemanda, request: Request, db: Session = Depends(get_db), current_user: models.Usuario = Depends(get_current_user)):
    ruta_plantilla = "templates/Borrador_Demanda_Auto_Moto.docx"
    
    # Crear carpeta dedicada para las demandas si no existe
    carpeta_salida = "demandas_generadas"
    os.makedirs(carpeta_salida, exist_ok=True)

    nombre_limpio = datos.NombreActor.replace(' ', '_')
    # Guardar el archivo dentro de la carpeta creada
    ruta_salida = os.path.join(carpeta_salida, f"temp_{nombre_limpio}.docx")

    # 3. Lógica Matemática Dura
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

    # 4. Asignación del Párrafo de Competencia
    texto_competencia = PARRAFOS_COMPETENCIA.get(
        datos.OpcionCompetencia, 
        PARRAFOS_COMPETENCIA[1] # Valor por defecto de seguridad
    )

# 5. Mapeo de variables listas para inyectar en el Word
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
        "ListaDocumental": datos.ListaDocumental,
        "CentroMedico": datos.CentroMedico,
        "CentroMedicoDireccion": datos.CentroMedicoDireccion,
        "LugarHecho": datos.LugarHecho,
        "FechaPresupuesto": datos.FechaPresupuesto
    }

    # ¡ATENCIÓN! HEMOS ELIMINADO EL PASO 6 TEMPORALMENTE (Sin RichText ni amarillo)

    
    try:
        doc = DocxTemplate(ruta_plantilla)

        # 🖼️ TAREA 4: Carga e inyección del logo dinámico
        ruta_logo = "assets/logo_defecto.png"
        if os.path.exists(ruta_logo):
            logo_imagen = InlineImage(doc, ruta_logo, width=Mm(40))
        else:
            logo_imagen = ""
            
        datos_procesados["logo_estudio"] = logo_imagen

        # Renderizado del Word
        doc.render(datos_procesados)
        doc.save(ruta_salida)

        # 🛡️ CAPTURA DE AUDITORÍA Y PERSISTENCIA (SQLAlchemy)
        ip_cliente = request.client.host if request.client else "Desconocida"
        user_agent_cliente = request.headers.get("user-agent", "Desconocido")

        try:
            dni_val = int(datos.DniActor) if hasattr(datos, "DniActor") and datos.DniActor else None
        except (ValueError, TypeError):
            dni_val = None

        nueva_demanda = models.DemandaGenerada(
            usuario_id=current_user.id,  # 👈 Acá vinculamos la demanda al usuario logueado
            dni_actor=dni_val,
            nombre_actor=datos.NombreActor,
            estado_operativo="Generada",
            ip_origen=ip_cliente,
            user_agent=user_agent_cliente
        )
        db.add(nueva_demanda)
        db.commit()
        db.refresh(nueva_demanda)

        print(f"🔒 [AUDITORÍA] Registro #{nueva_demanda.id} guardado. IP: {ip_cliente} | User-Agent: {user_agent_cliente}")

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error interno al procesar el documento: {str(e)}")

    return FileResponse(
        path=ruta_salida,
        filename=f"demanda_{nombre_limpio}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/historial", summary="Obtener el historial de demandas")
def get_historial(
    db: Session = Depends(get_db),
    current_user: models.Usuario = Depends(get_current_user)
):
    """Devuelve la lista con las demandas generadas pertenencientes al usuario autenticado."""
    registros = (
        db.query(models.DemandaGenerada)
        .filter(models.DemandaGenerada.usuario_id == current_user.id)
        .order_by(models.DemandaGenerada.id.desc())
        .all()
    )
    return registros

# 🔄 WORKFLOW: Endpoint para actualizar el estado operativo de una demanda
@app.patch("/demanda/{demanda_id}/estado", summary="Actualizar estado operativo de una demanda")
def actualizar_estado_demanda(demanda_id: int, nuevo_estado: str, db: Session = Depends(get_db)):
    """
    Estados permitidos recomendados: 'Generada', 'Presentada', 'En Notificacion', 'Archivada'
    """
    registro = db.query(DemandaGenerada).filter(DemandaGenerada.id == demanda_id).first()
    
    if not registro:
        raise HTTPException(status_code=404, detail="No se encontró la demanda especificada.")
        
    registro.estado_operativo = nuevo_estado
    db.commit()
    db.refresh(registro)
    
    return {
        "mensaje": f"Estado de la demanda #{demanda_id} actualizado a '{nuevo_estado}' con éxito.",
        "demanda": registro
    }


@app.get("/descargar-demanda/{demanda_id}", summary="Re-descargar una demanda del historial")
def descargar_demanda_historica(demanda_id: int, db: Session = Depends(get_db)):
    registro = db.query(DemandaGenerada).filter(DemandaGenerada.id == demanda_id).first()
    
    if not registro:
        raise HTTPException(status_code=404, detail="No se encontró el registro en la base de datos.")
    
    nombre_limpio = registro.nombre_actor.replace(' ', '_')
    # 📍 Buscamos dentro de la carpeta dedicada:
    ruta_archivo = os.path.join("demandas_generadas", f"temp_{nombre_limpio}.docx")
    
    if not os.path.exists(ruta_archivo):
        raise HTTPException(status_code=404, detail="El archivo físico .docx ya no existe en el servidor.")
        
    nombre_descarga = f"Demanda_{nombre_limpio}.docx"

    return FileResponse(
        path=ruta_archivo,
        filename=nombre_descarga,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    
    # PREVISUALIZACIÓN DE AUDITORÍA
@app.post("/preview-demanda/")
def preview_demanda(datos: DatosDemanda):
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
                "LesionesDetalles": datos.LesionesDetalles,
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