from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx import Document
from num2words import num2words
from docxtpl import DocxTemplate, RichText
import os
from fastapi.middleware.cors import CORSMiddleware # 1. Importar el middleware
from fastapi.responses import FileResponse
from fastapi import HTTPException
from database import inicializar_db, guardar_demanda, listar_historial, obtener_demanda_por_id
import os

app = FastAPI(title="SaaS Demandas Legal API", version="0.3.0")
inicializar_db()
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

@app.post("/generar-demanda/")
def generar_demanda(datos: DatosDemanda, background_tasks: BackgroundTasks):
    ruta_plantilla = "templates/Borrador_Demanda_Auto_Moto.docx"
    nombre_limpio = datos.NombreActor.replace(' ', '_')
    ruta_salida = f"temp_{nombre_limpio}.docx"

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
        doc.render(datos_procesados) # Renderiza texto plano, más seguro a prueba de fallos
        doc.save(ruta_salida)

        # 💾 Guardar en Base de Datos para el Historial
        payload_dict = datos.model_dump() if hasattr(datos, "model_dump") else datos.dict()
        payload_dict["MontoTotal"] = liquidacion_total  # Guardamos el total calculado
        
        id_guardado = guardar_demanda(payload=payload_dict, ruta_archivo=ruta_salida)
        print(f"Demanda guardada en BD con ID: {id_guardado}")
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error interno al procesar el documento: {str(e)}")

    return FileResponse(
        path=ruta_salida, 
        filename=f"demanda_{nombre_limpio}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/historial", summary="Obtener el historial de demandas")
def get_historial():
    """Devuelve la lista con todas las demandas generadas previamente."""
    return listar_historial()


@app.get("/descargar-demanda/{demanda_id}", summary="Re-descargar una demanda del historial")
def descargar_demanda_historica(demanda_id: int):
    """Busca el archivo en el historial por su ID y lo entrega para descarga."""
    registro = obtener_demanda_por_id(demanda_id)
    
    if not registro:
        raise HTTPException(status_code=404, detail="No se encontró el registro en la base de datos.")
    
    ruta_archivo = registro["ruta_archivo"]
    
    if not os.path.exists(ruta_archivo):
        raise HTTPException(status_code=404, detail="El archivo físico .docx ya no existe en el servidor.")
        
    nombre_descarga = f"Demanda_{registro['nombre_actor']}_vs_{registro['nombre_demandado']}.docx".replace(" ", "_")

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