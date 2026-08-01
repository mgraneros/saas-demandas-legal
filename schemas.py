from datetime import datetime
from typing import Optional, List
from pydantic import BaseModel, EmailStr

# ==========================================
# 1. ESQUEMAS DE AUTENTICACIÓN Y USUARIO
# ==========================================

class UsuarioCreate(BaseModel):
    email: EmailStr
    password: str
    nombre_estudio: Optional[str] = None


class UsuarioResponse(BaseModel):
    id: int
    email: str
    nombre_estudio: Optional[str] = None

    class Config:
        from_attributes = True


class UsuarioOut(BaseModel):
    id: int
    email: EmailStr

    class Config:
        from_attributes = True


class Token(BaseModel):
    access_token: str
    token_type: str


class TokenData(BaseModel):
    email: Optional[str] = None


# ==========================================
# 2. ESQUEMAS DE DEMANDAS
# ==========================================

class DatosDemanda(BaseModel):
    plantilla_id: int = 1
    
    # 1. Datos del Actor
    NombreActor: str
    DniActor: int
    DomicilioActor: str
    VehiculoActor: str
    
    # 2. Datos Demandado y Seguro
    NombreDemandado: str
    DomicilioDemandado: str
    AutoDemandado: str
    NombreAseguradora: str
    CuitAseguradora: str
    DomicilioAseguradora: str
    
    # 3. Hechos y Lesiones
    FechaHecho: str
    LugarHecho: str
    DescripcionHechos: str
    LesionesDetalles: str  
    PorcentajeDanoPsicologico: str
    ListadoSecuelas: str  
    
    # 4. Prueba y Atención
    Intervencion: str
    CentroMedico: str
    CentroMedicoDireccion: str
    TallerNombre: str
    DirecciónTaller: str
    FechaPresupuesto: str
    ListaDocumental: str
    
    # 5. Competencia y Liquidación
    OpcionCompetencia: str
    LiquiDanoMaterialNum: float
    PuntosdeIncapacidad: float


class DemandaHistorialOut(BaseModel):
    id: int
    plantilla_id: Optional[int]
    nombre_actor: Optional[str]
    dni_actor: Optional[int]
    estado_operativo: str
    fecha_creacion: datetime

    class Config:
        from_attributes = True

class PlantillaOut(BaseModel):
    id: int
    nombre: str
    descripcion: Optional[str] = None
    activa: bool

    class Config:
        from_attributes = True