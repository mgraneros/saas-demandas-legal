from datetime import datetime
from sqlalchemy import Column, Integer, String, DateTime
from database import Base

class DemandaGenerada(Base):
    __tablename__ = "demanda_generada"

    # Identificador
    id = Column(Integer, primary_key=True, index=True)

    # Datos del Negocio
    dni_actor = Column(Integer, nullable=True, index=True)
    nombre_actor = Column(String, nullable=False)
    
    # Trazabilidad y Estado Operativo
    fecha_creacion = Column(DateTime, default=datetime.now)
    estado_operativo = Column(String, default="Generada")

    # Auditoría / DevSecOps
    ip_origen = Column(String, nullable=True)
    user_agent = Column(String, nullable=True)