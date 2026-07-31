from sqlalchemy import Column, Integer, String, DateTime, ForeignKey
from sqlalchemy.orm import relationship
from datetime import datetime
from database import Base

class Usuario(Base):
    __tablename__ = "usuarios"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String, unique=True, index=True, nullable=False)
    hashed_password = Column(String, nullable=False)
    nombre_estudio = Column(String, nullable=True)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)

    # Relación 1 a N con demandas
    demandas = relationship("DemandaGenerada", back_populates="usuario")


class DemandaGenerada(Base):
    __tablename__ = "demanda_generada"

    id = Column(Integer, primary_key=True, index=True)
    dni_actor = Column(Integer, nullable=True, index=True)
    nombre_actor = Column(String, nullable=False)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)
    estado_operativo = Column(String, default="Generada")
    
    ip_origen = Column(String, nullable=True)
    user_agent = Column(String, nullable=True)

    # 🔗 Clave Foránea para Multi-Tenancy (SaaS)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    usuario = relationship("Usuario", back_populates="demandas")