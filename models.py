from sqlalchemy import Column, Integer, String, Float, Boolean, DateTime, ForeignKey, Text, JSON
from sqlalchemy.orm import relationship
from datetime import datetime
from database import Base

# --- TABLA DE USUARIOS ---
class Usuario(Base):
    __tablename__ = "usuarios"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String, unique=True, index=True, nullable=False)
    hashed_password = Column(String, nullable=False)
    nombre_estudio = Column(String, nullable=True)
    activo = Column(Boolean, default=True)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)

    # Relaciones
    demandas = relationship("DemandaGenerada", back_populates="usuario")
    suscripcion = relationship("Suscripcion", back_populates="usuario", uselist=False)
    logs = relationship("AuditoriaLog", back_populates="usuario")


# --- TABLA DE PLANTILLAS Y TIPOS DE DEMANDA ---
class Plantilla(Base):
    __tablename__ = "plantillas"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String, nullable=False)            # Ej: "Demanda Tránsito Auto/Moto"
    categoria = Column(String, nullable=False)         # Ej: "Daños y Perjuicios", "Laboral"
    descripcion = Column(Text, nullable=True)
    ruta_archivo = Column(String, nullable=False)      # Ej: "templates/Borrador_Demanda_Auto_Moto.docx"
    activa = Column(Boolean, default=True)

    # Relaciones
    demandas = relationship("DemandaGenerada", back_populates="plantilla")


# --- TABLA DE DEMANDAS GENERADAS (Actualizada con FK a Plantilla) ---
class DemandaGenerada(Base):
    __tablename__ = "demandas_generadas"

    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"), nullable=False)
    plantilla_id = Column(Integer, ForeignKey("plantillas.id"), nullable=True) # Opcional por si tenés registros viejos

    dni_actor = Column(Integer, nullable=True)
    nombre_actor = Column(String, nullable=True)
    estado_operativo = Column(String, default="Generada")
    ip_origen = Column(String, nullable=True)
    user_agent = Column(String, nullable=True)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)

    # Relaciones
    usuario = relationship("Usuario", back_populates="demandas")
    plantilla = relationship("Plantilla", back_populates="demandas")


# --- TABLA DE SUSCRIPCIONES Y PLANES ---
class Suscripcion(Base):
    __tablename__ = "suscripciones"

    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"), unique=True, nullable=False)
    plan = Column(String, default="Free")             # "Free", "Pro", "Premium"
    demandas_restantes = Column(Integer, default=3)   # Límite mensual (ej: 3 para Free)
    fecha_inicio = Column(DateTime, default=datetime.utcnow)
    fecha_vencimiento = Column(DateTime, nullable=True)

    # Relaciones
    usuario = relationship("Usuario", back_populates="suscripcion")


# --- TABLA DE LOGS DE AUDITORÍA ---
class AuditoriaLog(Base):
    __tablename__ = "auditoria_logs"

    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    accion = Column(String, nullable=False)            # Ej: "LOGIN", "GENERAR_DEMANDA", "DESCARGAR_DOC"
    ip_origen = Column(String, nullable=True)
    detalles = Column(Text, nullable=True)             # Info adicional en formato texto/JSON
    fecha = Column(DateTime, default=datetime.utcnow)

    # Relaciones
    usuario = relationship("Usuario", back_populates="logs")