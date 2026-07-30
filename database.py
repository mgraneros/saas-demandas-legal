import sqlite3
import json
from datetime import datetime

DB_NAME = "historial.db"

def obtener_conexion():
    """Crea y devuelve una conexión a la base de datos SQLite."""
    conn = sqlite3.connect(DB_NAME)
    # Permite acceder a las columnas por nombre como si fuera un diccionario
    conn.row_factory = sqlite3.Row
    return conn

def inicializar_db():
    """Crea la tabla 'historial_demandas' si todavía no existe."""
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historial_demandas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_creacion TEXT NOT NULL,
            nombre_actor TEXT NOT NULL,
            nombre_demandado TEXT NOT NULL,
            monto_total REAL,
            payload_json TEXT NOT NULL,
            ruta_archivo TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

def guardar_demanda(payload: dict, ruta_archivo: str) -> int:
    """Guarda un nuevo registro en el historial y devuelve el ID generado."""
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Extraemos valores con fallback por si alguna clave cambia en el JSON
    actor = payload.get("NombreActor") or payload.get("nombre_actor") or "Sin especificar"
    demandado = payload.get("NombreDemandado") or payload.get("nombre_demandado") or "Sin especificar"
    
    # Intentamos convertir el monto a float si viene en el payload
    monto_raw = payload.get("MontoTotal") or payload.get("monto_total") or 0
    try:
        monto = float(monto_raw)
    except (ValueError, TypeError):
        monto = 0.0

    cursor.execute('''
        INSERT INTO historial_demandas 
        (fecha_creacion, nombre_actor, nombre_demandado, monto_total, payload_json, ruta_archivo)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        actor,
        demandado,
        monto,
        json.dumps(payload, ensure_ascii=False),
        ruta_archivo
    ))
    
    nuevo_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return nuevo_id

def listar_historial():
    """Devuelve la lista de todas las demandas guardadas ordenadas por la más reciente."""
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT id, fecha_creacion, nombre_actor, nombre_demandado, monto_total, ruta_archivo 
        FROM historial_demandas 
        ORDER BY id DESC
    ''')
    filas = cursor.fetchall()
    conn.close()
    return [dict(fila) for fila in filas]

def obtener_demanda_por_id(demanda_id: int):
    """Busca un registro específico por su ID."""
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM historial_demandas WHERE id = ?', (demanda_id,))
    fila = cursor.fetchone()
    conn.close()
    return dict(fila) if fila else None