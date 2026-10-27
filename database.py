# ============================================================================
# database.py — Inicialización y Estructura de la Base de Datos
# ============================================================================
# Este archivo es el CORAZÓN de la persistencia de datos del sistema SISGEMOSI.
# Contiene:
#   1. Funciones de conexión a SQLite (get_db, get_biblioteca_db)
#   2. Función init_db() que CREA TODAS LAS TABLAS si no existen
#   3. Migraciones automáticas (agrega columnas nuevas sin perder datos)
#   4. Creación del usuario administrador inicial (admin/admin123)
#   5. Esquema de la matriz de accesos para cargos y sistemas del PAE
#   6. Índices de rendimiento para consultas frecuentes
#
# IMPORTANTE: El sistema usa SQLite (archivo local), NO requiere instalar
# ningún servidor de base de datos. Los archivos .db se crean automáticamente.
#
# Tablas principales:
#   - usuarios_sistema: Usuarios que acceden al sistema (login)
#   - trabajadores: Padrón de personal de la oficina
#   - equipos: Inventario de equipos informáticos, móviles, tokens, etc.
#   - prestamos: Registro de préstamos de equipos a trabajadores
#   - actividades_soporte: Bitácora de actividades de soporte TI
#   - documentos_registrados: Registro de documentos oficiales (informes, memorandos)
#   - reportes_sgd_cortes: Cortes de reportes del SGD (documentos no leídos)
#   - configuraciones_sistema: Configuraciones generales (patrón de numeración, etc.)
#   - certificados_digitales: Tokens y certificados de firma digital
#   - contratos_trabajadores: Archivos PDF de contratos OS/CAS
#   - grabaciones_apowerrec: Registro de grabaciones de pantalla sincronizadas
#   - sistemas_modulos_ut: Catálogo de sistemas del PAE (SIGO, SIGA, SGD, etc.)
#   - perfiles_cargo_ut: Perfiles de acceso por cargo
#   - accesos_perfil_cargo_ut: Matriz cargo×sistema con nivel de acceso
#   - solicitudes_acceso_trabajador_ut: Solicitudes individuales de acceso
#   - biblioteca_pae: Enlaces a eventos/capacitaciones (BD separada)
# ============================================================================

from werkzeug.security import generate_password_hash

import re
import sqlite3


def normalizar_numero_celular(numero):
    """Elimina todos los caracteres no numéricos de un número de celular.
    
    Se usa para normalizar números como '951-234-567' o '+51 951234567'
    a formato puro '951234567', facilitando búsquedas y evitando duplicados.
    
    Args:
        numero: Número de celular en cualquier formato (puede ser None).
    
    Returns:
        String con solo dígitos, o cadena vacía si el input es None/vacío.
    """
    return re.sub(r"\D", "", str(numero or ""))


def ensure_input_validation_schema(conn):
    """Applies database-level checks without rebuilding existing user databases."""
    conn.executescript("""
    CREATE TRIGGER IF NOT EXISTS validar_trabajadores_insert
    BEFORE INSERT ON trabajadores
    WHEN NEW.dni IS NULL OR length(NEW.dni) != 8 OR NEW.dni GLOB '*[^0-9]*'
      OR trim(NEW.nombres) = '' OR length(NEW.nombres) > 120
      OR trim(NEW.apellido_paterno) = '' OR length(NEW.apellido_paterno) > 100
      OR trim(NEW.apellido_materno) = '' OR length(NEW.apellido_materno) > 100
      OR length(COALESCE(NEW.cargo, '')) > 150
      OR length(COALESCE(NEW.correo_personal, '')) > 254
      OR length(COALESCE(NEW.correo_institucional, '')) > 254
      OR length(COALESCE(NEW.tipo_contrato, '')) > 100
      OR NEW.activo NOT IN (0, 1)
      OR (NEW.telefono IS NOT NULL AND NEW.telefono != ''
          AND (length(NEW.telefono) > 15 OR NEW.telefono GLOB '*[^0-9]*'))
      OR (NEW.celular_institucional IS NOT NULL AND NEW.celular_institucional != ''
          AND (length(NEW.celular_institucional) > 15 OR NEW.celular_institucional GLOB '*[^0-9]*'))
      OR (NEW.fecha_nacimiento IS NOT NULL AND NEW.fecha_nacimiento != ''
          AND (length(NEW.fecha_nacimiento) != 10
               OR NEW.fecha_nacimiento NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
               OR date(NEW.fecha_nacimiento, '+0 days') != NEW.fecha_nacimiento))
    BEGIN
        SELECT RAISE(ABORT, 'DNI, teléfono, celular o fecha de nacimiento inválidos');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_trabajadores_update
    BEFORE UPDATE OF dni, telefono, celular_institucional, fecha_nacimiento ON trabajadores
    WHEN (NEW.dni IS NOT OLD.dni
          AND (NEW.dni IS NULL OR length(NEW.dni) != 8 OR NEW.dni GLOB '*[^0-9]*'))
      OR (NEW.telefono IS NOT OLD.telefono AND NEW.telefono IS NOT NULL AND NEW.telefono != ''
          AND (length(NEW.telefono) > 15 OR NEW.telefono GLOB '*[^0-9]*'))
      OR (NEW.celular_institucional IS NOT OLD.celular_institucional
          AND NEW.celular_institucional IS NOT NULL AND NEW.celular_institucional != ''
          AND (length(NEW.celular_institucional) > 15 OR NEW.celular_institucional GLOB '*[^0-9]*'))
      OR (NEW.fecha_nacimiento IS NOT OLD.fecha_nacimiento
          AND NEW.fecha_nacimiento IS NOT NULL AND NEW.fecha_nacimiento != ''
          AND (length(NEW.fecha_nacimiento) != 10
               OR NEW.fecha_nacimiento NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
               OR date(NEW.fecha_nacimiento, '+0 days') != NEW.fecha_nacimiento))
    BEGIN
        SELECT RAISE(ABORT, 'DNI, teléfono, celular o fecha de nacimiento inválidos');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_trabajadores_texto_update
    BEFORE UPDATE OF nombres, apellido_paterno, apellido_materno, cargo,
                     correo_personal, correo_institucional, tipo_contrato, activo ON trabajadores
    WHEN (NEW.nombres IS NOT OLD.nombres
          AND (trim(NEW.nombres) = '' OR length(NEW.nombres) > 120))
      OR (NEW.apellido_paterno IS NOT OLD.apellido_paterno
          AND (trim(NEW.apellido_paterno) = '' OR length(NEW.apellido_paterno) > 100))
      OR (NEW.apellido_materno IS NOT OLD.apellido_materno
          AND (trim(NEW.apellido_materno) = '' OR length(NEW.apellido_materno) > 100))
      OR (NEW.cargo IS NOT OLD.cargo AND length(COALESCE(NEW.cargo, '')) > 150)
      OR (NEW.correo_personal IS NOT OLD.correo_personal
          AND length(COALESCE(NEW.correo_personal, '')) > 254)
      OR (NEW.correo_institucional IS NOT OLD.correo_institucional
          AND length(COALESCE(NEW.correo_institucional, '')) > 254)
      OR (NEW.tipo_contrato IS NOT OLD.tipo_contrato
          AND length(COALESCE(NEW.tipo_contrato, '')) > 100)
      OR (NEW.activo IS NOT OLD.activo AND NEW.activo NOT IN (0, 1))
    BEGIN
        SELECT RAISE(ABORT, 'Los campos del trabajador exceden los límites permitidos');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_equipos_celular_insert
    BEFORE INSERT ON equipos
    WHEN (NEW.categoria NOT IN ('MOVILES', 'TOKENS', 'INFORMATICOS', 'MOBILIARIO'))
      OR (NEW.celular_institucional IS NOT NULL AND NEW.celular_institucional != ''
          AND (length(NEW.celular_institucional) > 15 OR NEW.celular_institucional GLOB '*[^0-9]*'))
    BEGIN
        SELECT RAISE(ABORT, 'El celular debe contener solo números');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_equipos_celular_update
    BEFORE UPDATE OF categoria, celular_institucional ON equipos
    WHEN (NEW.categoria IS NOT OLD.categoria
          AND NEW.categoria NOT IN ('MOVILES', 'TOKENS', 'INFORMATICOS', 'MOBILIARIO'))
      OR (NEW.celular_institucional IS NOT OLD.celular_institucional
          AND NEW.celular_institucional IS NOT NULL AND NEW.celular_institucional != ''
          AND (length(NEW.celular_institucional) > 15 OR NEW.celular_institucional GLOB '*[^0-9]*'))
    BEGIN
        SELECT RAISE(ABORT, 'El celular debe contener solo números');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_certificados_insert
    BEFORE INSERT ON certificados_digitales
    WHEN NEW.dni IS NULL OR length(NEW.dni) < 1 OR length(NEW.dni) > 30
      OR NEW.dni GLOB '*[^0-9]*'
      OR trim(NEW.apellido_paterno) = '' OR length(NEW.apellido_paterno) > 100
      OR trim(NEW.apellido_materno) = '' OR length(NEW.apellido_materno) > 100
      OR trim(NEW.nombres) = '' OR length(NEW.nombres) > 120
      OR length(COALESCE(NEW.cargo, '')) > 150
      OR length(COALESCE(NEW.oficina, '')) > 150
      OR length(COALESCE(NEW.correo, '')) > 254
      OR trim(NEW.tipo_certificado) = '' OR length(NEW.tipo_certificado) > 100
      OR NEW.fecha_inicio IS NULL OR length(NEW.fecha_inicio) != 10
      OR NEW.fecha_inicio NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
      OR date(NEW.fecha_inicio, '+0 days') != NEW.fecha_inicio
      OR NEW.fecha_caducidad IS NULL OR length(NEW.fecha_caducidad) != 10
      OR NEW.fecha_caducidad NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
      OR date(NEW.fecha_caducidad, '+0 days') != NEW.fecha_caducidad
      OR NEW.fecha_caducidad < NEW.fecha_inicio
    BEGIN
        SELECT RAISE(ABORT, 'DNI o vigencia del certificado inválidos');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_certificados_update
    BEFORE UPDATE OF dni, apellido_paterno, apellido_materno, nombres, cargo, oficina,
                     correo, tipo_certificado, fecha_inicio, fecha_caducidad
    ON certificados_digitales
    WHEN (NEW.dni IS NOT OLD.dni
          AND (NEW.dni IS NULL OR length(NEW.dni) < 1 OR length(NEW.dni) > 30
               OR NEW.dni GLOB '*[^0-9]*'))
      OR (NEW.fecha_inicio IS NOT OLD.fecha_inicio
          AND (NEW.fecha_inicio IS NULL OR length(NEW.fecha_inicio) != 10
               OR NEW.fecha_inicio NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
               OR date(NEW.fecha_inicio, '+0 days') != NEW.fecha_inicio))
      OR (NEW.fecha_caducidad IS NOT OLD.fecha_caducidad
          AND (NEW.fecha_caducidad IS NULL OR length(NEW.fecha_caducidad) != 10
               OR NEW.fecha_caducidad NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
               OR date(NEW.fecha_caducidad, '+0 days') != NEW.fecha_caducidad))
      OR (NEW.apellido_paterno IS NOT OLD.apellido_paterno
          AND (trim(NEW.apellido_paterno) = '' OR length(NEW.apellido_paterno) > 100))
      OR (NEW.apellido_materno IS NOT OLD.apellido_materno
          AND (trim(NEW.apellido_materno) = '' OR length(NEW.apellido_materno) > 100))
      OR (NEW.nombres IS NOT OLD.nombres
          AND (trim(NEW.nombres) = '' OR length(NEW.nombres) > 120))
      OR (NEW.cargo IS NOT OLD.cargo AND length(COALESCE(NEW.cargo, '')) > 150)
      OR (NEW.oficina IS NOT OLD.oficina AND length(COALESCE(NEW.oficina, '')) > 150)
      OR (NEW.correo IS NOT OLD.correo AND length(COALESCE(NEW.correo, '')) > 254)
      OR (NEW.tipo_certificado IS NOT OLD.tipo_certificado
          AND (trim(NEW.tipo_certificado) = '' OR length(NEW.tipo_certificado) > 100))
      OR ((NEW.fecha_inicio IS NOT OLD.fecha_inicio
           OR NEW.fecha_caducidad IS NOT OLD.fecha_caducidad)
          AND NEW.fecha_caducidad < NEW.fecha_inicio)
    BEGIN
        SELECT RAISE(ABORT, 'DNI o vigencia del certificado inválidos');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_certificados_tipo_insert
    BEFORE INSERT ON certificados_digitales
    WHEN NEW.tipo_certificado NOT IN ('TOKEN', 'PARA PC')
      OR length(COALESCE(NEW.numero_resolucion, '')) > 100
    BEGIN
        SELECT RAISE(ABORT, 'El tipo o número de resolución del certificado no es válido');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_certificados_tipo_update
    BEFORE UPDATE OF tipo_certificado, numero_resolucion ON certificados_digitales
    WHEN (NEW.tipo_certificado IS NOT OLD.tipo_certificado
          AND NEW.tipo_certificado NOT IN ('TOKEN', 'PARA PC'))
      OR (NEW.numero_resolucion IS NOT OLD.numero_resolucion
          AND length(COALESCE(NEW.numero_resolucion, '')) > 100)
    BEGIN
        SELECT RAISE(ABORT, 'El tipo o número de resolución del certificado no es válido');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_actividades_fecha_insert
    BEFORE INSERT ON actividades_soporte
    WHEN trim(NEW.titulo) = '' OR length(NEW.titulo) > 300
      OR length(COALESCE(NEW.descripcion, '')) > 5000
      OR length(COALESCE(NEW.area_solicitante, '')) > 200
      OR length(COALESCE(NEW.solucion_aplicada, '')) > 5000
      OR NEW.categoria IS NULL OR NEW.categoria NOT IN (
          'HARDWARE', 'SOFTWARE', 'REDES', 'TOKENS', 'EQUIPOS MOVILES',
          'CAPACITACION', 'GRABAR SESION', 'TRAMITE UTI', 'SOPORTE TECNICO'
      )
      OR NEW.prioridad IS NULL OR NEW.prioridad NOT IN ('ALTA', 'MEDIA', 'BAJA')
      OR NEW.estado IS NULL OR NEW.estado NOT IN ('PENDIENTE', 'EN PROGRESO', 'COMPLETADA')
      OR (NEW.fecha_programada IS NOT NULL AND NEW.fecha_programada != ''
          AND (length(NEW.fecha_programada) != 10
               OR NEW.fecha_programada NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
               OR date(NEW.fecha_programada, '+0 days') != NEW.fecha_programada))
    BEGIN
        SELECT RAISE(ABORT, 'La fecha programada debe ser una fecha válida YYYY-MM-DD');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_actividades_fecha_update
    BEFORE UPDATE OF titulo, descripcion, area_solicitante, categoria, prioridad,
                     estado, fecha_programada, solucion_aplicada ON actividades_soporte
    WHEN (NEW.titulo IS NOT OLD.titulo
          AND (trim(NEW.titulo) = '' OR length(NEW.titulo) > 300))
      OR (NEW.descripcion IS NOT OLD.descripcion
          AND length(COALESCE(NEW.descripcion, '')) > 5000)
      OR (NEW.area_solicitante IS NOT OLD.area_solicitante
          AND length(COALESCE(NEW.area_solicitante, '')) > 200)
      OR (NEW.solucion_aplicada IS NOT OLD.solucion_aplicada
          AND length(COALESCE(NEW.solucion_aplicada, '')) > 5000)
      OR (NEW.categoria IS NOT OLD.categoria
          AND (NEW.categoria IS NULL OR NEW.categoria NOT IN (
              'HARDWARE', 'SOFTWARE', 'REDES', 'TOKENS', 'EQUIPOS MOVILES',
              'CAPACITACION', 'GRABAR SESION', 'TRAMITE UTI', 'SOPORTE TECNICO'
          )))
      OR (NEW.prioridad IS NOT OLD.prioridad
          AND (NEW.prioridad IS NULL OR NEW.prioridad NOT IN ('ALTA', 'MEDIA', 'BAJA')))
      OR (NEW.estado IS NOT OLD.estado
          AND (NEW.estado IS NULL OR NEW.estado NOT IN ('PENDIENTE', 'EN PROGRESO', 'COMPLETADA')))
      OR (NEW.fecha_programada IS NOT OLD.fecha_programada
          AND NEW.fecha_programada IS NOT NULL AND NEW.fecha_programada != ''
          AND (length(NEW.fecha_programada) != 10
               OR NEW.fecha_programada NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
               OR date(NEW.fecha_programada, '+0 days') != NEW.fecha_programada))
    BEGIN
        SELECT RAISE(ABORT, 'La fecha programada debe ser una fecha válida YYYY-MM-DD');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_actividades_hora_insert
    BEFORE INSERT ON actividades_soporte
    WHEN (NEW.hora_inicio IS NOT NULL AND NEW.hora_inicio != ''
          AND (length(NEW.hora_inicio) != 5
               OR NEW.hora_inicio NOT GLOB '[0-2][0-9]:[0-5][0-9]'
               OR substr(NEW.hora_inicio, 1, 2) > '23'))
      OR (NEW.hora_fin IS NOT NULL AND NEW.hora_fin != ''
          AND (length(NEW.hora_fin) != 5
               OR NEW.hora_fin NOT GLOB '[0-2][0-9]:[0-5][0-9]'
               OR substr(NEW.hora_fin, 1, 2) > '23'))
    BEGIN
        SELECT RAISE(ABORT, 'La hora debe tener un formato HH:MM válido');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_actividades_hora_update
    BEFORE UPDATE OF hora_inicio, hora_fin ON actividades_soporte
    WHEN (NEW.hora_inicio IS NOT OLD.hora_inicio
          AND NEW.hora_inicio IS NOT NULL AND NEW.hora_inicio != ''
          AND (length(NEW.hora_inicio) != 5
               OR NEW.hora_inicio NOT GLOB '[0-2][0-9]:[0-5][0-9]'
               OR substr(NEW.hora_inicio, 1, 2) > '23'))
      OR (NEW.hora_fin IS NOT OLD.hora_fin
          AND NEW.hora_fin IS NOT NULL AND NEW.hora_fin != ''
          AND (length(NEW.hora_fin) != 5
               OR NEW.hora_fin NOT GLOB '[0-2][0-9]:[0-5][0-9]'
               OR substr(NEW.hora_fin, 1, 2) > '23'))
    BEGIN
        SELECT RAISE(ABORT, 'La hora debe tener un formato HH:MM válido');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_documentos_fecha_insert
    BEFORE INSERT ON documentos_registrados
    WHEN NEW.fecha_emision IS NOT NULL AND NEW.fecha_emision != ''
      AND (length(NEW.fecha_emision) != 10
           OR NEW.fecha_emision NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
           OR date(NEW.fecha_emision, '+0 days') != NEW.fecha_emision)
    BEGIN
        SELECT RAISE(ABORT, 'La fecha de emisión debe ser una fecha válida YYYY-MM-DD');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_documentos_fecha_update
    BEFORE UPDATE OF fecha_emision ON documentos_registrados
    WHEN NEW.fecha_emision IS NOT OLD.fecha_emision
      AND NEW.fecha_emision IS NOT NULL AND NEW.fecha_emision != ''
      AND (length(NEW.fecha_emision) != 10
           OR NEW.fecha_emision NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
           OR date(NEW.fecha_emision, '+0 days') != NEW.fecha_emision)
    BEGIN
        SELECT RAISE(ABORT, 'La fecha de emisión debe ser una fecha válida YYYY-MM-DD');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_documentos_ordinal_insert
    BEFORE INSERT ON documentos_registrados
    WHEN NEW.numero_ordinal IS NULL
      OR typeof(NEW.numero_ordinal) != 'integer' OR NEW.numero_ordinal < 0
    BEGIN
        SELECT RAISE(ABORT, 'El correlativo del documento debe ser un entero no negativo');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_documentos_ordinal_update
    BEFORE UPDATE OF numero_ordinal ON documentos_registrados
    WHEN NEW.numero_ordinal IS NOT OLD.numero_ordinal
      AND (NEW.numero_ordinal IS NULL
           OR typeof(NEW.numero_ordinal) != 'integer' OR NEW.numero_ordinal < 0)
    BEGIN
        SELECT RAISE(ABORT, 'El correlativo del documento debe ser un entero no negativo');
    END;
    """)


def ensure_biblioteca_validation_schema(conn):
    conn.executescript("""
    CREATE TRIGGER IF NOT EXISTS validar_biblioteca_fecha_insert
    BEFORE INSERT ON biblioteca_pae
    WHEN NEW.fecha_publicacion IS NOT NULL AND NEW.fecha_publicacion != ''
      AND (length(NEW.fecha_publicacion) != 10
           OR NEW.fecha_publicacion NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
           OR date(NEW.fecha_publicacion, '+0 days') != NEW.fecha_publicacion)
    BEGIN
        SELECT RAISE(ABORT, 'La fecha de publicación debe ser una fecha válida YYYY-MM-DD');
    END;

    CREATE TRIGGER IF NOT EXISTS validar_biblioteca_fecha_update
    BEFORE UPDATE OF fecha_publicacion ON biblioteca_pae
    WHEN NEW.fecha_publicacion IS NOT OLD.fecha_publicacion
      AND NEW.fecha_publicacion IS NOT NULL AND NEW.fecha_publicacion != ''
      AND (length(NEW.fecha_publicacion) != 10
           OR NEW.fecha_publicacion NOT GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
           OR date(NEW.fecha_publicacion, '+0 days') != NEW.fecha_publicacion)
    BEGIN
        SELECT RAISE(ABORT, 'La fecha de publicación debe ser una fecha válida YYYY-MM-DD');
    END;
    """)


# ════════════════════════════════════════════════════════════
# FUNCIONES DE CONEXIÓN A BASE DE DATOS
# ════════════════════════════════════════════════════════════

def get_db():
    """Abre y retorna una conexión a la base de datos PRINCIPAL (sistema_prestamos.db).
    
    Configura SQLite con optimizaciones de rendimiento:
      - foreign_keys = ON: Activa integridad referencial (las FK se respetan)
      - journal_mode = WAL: Permite lecturas simultáneas con escrituras (Write-Ahead Logging)
      - busy_timeout = 15000: Espera hasta 15 segundos si la BD está bloqueada
      - synchronous = NORMAL: Balance entre velocidad y seguridad de datos
      - cache_size = -32000: Usa 32 MB de caché en memoria para consultas rápidas
    
    IMPORTANTE: Cada llamada crea una NUEVA conexión. Debe cerrarse con conn.close()
    después de usarla, o usar un bloque try/finally.
    
    Returns:
        sqlite3.Connection configurada con row_factory=sqlite3.Row
        (permite acceder a columnas por nombre: row["campo"])
    """
    conn = sqlite3.connect("sistema_prestamos.db", timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA busy_timeout = 15000;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    conn.execute("PRAGMA cache_size = -32000;")
    return conn

def get_biblioteca_db():
    """Abre y retorna una conexión a la base de datos de BIBLIOTECA PAE (biblioteca_pae.db).
    
    Base de datos separada que almacena enlaces a eventos y capacitaciones
    del Programa de Alimentación Escolar. Se mantiene separada de la BD
    principal para modularidad.
    
    Returns:
        sqlite3.Connection con row_factory configurada.
    """
    conn = sqlite3.connect("biblioteca_pae.db", timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=15000;")
    return conn


def apply_input_validation_migrations():
    """Installs validation triggers on already-initialized databases only."""
    conn = get_db()
    try:
        tablas = {
            fila[0] for fila in conn.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
        requeridas = {
            "trabajadores", "equipos", "certificados_digitales",
            "actividades_soporte", "documentos_registrados",
        }
        faltantes = sorted(requeridas - tablas)
        if faltantes:
            raise RuntimeError(
                "No se aplicaron las restricciones: falta inicializar el esquema "
                f"de la base de datos ({', '.join(faltantes)})."
            )
        columnas_certificados = {
            fila[1] for fila in conn.execute(
                "PRAGMA table_info(certificados_digitales)"
            ).fetchall()
        }
        if "trabajador_id" not in columnas_certificados:
            conn.execute(
                "ALTER TABLE certificados_digitales ADD COLUMN trabajador_id INTEGER"
            )
        if "numero_resolucion" not in columnas_certificados:
            conn.execute(
                "ALTER TABLE certificados_digitales ADD COLUMN numero_resolucion TEXT"
            )
        ensure_input_validation_schema(conn)
        conn.commit()
    finally:
        conn.close()

    biblioteca_conn = get_biblioteca_db()
    try:
        existe_biblioteca = biblioteca_conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'biblioteca_pae'"
        ).fetchone()
        if not existe_biblioteca:
            raise RuntimeError(
                "No se aplicaron las restricciones: falta inicializar la tabla biblioteca_pae."
            )
        ensure_biblioteca_validation_schema(biblioteca_conn)
        biblioteca_conn.commit()
    finally:
        biblioteca_conn.close()


# ════════════════════════════════════════════════════════════
# ESQUEMA DE GRABACIONES APOWERREC
# ════════════════════════════════════════════════════════════

def ensure_apowerrec_schema(conn):
    """Crea la tabla de grabaciones ApowerREC si no existe y migra columnas.
    
    ApowerREC es un software de grabación de pantalla. Este módulo permite
    sincronizar automáticamente las grabaciones de sesiones de soporte TI
    con el sistema, subiéndolas a Google Drive y vinculándolas con la
    actividad correspondiente.
    
    La tabla 'grabaciones_apowerrec' almacena:
      - sha256: Hash único del archivo de video (evita duplicados)
      - archivo_nombre: Nombre del archivo MP4
      - actividad_id: FK a la actividad de soporte que se grabó
      - fecha_grabacion: Cuándo se realizó la grabación
      - estado: PROCESANDO → COMPLETADO o ERROR
      - drive_file_id/drive_url: Referencia al archivo en Google Drive
      - biblioteca_id: FK a biblioteca_pae si se publicó como evento
    """
    # Agregar columna estado_iniciado_el a actividades_soporte si no existe
    # (para rastrear cuándo una actividad pasó a "EN PROGRESO")
    columnas_actividades = {
        fila[1] for fila in conn.execute("PRAGMA table_info(actividades_soporte)").fetchall()
    }
    if "estado_iniciado_el" not in columnas_actividades:
        conn.execute("ALTER TABLE actividades_soporte ADD COLUMN estado_iniciado_el REAL")
    
    # Inicializar timestamp para actividades que ya estaban EN PROGRESO
    conn.execute(
        "UPDATE actividades_soporte SET estado_iniciado_el = strftime('%s', 'now') "
        "WHERE estado = 'EN PROGRESO' AND estado_iniciado_el IS NULL"
    )
    
    conn.execute("""
    CREATE TABLE IF NOT EXISTS grabaciones_apowerrec (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sha256 TEXT NOT NULL UNIQUE,          -- Hash SHA-256 del video (clave de deduplicación)
        archivo_nombre TEXT NOT NULL,          -- Nombre original del archivo MP4
        actividad_id INTEGER NOT NULL,         -- Actividad de soporte vinculada
        fecha_grabacion TEXT NOT NULL,          -- Fecha/hora de la grabación
        estado TEXT NOT NULL DEFAULT 'PROCESANDO',  -- PROCESANDO, COMPLETADO, ERROR
        drive_file_id TEXT,                    -- ID del archivo en Google Drive
        drive_url TEXT,                        -- URL pública de Google Drive
        biblioteca_id INTEGER,                 -- FK a biblioteca_pae (si se publicó)
        error TEXT,                            -- Mensaje de error (si falló)
        actualizado_el REAL NOT NULL,          -- Timestamp Unix de última actualización
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.commit()


# ════════════════════════════════════════════════════════════
# ESQUEMA DE MATRIZ DE ACCESOS (CARGOS × SISTEMAS)
# ════════════════════════════════════════════════════════════

def ensure_access_matrix_schema(conn):
    """Crea las tablas del módulo de Cargos y Asignaciones / Matriz de Accesos.
    
    Este módulo gestiona QUÉ SISTEMAS del PAE debe acceder CADA CARGO.
    Por ejemplo: un "Supervisor de Planta" necesita acceso a SIGO (módulo
    Vigilancia y Monitoreo) con nivel REGISTRO, pero un "Monitor de Gestión
    Local" necesita acceso a SIGO (módulo Gestión de Actores) con nivel LECTURA.
    
    Tablas creadas:
      - sistemas_modulos_ut: Catálogo jerárquico de sistemas
        (SISTEMA → MÓDULO → SUBMÓDULO, ej: SIGO → Prestación de Servicio → Programación de Raciones)
      - perfiles_cargo_ut: Un perfil por cada cargo diferente
        (con funciones, autorizador, acción de acceso: NUEVO/MANTENER/MODIFICAR/RETIRAR)
      - accesos_perfil_cargo_ut: Tabla pivote cargo×sistema con nivel de acceso
        (CUENTA, LECTURA, REGISTRO, EDICION, ADMINISTRACION)
      - historial_perfiles_cargo_ut: Auditoría de cambios en perfiles
      - solicitudes_acceso_trabajador_ut: Solicitudes individuales de acceso
      - modulos_solicitud_acceso_trabajador_ut: Módulos solicitados en cada solicitud
    
    También precarga el catálogo completo de sistemas del PAE:
      - SGD, Correo electrónico, Mesa de Ayuda (obligatorios para todos)
      - SIGO con todos sus módulos y submódulos
      - SIGDEL, SADE, SIGA, SIAF CONSULTA
      - Aplicativos móviles del PAE
    """
    cursor = conn.cursor()
    
    # ── Tabla: Catálogo de Sistemas/Módulos del PAE ──
    # Estructura jerárquica: SISTEMA > MÓDULO > SUBMÓDULO
    # El campo padre_id permite la jerarquía (self-referencing FK)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sistemas_modulos_ut (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,  -- Nombre del sistema/módulo
        tipo TEXT NOT NULL CHECK(tipo IN ('SISTEMA', 'MODULO', 'SUBMODULO', 'APK', 'SERVICIO')),
        padre_id INTEGER,                             -- FK al sistema padre (NULL = raíz)
        obligatorio_todos INTEGER NOT NULL DEFAULT 0 CHECK(obligatorio_todos IN (0, 1)),
        activo INTEGER NOT NULL DEFAULT 1 CHECK(activo IN (0, 1)),
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(padre_id) REFERENCES sistemas_modulos_ut(id) ON DELETE RESTRICT
    )
    """)
    
    # ── Tabla: Perfiles de Acceso por Cargo ──
    # Cada cargo único de la tabla trabajadores tiene un perfil aquí
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS perfiles_cargo_ut (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cargo TEXT NOT NULL COLLATE NOCASE UNIQUE,  -- Nombre del cargo (ej: "Monitor de Gestión Local")
        funciones TEXT NOT NULL DEFAULT '',           -- Funciones principales del cargo
        autorizador TEXT NOT NULL DEFAULT '',         -- Quién autoriza el acceso
        reportes_requeridos TEXT NOT NULL DEFAULT '', -- Reportes que necesita generar
        accion_acceso TEXT NOT NULL DEFAULT 'NUEVO'   -- MANTENER, MODIFICAR, NUEVO o RETIRAR
            CHECK(accion_acceso IN ('MANTENER', 'MODIFICAR', 'NUEVO', 'RETIRAR')),
        validado INTEGER NOT NULL DEFAULT 0 CHECK(validado IN (0, 1)),  -- Si fue revisado
        activo INTEGER NOT NULL DEFAULT 1 CHECK(activo IN (0, 1)),
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        actualizado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # Migraciones: agregar columnas nuevas si no existen
    columnas_perfiles = {
        fila[1] for fila in cursor.execute(
            "PRAGMA table_info(perfiles_cargo_ut)"
        ).fetchall()
    }
    if "reportes_requeridos" not in columnas_perfiles:
        cursor.execute("""
            ALTER TABLE perfiles_cargo_ut
            ADD COLUMN reportes_requeridos TEXT NOT NULL DEFAULT ''
        """)
    if "accion_acceso" not in columnas_perfiles:
        cursor.execute("""
            ALTER TABLE perfiles_cargo_ut
            ADD COLUMN accion_acceso TEXT NOT NULL DEFAULT 'NUEVO'
        """)
    
    # ── Tabla Pivote: Accesos de cada Perfil/Cargo a cada Sistema/Módulo ──
    # Esta es la MATRIZ DE ACCESOS propiamente dicha
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS accesos_perfil_cargo_ut (
        perfil_id INTEGER NOT NULL,
        sistema_modulo_id INTEGER NOT NULL,
        nivel_acceso TEXT NOT NULL CHECK(
            nivel_acceso IN ('CUENTA', 'LECTURA', 'REGISTRO', 'EDICION', 'ADMINISTRACION')
        ),
        PRIMARY KEY(perfil_id, sistema_modulo_id),
        FOREIGN KEY(perfil_id) REFERENCES perfiles_cargo_ut(id) ON DELETE CASCADE,
        FOREIGN KEY(sistema_modulo_id) REFERENCES sistemas_modulos_ut(id) ON DELETE RESTRICT
    )
    """)
    
    # ── Tabla: Historial/Auditoría de cambios en perfiles ──
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS historial_perfiles_cargo_ut (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        perfil_id INTEGER,
        cargo TEXT NOT NULL,
        cambios TEXT NOT NULL,           -- Descripción de qué cambió
        justificacion TEXT NOT NULL,     -- Por qué se realizó el cambio
        autorizado_por TEXT NOT NULL,    -- Quién autorizó
        usuario_id INTEGER,             -- Usuario del sistema que lo registró
        registrado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(perfil_id) REFERENCES perfiles_cargo_ut(id) ON DELETE SET NULL,
        FOREIGN KEY(usuario_id) REFERENCES usuarios_sistema(id) ON DELETE SET NULL
    )
    """)
    
    # ── Tabla: Solicitudes individuales de acceso por trabajador ──
    # Cuando un trabajador específico necesita acceso diferente al de su cargo
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS solicitudes_acceso_trabajador_ut (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        trabajador_id INTEGER NOT NULL,
        funciones TEXT NOT NULL,
        nivel_acceso TEXT NOT NULL CHECK(
            nivel_acceso IN ('CONSULTA', 'REGISTRO', 'MODIFICACION', 'APROBACION', 'REPORTES')
        ),
        reportes_requeridos TEXT NOT NULL,
        accion TEXT NOT NULL CHECK(accion IN ('MANTENER', 'MODIFICAR', 'NUEVO', 'RETIRAR')),
        autorizador TEXT NOT NULL,
        estado TEXT NOT NULL DEFAULT 'PENDIENTE' CHECK(estado IN ('PENDIENTE', 'VALIDADA')),
        usuario_id INTEGER,
        registrado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(trabajador_id) REFERENCES trabajadores(id) ON DELETE CASCADE,
        FOREIGN KEY(usuario_id) REFERENCES usuarios_sistema(id) ON DELETE SET NULL
    )
    """)
    
    # ── Tabla: Módulos solicitados en cada solicitud individual ──
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS modulos_solicitud_acceso_trabajador_ut (
        solicitud_id INTEGER NOT NULL,
        sistema_modulo_id INTEGER NOT NULL,
        nivel_acceso TEXT NOT NULL DEFAULT 'CONSULTA' CHECK(
            nivel_acceso IN (
                'CUENTA', 'LECTURA', 'REGISTRO', 'EDICION', 'ADMINISTRACION',
                'CONSULTA', 'MODIFICACION', 'APROBACION', 'REPORTES'
            )
        ),
        PRIMARY KEY(solicitud_id, sistema_modulo_id),
        FOREIGN KEY(solicitud_id) REFERENCES solicitudes_acceso_trabajador_ut(id) ON DELETE CASCADE,
        FOREIGN KEY(sistema_modulo_id) REFERENCES sistemas_modulos_ut(id) ON DELETE RESTRICT
    )
    """)
    
    # Migración: agregar nivel_acceso si no existe
    columnas_modulos_solicitud = {
        fila[1] for fila in cursor.execute(
            "PRAGMA table_info(modulos_solicitud_acceso_trabajador_ut)"
        ).fetchall()
    }
    if "nivel_acceso" not in columnas_modulos_solicitud:
        cursor.execute("""
            ALTER TABLE modulos_solicitud_acceso_trabajador_ut
            ADD COLUMN nivel_acceso TEXT NOT NULL DEFAULT 'CONSULTA'
        """)
    
    # Índice para búsquedas rápidas de solicitudes por trabajador
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_solicitudes_acceso_trabajador
        ON solicitudes_acceso_trabajador_ut(trabajador_id, id)
    """)

    # ── Normalización de nombres de sistemas del catálogo ──
    cursor.execute("""
        UPDATE sistemas_modulos_ut
        SET nombre = 'SIGO'
        WHERE nombre = 'SIGO PAE' COLLATE NOCASE
          AND NOT EXISTS (
              SELECT 1 FROM sistemas_modulos_ut WHERE nombre = 'SIGO' COLLATE NOCASE
          )
    """)
    cursor.execute("""
        UPDATE sistemas_modulos_ut
        SET nombre = 'CORREO ELECTRONICO'
        WHERE nombre IN ('Correo electrónico', 'Correo electronico') COLLATE NOCASE
          AND NOT EXISTS (
              SELECT 1 FROM sistemas_modulos_ut WHERE nombre = 'CORREO ELECTRONICO' COLLATE NOCASE
          )
    """)
    cursor.execute("""
        DELETE FROM sistemas_modulos_ut
        WHERE id != 2 AND (nombre LIKE 'Correo electr%' OR nombre LIKE 'Correo electron%')
          AND NOT EXISTS (SELECT 1 FROM accesos_perfil_cargo_ut WHERE sistema_modulo_id = sistemas_modulos_ut.id)
          AND NOT EXISTS (SELECT 1 FROM modulos_solicitud_acceso_trabajador_ut WHERE sistema_modulo_id = sistemas_modulos_ut.id)
    """)
    cursor.execute("""
        DELETE FROM sistemas_modulos_ut
        WHERE id != 3 AND (nombre LIKE 'Usuario de Mesa%' OR nombre LIKE 'Mesa de Ayuda%')
          AND NOT EXISTS (SELECT 1 FROM accesos_perfil_cargo_ut WHERE sistema_modulo_id = sistemas_modulos_ut.id)
          AND NOT EXISTS (SELECT 1 FROM modulos_solicitud_acceso_trabajador_ut WHERE sistema_modulo_id = sistemas_modulos_ut.id)
    """)

    # ── Asegurar el sistema contenedor de aplicaciones móviles: "APK - APLICATIVOS MOBILES" ──
    # Si existía previamente como "APK - APLICATIVOS MOVILES", renombrarlo a "APK - APLICATIVOS MOBILES"
    cursor.execute("""
        UPDATE sistemas_modulos_ut
        SET nombre = 'APK - APLICATIVOS MOBILES'
        WHERE nombre = 'APK - APLICATIVOS MOVILES' COLLATE NOCASE
    """)
    cursor.execute("""
        INSERT INTO sistemas_modulos_ut (nombre, tipo, obligatorio_todos, activo)
        VALUES ('APK - APLICATIVOS MOBILES', 'SISTEMA', 0, 1)
        ON CONFLICT(nombre) DO UPDATE SET activo = 1
    """)
    apk_parent = cursor.execute("""
        SELECT id FROM sistemas_modulos_ut WHERE nombre = 'APK - APLICATIVOS MOBILES' COLLATE NOCASE
    """).fetchone()
    apk_parent_id = apk_parent["id"] if apk_parent else None

    # Vincular los 4 aplicativos móviles bajo "APK - APLICATIVOS MOBILES"
    apks_oficiales = (
        "APLICATIVO MOVIL: PAE IIEE",
        "APLICATIVO MOVIL: PAE ESTABLECIMIENTO",
        "APLICATIVO MOVIL: ALMACEN EN IE",
        "APLICATIVO MOVIL: IIEE",
    )
    for apk_nombre in apks_oficiales:
        cursor.execute("""
            INSERT INTO sistemas_modulos_ut (nombre, tipo, padre_id, obligatorio_todos, activo)
            VALUES (?, 'APK', ?, 0, 1)
            ON CONFLICT(nombre) DO UPDATE SET
                tipo = 'APK',
                padre_id = excluded.padre_id,
                activo = 1
        """, (apk_nombre, apk_parent_id))

    # Asegurar incondicionalmente que todos los aplicativos móviles tengan como padre el contenedor APK
    if apk_parent_id:
        cursor.execute("""
            UPDATE sistemas_modulos_ut
            SET padre_id = ?, tipo = 'APK', activo = 1
            WHERE (nombre LIKE 'APLICATIVO MOVIL:%' OR tipo = 'APK')
              AND id != ?
        """, (apk_parent_id, apk_parent_id))

    # Eliminar contenedores obsoletos si existen
    cursor.execute("""
        DELETE FROM sistemas_modulos_ut
        WHERE nombre = 'Aplicativos móviles del PAE' COLLATE NOCASE
          AND NOT EXISTS (
              SELECT 1 FROM sistemas_modulos_ut s2 WHERE s2.padre_id = sistemas_modulos_ut.id
          )
    """)

    # Desactivar registros obsoletos o temporales
    cursor.execute("""
        UPDATE sistemas_modulos_ut
        SET activo = 0
        WHERE nombre IN ('Sistemas de reportes', 'Otros sistemas/aplicativos', 'SADCIL') COLLATE NOCASE
          AND NOT EXISTS (
              SELECT 1 FROM sistemas_modulos_ut s2 WHERE s2.padre_id = sistemas_modulos_ut.id
          )
    """)

    # ── Precarga de sistemas oficiales principales de la UTI ──
    sistemas_oficiales = (
        ("SIGO", "SISTEMA", 0),
        ("SIGDEL", "SISTEMA", 0),
        ("SGD", "SISTEMA", 1),
        ("SADE", "SISTEMA", 0),
        ("CORREO ELECTRONICO", "SERVICIO", 1),
        ("MESA DE AYUDA", "SERVICIO", 1),
        ("APK - APLICATIVOS MOBILES", "SISTEMA", 0),
        ("SIGA", "SISTEMA", 0),
        ("SIAF CONSULTA", "SISTEMA", 0),
    )
    for nombre, tipo, obligatorio in sistemas_oficiales:
        cursor.execute("""
            INSERT INTO sistemas_modulos_ut (nombre, tipo, obligatorio_todos, activo)
            VALUES (?, ?, ?, 1)
            ON CONFLICT(nombre) DO UPDATE SET
                obligatorio_todos = excluded.obligatorio_todos,
                activo = 1
        """, (nombre, tipo, obligatorio))

    # ── Precarga de Módulos y Submódulos del SIGO ──
    # SIGO es el sistema más grande del PAE, con muchos módulos internos.
    # Se organizan jerárquicamente: SIGO > Módulo > Submódulo
    modulos_sigo = {
        "Prestación de Servicio": (
            "Grupo y Productos", "Combinación raciones", "Programación de Raciones",
            "Programación de Colegios", "Volúmenes de Compra", "Presentaciones Reales",
            "Volumen Evaluación",
        ),
        "Vigilancia y Monitoreo": (
            "Lib. Establecimientos", "Prog. EIA", "ETE (Sup. Inicial)", "IIEE",
            "Prog. AGL", "Vig. IIEE", "Actualización IIEE", "Administración IIEE",
            "Incidencias", "Asistencia Técnica", "Vigilancia Social", "Sectorización",
            "Atención alimentaria", "Sectorización SyL",
        ),
        "TRANSFERENCIA MONETARIA": (
            "Contratistas", "Convocatorias", "Contratos",
            "Acta Contratista - Comida Caliente",
            "Acta Contratista - Alimentos Listos Para El Consumo",
            "Valorización", "Solicitud de Transferencia", "Liquidación",
            "Rendición de Cuentas", "Datos de Vehículo", "Consulta Postores",
            "Reportes UGCTR", "DATOS ADENDA IIEE", "VOC",
        ),
        "Gestión de Actores": (
            "Perfiles y Usuarios", "CAE", "Comités de Contrataciones",
            "Capacitaciones", "Asist. Técnica", "Familia", "CGAE",
        ),
        "Tesorería": (
            "Elaboración de Cartas Órdenes", "Elaboración de Recibos de Ingreso",
            "Elaboración de Devolución de Garantía", "Elaboración de Carta Fianza",
            "Devolución de Garantía de Seriedad",
        ),
        "Informáticos": (
            "IMEIs", "Asig. Equipos Celulares", "URNA DE CRISTAL VIDEOS",
        ),
        "UDPC": (
            "Asignación de IIEE a ítems", "Gestión de Ítems", "Mapas",
            "METAS FÍSICAS Y FINANCIERAS", "FORM. DE ESTRUCTURA DE COSTOS",
            "Gestión Usuarios",
        ),
        "Administrativos": ("SADE (módulo SIGO)",),
        "Proceso de Contratación": (
            "Trámite Actas", "Administración Plantilla Acta",
            "Asignación Comité de Compra", "Carga Valor Referencial",
            "Gestión de Pestañas y Enlaces", "Eval. Participante", "D.J Interés",
            "Reportes", "Expresión de intereses", "Expresión de intereses UT",
            "Evaluación Expresión de Intereses",
        ),
    }
    
    # Insertar módulos y submódulos del SIGO en la jerarquía
    sigo = cursor.execute(
        "SELECT id FROM sistemas_modulos_ut WHERE nombre = 'SIGO' COLLATE NOCASE"
    ).fetchone()
    if sigo:
        for nombre_modulo, submodulos in modulos_sigo.items():
            cursor.execute("""
                INSERT INTO sistemas_modulos_ut (nombre, tipo, padre_id)
                VALUES (?, 'MODULO', ?)
                ON CONFLICT(nombre) DO NOTHING
            """, (nombre_modulo, sigo["id"]))
            modulo = cursor.execute(
                "SELECT id FROM sistemas_modulos_ut WHERE nombre = ? COLLATE NOCASE",
                (nombre_modulo,),
            ).fetchone()
            if modulo:
                cursor.executemany("""
                    INSERT INTO sistemas_modulos_ut (nombre, tipo, padre_id)
                    VALUES (?, 'SUBMODULO', ?)
                    ON CONFLICT(nombre) DO NOTHING
                """, [(nombre, modulo["id"]) for nombre in submodulos])

    # Sincronizar cargos de trabajadores con perfiles_cargo_ut
    # (cada cargo único que aparezca en la tabla trabajadores genera un perfil)
    cursor.execute("""
        INSERT OR IGNORE INTO perfiles_cargo_ut (cargo)
        SELECT DISTINCT TRIM(cargo)
        FROM trabajadores
        WHERE cargo IS NOT NULL AND TRIM(cargo) != ''
    """)
    conn.commit()


def _migrar_tipo_submodulo(conn):
    """Migración de esquema: agrega 'SUBMODULO' al CHECK constraint de tipo.
    
    Esta migración se ejecuta UNA SOLA VEZ cuando se actualiza el sistema
    desde una versión anterior que no soportaba el tipo 'SUBMODULO'.
    Recrea la tabla sistemas_modulos_ut con el constraint actualizado,
    preservando todos los datos existentes.
    
    Proceso:
    1. Verifica si el constraint actual ya incluye 'SUBMODULO'
    2. Si no, desactiva FK temporalmente
    3. Crea tabla nueva con constraint correcto
    4. Copia todos los datos
    5. Elimina tabla vieja y renombra la nueva
    6. Reactiva FK
    """
    fila = conn.execute("""
        SELECT sql
        FROM sqlite_master
        WHERE type = 'table' AND name = 'sistemas_modulos_ut'
    """).fetchone()
    if not fila or "SUBMODULO" in (fila["sql"] or ""):
        return  # Ya tiene el constraint correcto, no hacer nada
    if conn.in_transaction:
        raise RuntimeError(
            "No se puede actualizar el catálogo de sistemas durante una transacción activa."
        )

    conn.execute("PRAGMA foreign_keys = OFF")
    try:
        conn.execute("BEGIN")
        conn.execute("""
            CREATE TABLE sistemas_modulos_ut_nueva (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
                tipo TEXT NOT NULL CHECK(
                    tipo IN ('SISTEMA', 'MODULO', 'SUBMODULO', 'APK', 'SERVICIO')
                ),
                padre_id INTEGER,
                obligatorio_todos INTEGER NOT NULL DEFAULT 0 CHECK(obligatorio_todos IN (0, 1)),
                activo INTEGER NOT NULL DEFAULT 1 CHECK(activo IN (0, 1)),
                creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(padre_id) REFERENCES sistemas_modulos_ut_nueva(id) ON DELETE RESTRICT
            )
        """)
        conn.execute("""
            INSERT INTO sistemas_modulos_ut_nueva
                (id, nombre, tipo, padre_id, obligatorio_todos, activo, creado_el)
            SELECT id, nombre, tipo, padre_id, obligatorio_todos, activo, creado_el
            FROM sistemas_modulos_ut
        """)
        conn.execute("DROP TABLE sistemas_modulos_ut")
        conn.execute("""
            ALTER TABLE sistemas_modulos_ut_nueva
            RENAME TO sistemas_modulos_ut
        """)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.execute("PRAGMA foreign_keys = ON")


# ════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL DE INICIALIZACIÓN DE LA BASE DE DATOS
# ════════════════════════════════════════════════════════════

def init_db():
    """Inicializa la base de datos principal: crea tablas, migra columnas, crea admin.
    
    Esta función se ejecuta cada vez que arranca el servidor (tanto en desarrollo
    como en producción con Gunicorn). Todas las operaciones usan "CREATE TABLE IF
    NOT EXISTS" y verifican columnas antes de ALTER TABLE, por lo que es seguro
    ejecutarla múltiples veces (idempotente).
    
    Orden de operaciones:
    1. Migración de esquema (si es necesaria)
    2. Crear tabla usuarios_sistema + admin por defecto
    3. Crear tabla trabajadores + migraciones de columnas
    4. Crear tablas de equipos, préstamos, actividades
    5. Crear tablas de SGD, documentos, configuraciones
    6. Crear tablas de certificados digitales
    7. Crear esquema de matriz de accesos
    8. Crear índices de rendimiento
    """
    conn = get_db()
    _migrar_tipo_submodulo(conn)
    cursor = conn.cursor()

    # ── 1. TABLA DE USUARIOS DEL SISTEMA ──
    # Almacena las credenciales y permisos de quienes acceden al sistema.
    # Las contraseñas se guardan hasheadas con pbkdf2 (generate_password_hash).
    # El campo 'permisos' es un JSON con acceso granular por módulo y acción.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios_sistema (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE NOT NULL,       -- Nombre de usuario para login
        password TEXT NOT NULL,              -- Hash pbkdf2 de la contraseña
        nombre_completo TEXT NOT NULL,       -- Nombre visible del usuario
        rol TEXT NOT NULL DEFAULT 'USUARIO', -- ADMIN o USUARIO
        permisos TEXT NOT NULL DEFAULT '{}', -- JSON: {"modulo": ["ver","modificar","eliminar","reportar"]}
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Migración: agregar columna de permisos si no existe (sistemas antiguos)
    columnas_usuarios = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(usuarios_sistema)").fetchall()
    }
    if "permisos" not in columnas_usuarios:
        cursor.execute("ALTER TABLE usuarios_sistema ADD COLUMN permisos TEXT NOT NULL DEFAULT '{}'")
    
    # Asignar permisos completos a usuarios que tengan permisos=NULL (migración)
    permisos_legacy = '{"personal":["ver","modificar","eliminar","reportar"],"prestamos":["ver","modificar","eliminar","reportar"],"actividades":["ver","modificar","eliminar","reportar"],"documentaria":["ver","modificar","eliminar","reportar"],"sgd":["ver","modificar","eliminar","reportar"],"reportes":["ver","reportar"],"biblioteca":["ver","modificar","eliminar","reportar"],"tokens":["ver","modificar","eliminar","reportar"]}'
    cursor.execute("UPDATE usuarios_sistema SET permisos = ? WHERE permisos IS NULL", (permisos_legacy,))

    # ── Crear usuario administrador por defecto si la BD está vacía ──
    # Esto permite el primer acceso al sistema recién instalado.
    # Credenciales: usuario="admin", contraseña="admin123"
    # ¡IMPORTANTE! Cambiar la contraseña después del primer acceso.
    cant_usuarios = cursor.execute("SELECT COUNT(*) FROM usuarios_sistema").fetchone()[0]
    if cant_usuarios == 0:
        admin_pass = generate_password_hash("admin123")
        permisos_admin = '{"personal":["ver","modificar","eliminar","reportar"],"prestamos":["ver","modificar","eliminar","reportar"],"actividades":["ver","modificar","eliminar","reportar"],"documentaria":["ver","modificar","eliminar","reportar"],"sgd":["ver","modificar","eliminar","reportar"],"reportes":["ver","reportar"],"biblioteca":["ver","modificar","eliminar","reportar"],"tokens":["ver","modificar","eliminar","reportar"]}'
        cursor.execute("""
            INSERT INTO usuarios_sistema (usuario, password, nombre_completo, rol, permisos)
            VALUES (?, ?, ?, ?, ?)
        """, ("admin", admin_pass, "Administrador del Sistema", "ADMIN", permisos_admin))
        conn.commit()

    # ── 2. TABLA DE TRABAJADORES ──
    # Padrón de todo el personal de la oficina. Se puede importar desde Excel.
    # Campos clave: DNI (único), nombres, cargo, tipo de contrato (CAS/OS).
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS trabajadores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        dni TEXT UNIQUE NOT NULL,              -- Documento Nacional de Identidad (8 dígitos)
        nombres TEXT NOT NULL,
        apellido_paterno TEXT NOT NULL,
        apellido_materno TEXT NOT NULL,
        cargo TEXT,                            -- Cargo en la oficina
        correo_personal TEXT,
        correo_institucional TEXT,             -- Correo @qaliwarma.gob.pe
        tipo_contrato TEXT,                    -- CAS, OS (Orden de Servicio), etc.
        telefono TEXT,
        fecha_nacimiento TEXT,
        activo INTEGER NOT NULL DEFAULT 1      -- 1=activo, 0=cesado/inactivo
    )
    """)

    # Migraciones de columnas para trabajadores
    columnas_trabajadores = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(trabajadores)").fetchall()
    }
    if "celular_institucional" not in columnas_trabajadores:
        cursor.execute("ALTER TABLE trabajadores ADD COLUMN celular_institucional TEXT")
    if "celular_institucional_normalizado" not in columnas_trabajadores:
        cursor.execute("ALTER TABLE trabajadores ADD COLUMN celular_institucional_normalizado TEXT")
    if "activo" not in columnas_trabajadores:
        cursor.execute("ALTER TABLE trabajadores ADD COLUMN activo INTEGER NOT NULL DEFAULT 1")

    # Normalizar números de celular para búsquedas consistentes
    trabajadores_con_celular = cursor.execute(
        "SELECT id, celular_institucional FROM trabajadores"
    ).fetchall()
    celulares_normalizados = {}
    for trabajador in trabajadores_con_celular:
        celular_normalizado = normalizar_numero_celular(trabajador["celular_institucional"])
        cursor.execute(
            "UPDATE trabajadores SET celular_institucional_normalizado = ? WHERE id = ?",
            (celular_normalizado, trabajador["id"]),
        )
        if celular_normalizado:
            celulares_normalizados.setdefault(celular_normalizado, []).append(trabajador["id"])

    # Crear índice único solo si no hay duplicados de celular
    if all(len(ids) == 1 for ids in celulares_normalizados.values()):
        cursor.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS idx_trabajadores_celular_unico
            ON trabajadores(celular_institucional_normalizado)
            WHERE celular_institucional_normalizado != ''
        """)

    # ── 3. TABLA DE CONTRATOS DE TRABAJADORES ──
    # Almacena los archivos PDF de contratos (Orden de Servicio u CAS)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS contratos_trabajadores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        trabajador_id INTEGER NOT NULL,
        tipo_documento TEXT NOT NULL CHECK(tipo_documento IN ('OS', 'CAS')),
        nombre_archivo TEXT NOT NULL,
        ruta_archivo TEXT NOT NULL,
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(trabajador_id) REFERENCES trabajadores(id) ON DELETE CASCADE
    )
    """)

    # ── 4. TABLA DE EQUIPOS ──
    # Inventario de todos los bienes asignados a la UTI.
    # Se clasifican por categoría: INFORMATICOS, MOVILES, TOKENS, MOBILIARIO.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS equipos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo_margesi TEXT,          -- Código patrimonial MARGESI del bien
        codigo_interno TEXT,          -- Código interno de la oficina
        categoria TEXT NOT NULL,      -- INFORMATICOS, MOVILES, TOKENS, MOBILIARIO
        tipo_equipo TEXT,             -- Laptop, PC, Celular, Token USB, etc.
        descripcion TEXT,             -- Descripción detallada del equipo
        marca TEXT,
        modelo TEXT,
        serie_medidas TEXT,           -- Número de serie o medidas físicas
        numero_serie TEXT,
        estado_conservacion TEXT,     -- Bueno, Regular, Malo
        color TEXT,
        sede TEXT,                    -- Sede/oficina donde está el equipo
        valor_inicial TEXT,           -- Valor de adquisición
        valor_neto TEXT,              -- Valor actual (depreciado)
        observaciones TEXT,
        estado TEXT DEFAULT 'EN ALMACEN'  -- EN ALMACEN, PRESTADO, BAJA
    )
    """)

    # Migración: agregar campo celular_institucional a equipos (para tokens/chips)
    columnas_equipos = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(equipos)").fetchall()
    }
    if "celular_institucional" not in columnas_equipos:
        cursor.execute("ALTER TABLE equipos ADD COLUMN celular_institucional TEXT")
    
    # ── 5. TABLA DE PRÉSTAMOS ──
    # Registra cada entrega y devolución de equipos a trabajadores.
    # Un equipo PRESTADO no puede volver a prestarse hasta que se devuelva.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS prestamos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        equipo_id INTEGER,                    -- FK al equipo prestado
        trabajador_id INTEGER,                -- FK al trabajador que recibe
        fecha_prestamo TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        fecha_devolucion TIMESTAMP,           -- NULL si aún no se devolvió
        estado TEXT DEFAULT 'ACTIVO',         -- ACTIVO o DEVUELTO
        observaciones_entrega TEXT,
        observaciones_devolucion TEXT,
        archivo_acta_entrega TEXT,            -- Ruta al archivo Word del acta
        FOREIGN KEY(equipo_id) REFERENCES equipos(id),
        FOREIGN KEY(trabajador_id) REFERENCES trabajadores(id)
    )
    """)

    # ── 6. TABLA DE ACTIVIDADES DE SOPORTE TI ──
    # Bitácora diaria de todas las tareas de soporte técnico.
    # Incluye: título, categoría (HARDWARE/SOFTWARE/REDES/etc.), prioridad,
    # estado (PENDIENTE/EN PROGRESO/COMPLETADA), horarios, solución aplicada.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS actividades_soporte (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,                  -- Título descriptivo de la actividad
        descripcion TEXT,                      -- Detalle del requerimiento
        categoria TEXT DEFAULT 'SOPORTE TECNICO',  -- HARDWARE, SOFTWARE, REDES, TOKENS, etc.
        area_solicitante TEXT,                 -- Quién solicitó el soporte
        prioridad TEXT DEFAULT 'MEDIA',        -- ALTA, MEDIA, BAJA
        estado TEXT DEFAULT 'PENDIENTE',       -- PENDIENTE, EN PROGRESO, COMPLETADA
        fecha_programada TEXT,                 -- Fecha en que se realizará
        hora_inicio TEXT,                      -- Hora de inicio (HH:MM)
        hora_fin TEXT,                         -- Hora de fin (HH:MM)
        whatsapp_telefono TEXT,                -- Número internacional para recordatorios por WhatsApp
        solucion_aplicada TEXT,                -- Descripción de la solución (al completar)
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # Migración: agregar usuario_id para saber quién creó cada actividad
    columnas_actividades = {fila[1] for fila in cursor.execute("PRAGMA table_info(actividades_soporte)").fetchall()}
    if "usuario_id" not in columnas_actividades:
        cursor.execute("ALTER TABLE actividades_soporte ADD COLUMN usuario_id INTEGER")
    if "whatsapp_telefono" not in columnas_actividades:
        cursor.execute("ALTER TABLE actividades_soporte ADD COLUMN whatsapp_telefono TEXT")
    
    # Crear esquema de ApowerREC (grabaciones de pantalla)
    ensure_apowerrec_schema(conn)

    # ── 7. TABLA DE CORTES DE REPORTES SGD ──
    # Almacena los resultados procesados de los reportes del SGD nacional.
    # El SGD (Sistema de Gestión Documental del Estado) exporta un Excel con
    # documentos enviados/no leídos por destinatario. Este módulo los procesa
    # y genera estadísticas, rankings y mapas de calor por fecha.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS reportes_sgd_cortes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo_estado TEXT NOT NULL,       -- Tipo de reporte: "NO LEIDOS", "ENVIADOS", etc.
        fecha_corte TEXT NOT NULL,       -- Fecha del corte del reporte
        archivo_origen TEXT,             -- Nombre del archivo Excel original
        gran_total INTEGER NOT NULL DEFAULT 0,         -- Total general de documentos
        total_trabajadores INTEGER NOT NULL DEFAULT 0, -- Cantidad de destinatarios
        datos_json TEXT NOT NULL,        -- JSON con toda la matriz de datos procesados
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # Migración: agregar campo para ocultar destinatarios en el reporte
    columnas_cortes_sgd = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(reportes_sgd_cortes)").fetchall()
    }
    if "destinatarios_ocultos" not in columnas_cortes_sgd:
        cursor.execute("ALTER TABLE reportes_sgd_cortes ADD COLUMN destinatarios_ocultos TEXT NOT NULL DEFAULT '[]'")

    # ── 8. TABLA DE DOCUMENTOS REGISTRADOS ──
    # Registro de documentos oficiales: informes, memorandos, oficios, etc.
    # Soporta numeración automática configurable y subida a Google Drive.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS documentos_registrados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo_documento TEXT NOT NULL,    -- INFORME, MEMORANDO, OFICIO, etc.
        numero TEXT,                     -- Número del documento (ej: "INFORME N° 001-2026-VCEM-SI")
        numero_ordinal INTEGER DEFAULT 0,  -- Correlativo numérico para auto-incremento
        titulo TEXT NOT NULL,            -- Asunto/título del documento
        descripcion TEXT,
        remitente TEXT,                  -- Quién emite
        destinatario TEXT,               -- A quién va dirigido
        fecha_emision TEXT,
        requerimiento TEXT,              -- Referencia al requerimiento que lo origina
        uso_documento TEXT,              -- Para qué se usa (gestión interna, etc.)
        resumen_gemini TEXT,             -- Resumen generado por IA (si se configuró)
        archivo_nombre TEXT,             -- Nombre del archivo PDF
        archivo_ruta TEXT,               -- Ruta en disco
        drive_url TEXT,                  -- URL en Google Drive (si se subió)
        drive_file_id TEXT,              -- ID del archivo en Drive
        fuente TEXT DEFAULT 'MANUAL',    -- MANUAL o IMPORTADO
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # ── 9. TABLA DE CONFIGURACIONES DEL SISTEMA ──
    # Almacena configuraciones generales como el patrón de numeración de documentos.
    # Estructura clave-valor flexible para cualquier configuración futura.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS configuraciones_sistema (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        clave TEXT UNIQUE NOT NULL,  -- Nombre de la configuración
        valor TEXT,                  -- Valor (puede ser texto, JSON, etc.)
        actualizado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # ── 10. TABLA DE CERTIFICADOS DIGITALES / TOKENS ──
    # Registro de certificados de firma digital (tokens USB, certificados de software).
    # Rastreo de fechas de vigencia para alertar caducidad.
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS certificados_digitales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        dni TEXT NOT NULL,                  -- DNI del titular
        apellido_paterno TEXT NOT NULL,
        apellido_materno TEXT NOT NULL,
        nombres TEXT NOT NULL,
        cargo TEXT,
        oficina TEXT,
        correo TEXT,
        tipo_certificado TEXT NOT NULL,     -- TOKEN o PARA PC
        trabajador_id INTEGER,
        numero_resolucion TEXT,
        fecha_inicio TEXT NOT NULL,         -- Fecha de emisión del certificado
        fecha_caducidad TEXT NOT NULL,      -- Fecha de vencimiento
        estado TEXT NOT NULL DEFAULT 'VIGENTE'
            CHECK(estado IN ('VIGENTE', 'RETIRADO')),
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        actualizado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    columnas_certificados = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(certificados_digitales)").fetchall()
    }
    if "trabajador_id" not in columnas_certificados:
        cursor.execute("ALTER TABLE certificados_digitales ADD COLUMN trabajador_id INTEGER")
    if "numero_resolucion" not in columnas_certificados:
        cursor.execute("ALTER TABLE certificados_digitales ADD COLUMN numero_resolucion TEXT")

    # Índice para evitar duplicados exactos de certificados
    cursor.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_certificados_digitales_unico
        ON certificados_digitales(dni, tipo_certificado, fecha_inicio, fecha_caducidad)
    """)

    # Migraciones de columnas para documentos_registrados
    columnas_documentos = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(documentos_registrados)").fetchall()
    }
    if "numero_ordinal" not in columnas_documentos:
        cursor.execute("ALTER TABLE documentos_registrados ADD COLUMN numero_ordinal INTEGER DEFAULT 0")
    if "uso_documento" not in columnas_documentos:
        cursor.execute("ALTER TABLE documentos_registrados ADD COLUMN uso_documento TEXT")
    if "usuario_id" not in columnas_documentos:
        cursor.execute("ALTER TABLE documentos_registrados ADD COLUMN usuario_id INTEGER")

    # Crear esquema de la Matriz de Accesos (Cargos × Sistemas)
    ensure_access_matrix_schema(conn)
    
    # Corregir tipos MODULO/SUBMODULO según jerarquía
    cursor.execute("""
        UPDATE sistemas_modulos_ut
        SET tipo = CASE
            WHEN padre_id IS NULL
              OR (SELECT tipo FROM sistemas_modulos_ut padre WHERE padre.id = sistemas_modulos_ut.padre_id) = 'SISTEMA'
            THEN 'MODULO'
            ELSE 'SUBMODULO'
        END
        WHERE tipo IN ('MODULO', 'SUBMODULO')
    """)

    # ── 11. ÍNDICES DE RENDIMIENTO ──
    # Aceleran las consultas más frecuentes del sistema.
    # Sin estos índices, las búsquedas en tablas grandes serían lentas.
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_prestamos_equipo ON prestamos(equipo_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_prestamos_trabajador ON prestamos(trabajador_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_prestamos_estado ON prestamos(estado)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_equipos_categoria_estado ON equipos(categoria, estado)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_actividades_usuario_fecha ON actividades_soporte(usuario_id, fecha_programada)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_documentos_tipo_fecha ON documentos_registrados(tipo_documento, fecha_emision)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_trabajadores_activo ON trabajadores(activo)")
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_certificados_digitales_caducidad "
        "ON certificados_digitales(estado, fecha_caducidad)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_certificados_digitales_persona "
        "ON certificados_digitales(dni, tipo_certificado)"
    )

    ensure_input_validation_schema(conn)
    conn.commit()
    conn.close()


# ════════════════════════════════════════════════════════════
# INICIALIZACIÓN DE LA BASE DE DATOS DE BIBLIOTECA PAE
# ════════════════════════════════════════════════════════════

def init_biblioteca_db():
    """Inicializa la base de datos auxiliar de Biblioteca PAE.
    
    Crea la tabla biblioteca_pae que almacena enlaces a eventos,
    capacitaciones y sesiones del Programa de Alimentación Escolar.
    
    Campos:
      - fecha_publicacion: Cuándo se publicó/realizó el evento
      - nombre_evento: Título del evento o capacitación
      - link: URL del evento (puede ser enlace a video, documento, etc.)
      - observaciones: Notas adicionales
    """
    conn = get_biblioteca_db()
    conn.execute("""
    CREATE TABLE IF NOT EXISTS biblioteca_pae (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_publicacion TEXT,
        nombre_evento TEXT NOT NULL,
        link TEXT NOT NULL,
        observaciones TEXT,
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    ensure_biblioteca_validation_schema(conn)
    conn.commit()
    conn.close()