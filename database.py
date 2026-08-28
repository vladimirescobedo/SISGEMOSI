import sqlite3

def get_db():
    conn = sqlite3.connect("sistema_prestamos.db")
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios_sistema (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        nombre_completo TEXT NOT NULL,
        rol TEXT NOT NULL DEFAULT 'USUARIO',
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS trabajadores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        dni TEXT UNIQUE NOT NULL,
        nombres TEXT NOT NULL,
        apellido_paterno TEXT NOT NULL,
        apellido_materno TEXT NOT NULL,
        cargo TEXT,
        correo_personal TEXT,
        correo_institucional TEXT,
        tipo_contrato TEXT,
        telefono TEXT,
        fecha_nacimiento TEXT
    )
    """)

    columnas_trabajadores = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(trabajadores)").fetchall()
    }
    if "celular_institucional" not in columnas_trabajadores:
        cursor.execute("ALTER TABLE trabajadores ADD COLUMN celular_institucional TEXT")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS equipos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo_margesi TEXT,
        codigo_interno TEXT,
        categoria TEXT NOT NULL,
        tipo_equipo TEXT,
        descripcion TEXT,
        marca TEXT,
        modelo TEXT,
        serie_medidas TEXT,
        numero_serie TEXT,
        estado_conservacion TEXT,
        color TEXT,
        sede TEXT,
        valor_inicial TEXT,
        valor_neto TEXT,
        observaciones TEXT,
        estado TEXT DEFAULT 'EN ALMACEN'
    )
    """)

    columnas_equipos = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(equipos)").fetchall()
    }
    if "celular_institucional" not in columnas_equipos:
        cursor.execute("ALTER TABLE equipos ADD COLUMN celular_institucional TEXT")
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS prestamos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        equipo_id INTEGER,
        trabajador_id INTEGER,
        fecha_prestamo TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        fecha_devolucion TIMESTAMP,
        estado TEXT DEFAULT 'ACTIVO',
        observaciones_entrega TEXT,
        observaciones_devolucion TEXT,
        archivo_acta_entrega TEXT,
        FOREIGN KEY(equipo_id) REFERENCES equipos(id),
        FOREIGN KEY(trabajador_id) REFERENCES trabajadores(id)
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS actividades_soporte (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        descripcion TEXT,
        categoria TEXT DEFAULT 'SOPORTE TECNICO',
        area_solicitante TEXT,
        prioridad TEXT DEFAULT 'MEDIA',
        estado TEXT DEFAULT 'PENDIENTE',
        fecha_programada TEXT,
        hora_inicio TEXT,
        hora_fin TEXT,
        solucion_aplicada TEXT,
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # TABLA PARA GUARDAR PERMANENTEMENTE LOS CORTES DE REPORTE SGD
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS reportes_sgd_cortes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo_estado TEXT NOT NULL,
        fecha_corte TEXT NOT NULL,
        archivo_origen TEXT,
        gran_total INTEGER NOT NULL DEFAULT 0,
        total_trabajadores INTEGER NOT NULL DEFAULT 0,
        datos_json TEXT NOT NULL,
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    conn.commit()
    conn.close()