import sqlite3

def get_db():
    conn = sqlite3.connect("sistema_prestamos.db", timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA busy_timeout = 15000;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    conn.execute("PRAGMA cache_size = -32000;")
    return conn

def get_biblioteca_db():
    conn = sqlite3.connect("biblioteca_pae.db", timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=15000;")
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
        permisos TEXT NOT NULL DEFAULT '{}',
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    columnas_usuarios = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(usuarios_sistema)").fetchall()
    }
    if "permisos" not in columnas_usuarios:
        cursor.execute("ALTER TABLE usuarios_sistema ADD COLUMN permisos TEXT NOT NULL DEFAULT '{}'")
    permisos_legacy = '{"personal":["ver","modificar","eliminar","reportar"],"prestamos":["ver","modificar","eliminar","reportar"],"actividades":["ver","modificar","eliminar","reportar"],"documentaria":["ver","modificar","eliminar","reportar"],"sgd":["ver","modificar","eliminar","reportar"],"reportes":["ver","reportar"]}'
    cursor.execute("UPDATE usuarios_sistema SET permisos = ? WHERE permisos IS NULL", (permisos_legacy,))

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
        fecha_nacimiento TEXT,
        activo INTEGER NOT NULL DEFAULT 1
    )
    """)

    columnas_trabajadores = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(trabajadores)").fetchall()
    }
    if "celular_institucional" not in columnas_trabajadores:
        cursor.execute("ALTER TABLE trabajadores ADD COLUMN celular_institucional TEXT")
    if "activo" not in columnas_trabajadores:
        cursor.execute("ALTER TABLE trabajadores ADD COLUMN activo INTEGER NOT NULL DEFAULT 1")

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
    columnas_actividades = {fila[1] for fila in cursor.execute("PRAGMA table_info(actividades_soporte)").fetchall()}
    if "usuario_id" not in columnas_actividades:
        cursor.execute("ALTER TABLE actividades_soporte ADD COLUMN usuario_id INTEGER")

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
    columnas_cortes_sgd = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(reportes_sgd_cortes)").fetchall()
    }
    if "destinatarios_ocultos" not in columnas_cortes_sgd:
        cursor.execute("ALTER TABLE reportes_sgd_cortes ADD COLUMN destinatarios_ocultos TEXT NOT NULL DEFAULT '[]'")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS documentos_registrados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo_documento TEXT NOT NULL,
        numero TEXT,
        numero_ordinal INTEGER DEFAULT 0,
        titulo TEXT NOT NULL,
        descripcion TEXT,
        remitente TEXT,
        destinatario TEXT,
        fecha_emision TEXT,
        requerimiento TEXT,
        uso_documento TEXT,
        resumen_gemini TEXT,
        archivo_nombre TEXT,
        archivo_ruta TEXT,
        drive_url TEXT,
        drive_file_id TEXT,
        fuente TEXT DEFAULT 'MANUAL',
        creado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS configuraciones_sistema (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        clave TEXT UNIQUE NOT NULL,
        valor TEXT,
        actualizado_el TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    columnas_documentos = {
        fila[1] for fila in cursor.execute("PRAGMA table_info(documentos_registrados)").fetchall()
    }
    if "numero_ordinal" not in columnas_documentos:
        cursor.execute("ALTER TABLE documentos_registrados ADD COLUMN numero_ordinal INTEGER DEFAULT 0")
    if "uso_documento" not in columnas_documentos:
        cursor.execute("ALTER TABLE documentos_registrados ADD COLUMN uso_documento TEXT")
    if "usuario_id" not in columnas_documentos:
        cursor.execute("ALTER TABLE documentos_registrados ADD COLUMN usuario_id INTEGER")

    # ===== ÍNDICES DE RENDIMIENTO =====
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_prestamos_equipo ON prestamos(equipo_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_prestamos_trabajador ON prestamos(trabajador_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_prestamos_estado ON prestamos(estado)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_equipos_categoria_estado ON equipos(categoria, estado)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_actividades_usuario_fecha ON actividades_soporte(usuario_id, fecha_programada)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_documentos_tipo_fecha ON documentos_registrados(tipo_documento, fecha_emision)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_trabajadores_activo ON trabajadores(activo)")

    conn.commit()
    conn.close()

def init_biblioteca_db():
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
    conn.commit()
    conn.close()