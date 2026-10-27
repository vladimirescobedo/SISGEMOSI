# ============================================================================
# config.py — Configuración Central de SISGEMOSI
# ============================================================================
# Este archivo centraliza TODAS las configuraciones del sistema:
#   - Rutas de bases de datos SQLite
#   - Carpetas de almacenamiento de archivos subidos
#   - Categorías de equipos reconocidas
#   - Configuración del servidor web (host, puerto, tamaño máximo de uploads)
#
# Las variables de entorno (como SISGEMOSI_SECRET_KEY) se leen aquí
# para permitir configuración diferente entre desarrollo y producción.
# ============================================================================

import os

# Directorio raíz del proyecto (donde está este archivo config.py)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# ────────────────────────────────────────────────────────────
# Configuración de Seguridad
# ────────────────────────────────────────────────────────────
# Clave secreta para firmar cookies de sesión y tokens CSRF.
# En producción (Render.com) se define como variable de entorno.
# En desarrollo, app.py usa una clave fija por defecto.
SECRET_KEY = os.environ.get("SISGEMOSI_SECRET_KEY")

# ────────────────────────────────────────────────────────────
# Rutas de Bases de Datos SQLite
# ────────────────────────────────────────────────────────────
# El sistema usa DOS bases de datos SQLite separadas:
#   1. sistema_prestamos.db — BD principal con usuarios, trabajadores,
#      equipos, préstamos, actividades, documentos, SGD, tokens, etc.
#   2. biblioteca_pae.db — BD auxiliar para la Biblioteca PAE
#      (enlaces a eventos/capacitaciones del Programa de Alimentación Escolar)
DB_PATH = os.path.join(BASE_DIR, "sistema_prestamos.db")
BIBLIOTECA_DB_PATH = os.path.join(BASE_DIR, "biblioteca_pae.db")

# ────────────────────────────────────────────────────────────
# Carpetas de Almacenamiento de Archivos
# ────────────────────────────────────────────────────────────
# Cada módulo tiene su propia carpeta para archivos subidos.
# Se crean automáticamente si no existen (ver app.py, os.makedirs).
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")            # Archivos generales
ACTAS_FOLDER = os.path.join(BASE_DIR, "actas_generadas")     # Actas de entrega/devolución (Word)
PLANTILLAS_FOLDER = os.path.join(BASE_DIR, "plantillas")     # Plantillas .docx base
SGD_FOLDER = os.path.join(BASE_DIR, "uploads_sgd")           # Archivos Excel del SGD
DOCUMENTOS_FOLDER = os.path.join(BASE_DIR, "uploads_documentos")  # PDFs de documentos registrados
CONTRATOS_FOLDER = os.path.join(UPLOAD_FOLDER, "contratos")  # Contratos OS/CAS de trabajadores
BIBLIOTECA_FOLDER = os.path.join(UPLOAD_FOLDER, "biblioteca")  # Archivos de biblioteca PAE
BACKUPS_FOLDER = os.path.join(BASE_DIR, "backups")           # Respaldos de la BD

# ────────────────────────────────────────────────────────────
# Categorías de Equipos
# ────────────────────────────────────────────────────────────
# Estas son las categorías reconocidas en el módulo de Préstamos.
# Cada categoría tiene su propio dashboard con vista de tarjetas.
CATEGORIAS_EQUIPOS = ["MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"]

# ────────────────────────────────────────────────────────────
# Configuración del Servidor Web
# ────────────────────────────────────────────────────────────
HOST = "0.0.0.0"  # Escuchar en todas las interfaces de red
PORT = int(os.environ.get("PORT", 5000))  # Puerto (Render.com asigna uno vía $PORT)
MAX_CONTENT_LENGTH = 32 * 1024 * 1024  # Tamaño máximo de archivos subidos: 32 MB
