import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# Configuración de Seguridad
SECRET_KEY = os.environ.get("SISGEMOSI_SECRET_KEY")

# Rutas de Bases de Datos SQLite
DB_PATH = os.path.join(BASE_DIR, "sistema_prestamos.db")
BIBLIOTECA_DB_PATH = os.path.join(BASE_DIR, "biblioteca_pae.db")

# Carpetas de Almacenamiento
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
ACTAS_FOLDER = os.path.join(BASE_DIR, "actas_generadas")
PLANTILLAS_FOLDER = os.path.join(BASE_DIR, "plantillas")
SGD_FOLDER = os.path.join(BASE_DIR, "uploads_sgd")
DOCUMENTOS_FOLDER = os.path.join(BASE_DIR, "uploads_documentos")
CONTRATOS_FOLDER = os.path.join(UPLOAD_FOLDER, "contratos")
BIBLIOTECA_FOLDER = os.path.join(UPLOAD_FOLDER, "biblioteca")
BACKUPS_FOLDER = os.path.join(BASE_DIR, "backups")

# Categorías reconocidas en el sistema
CATEGORIAS_EQUIPOS = ["MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"]

# Configuración del Servidor
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", 5000))
MAX_CONTENT_LENGTH = 32 * 1024 * 1024  # 32 MB max upload

