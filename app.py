import os
import json
import hashlib
import hmac
import re
import secrets
import unicodedata
import urllib.parse
import urllib.request
from xml.sax.saxutils import escape as escape_xml
from io import BytesIO
from datetime import datetime, date, timedelta
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, session
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from database import get_db, init_db, get_biblioteca_db, init_biblioteca_db
from excel_loader import importar_trabajadores_excel, formatear_solo_fecha
from equipos_loader import importar_equipos_excel
from docx_generator import rellenar_plantilla_docx, generar_informe_mensual_docx
from docx_generator import rellenar_plantilla_docx, generar_informe_mensual_docx, generar_reporte_actividades_docx
from sgd_processor import procesar_sgd_excel, generar_excel_sgd, aplicar_destinatarios_ocultos

app = Flask(__name__)
PRODUCTION_MODE = os.environ.get("SISGEMOSI_ENV", "development").lower() == "production"
secret_key = os.environ.get("SISGEMOSI_SECRET_KEY")
if PRODUCTION_MODE and not secret_key:
    raise RuntimeError("Defina SISGEMOSI_SECRET_KEY antes de iniciar en producción.")
app.secret_key = secret_key or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=PRODUCTION_MODE,
    MAX_CONTENT_LENGTH=32 * 1024 * 1024,
)

UPLOAD_FOLDER = "uploads"
ACTAS_FOLDER = "actas_generadas"
PLANTILLAS_FOLDER = "plantillas"
SGD_FOLDER = "uploads_sgd"
DOCUMENTOS_FOLDER = "uploads_documentos"
CONTRATOS_FOLDER = os.path.join(UPLOAD_FOLDER, "contratos")
BIBLIOTECA_FOLDER = os.path.join(UPLOAD_FOLDER, "biblioteca")

def guardar_archivo_subido(archivo, carpeta, extensiones_permitidas):
    nombre_original = secure_filename(archivo.filename or "")
    extension = os.path.splitext(nombre_original)[1].lower()
    if not nombre_original or extension not in extensiones_permitidas:
        raise ValueError("El tipo de archivo no está permitido.")
    os.makedirs(carpeta, exist_ok=True)
    nombre_guardado = f"{secrets.token_hex(16)}{extension}"
    ruta = os.path.join(carpeta, nombre_guardado)
    archivo.save(ruta)
    return ruta, archivo.filename

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(ACTAS_FOLDER, exist_ok=True)
os.makedirs(PLANTILLAS_FOLDER, exist_ok=True)
os.makedirs(SGD_FOLDER, exist_ok=True)
os.makedirs(DOCUMENTOS_FOLDER, exist_ok=True)
os.makedirs(CONTRATOS_FOLDER, exist_ok=True)
os.makedirs(BIBLIOTECA_FOLDER, exist_ok=True)

def hash_pass(p):
    """Genera hash seguro con PBKDF2 (werkzeug) en vez de SHA-256 plano."""
    return generate_password_hash(p)

def _legacy_sha256(p):
    """Hash SHA-256 antiguo para compatibilidad con contraseñas existentes."""
    return hashlib.sha256(p.encode()).hexdigest()

def generar_credencial_trabajador(nombres, apellido_paterno, apellido_materno):
    def limpio(valor):
        texto = unicodedata.normalize("NFKD", valor or "").encode("ascii", "ignore").decode("ascii")
        return re.sub(r"[^A-Z]", "", texto.upper())

    nombres_limpios = limpio(nombres)
    paterno_limpio = limpio(apellido_paterno)
    materno_limpio = limpio(apellido_materno)
    return f"{nombres_limpios[:1]}{paterno_limpio}{materno_limpio[:1]}"

def sanitizar_texto_excel(val):
    """Limpia caracteres no imprimibles, emojis o ilegales para hojas de cálculo Excel (openpyxl)."""
    if val is None:
        return ""
    return ILLEGAL_CHARACTERS_RE.sub("", str(val))

def check_pass(p, h):
    """Verifica contraseña con soporte para hashes PBKDF2 y SHA-256 legacy.
    Si el hash es legacy SHA-256, lo migra automáticamente a PBKDF2."""
    # Intentar primero con werkzeug (PBKDF2)
    if h.startswith("pbkdf2:") or h.startswith("scrypt:"):
        return check_password_hash(h, p)
    # Fallback: verificar contra SHA-256 legacy
    if _legacy_sha256(p) == h:
        return True
    return False

def obtener_token_csrf():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)
    return session["csrf_token"]

@app.before_request
def proteger_formularios():
    if request.method != "POST" or request.endpoint in {"static", "login"}:
        return None

    token_formulario = request.form.get("csrf_token", "")
    token_sesion = session.get("csrf_token", "")
    if not token_sesion or not token_formulario or not hmac.compare_digest(token_formulario, token_sesion):
        return "Solicitud no válida.", 400
    return None

@app.after_request
def agregar_cabeceras_seguridad(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    if PRODUCTION_MODE:
        response.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
    return response

def get_document_number_pattern_config():
    config = {
        "year": datetime.now().year,
        "sufijo": "VCEM-SI",
        "patron": "{tipo} N° {numero:03d}-{year}-{sufijo}"
    }

    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT clave, valor FROM configuraciones_sistema WHERE clave IN (?, ?, ?)",
            ("document_number_pattern", "document_number_year", "document_number_suffix")
        ).fetchall()
        for row in rows:
            clave = row["clave"]
            valor = row["valor"]
            if clave == "document_number_pattern" and valor:
                try:
                    data = json.loads(valor)
                    if isinstance(data, dict):
                        config.update({k: v for k, v in data.items() if v not in (None, "")})
                except Exception:
                    config["patron"] = valor
            elif clave == "document_number_year" and valor:
                try:
                    config["year"] = int(valor)
                except ValueError:
                    pass
            elif clave == "document_number_suffix" and valor:
                config["sufijo"] = valor.strip() or config["sufijo"]
    finally:
        conn.close()

    return config


def guardar_document_number_pattern_config(config):
    conn = get_db()
    try:
        patron = (config.get("patron") or "{tipo} N° {numero:03d}-{year}-{sufijo}").strip()
        year = config.get("year") or datetime.now().year
        sufijo = (config.get("sufijo") or "VCEM-SI").strip()

        conn.execute(
            "INSERT INTO configuraciones_sistema (clave, valor) VALUES (?, ?) ON CONFLICT(clave) DO UPDATE SET valor = excluded.valor, actualizado_el = CURRENT_TIMESTAMP",
            ("document_number_pattern", json.dumps({"patron": patron}, ensure_ascii=False))
        )
        conn.execute(
            "INSERT INTO configuraciones_sistema (clave, valor) VALUES (?, ?) ON CONFLICT(clave) DO UPDATE SET valor = excluded.valor, actualizado_el = CURRENT_TIMESTAMP",
            ("document_number_year", str(int(year)))
        )
        conn.execute(
            "INSERT INTO configuraciones_sistema (clave, valor) VALUES (?, ?) ON CONFLICT(clave) DO UPDATE SET valor = excluded.valor, actualizado_el = CURRENT_TIMESTAMP",
            ("document_number_suffix", sufijo)
        )
        conn.commit()
    finally:
        conn.close()


def generar_numero_documento(tipo_documento, conn=None):
    tipo = (tipo_documento or "").upper().strip()
    if not tipo:
        return None

    config = get_document_number_pattern_config()
    year = int(config.get("year") or datetime.now().year)
    sufijo = str(config.get("sufijo") or "VCEM-SI")
    patron = (config.get("patron") or "{tipo} N° {numero:03d}-{year}-{sufijo}").strip()

    db_conn = conn or get_db()
    try:
        ultimo = db_conn.execute(
            "SELECT numero_ordinal FROM documentos_registrados WHERE tipo_documento = ? AND fecha_emision LIKE ? ORDER BY numero_ordinal DESC, id DESC LIMIT 1",
            (tipo, f"{year}%")
        ).fetchone()
    finally:
        if conn is None:
            db_conn.close()

    ultimo_numero = 0
    if ultimo and ultimo["numero_ordinal"]:
        ultimo_numero = int(ultimo["numero_ordinal"])

    siguiente = ultimo_numero + 1
    return patron.format(tipo=tipo, numero=siguiente, year=year, sufijo=sufijo)


def extraer_texto_archivo(ruta_archivo):
    nombre = (ruta_archivo or "").lower()
    if nombre.endswith((".txt", ".md", ".csv")):
        try:
            with open(ruta_archivo, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception:
            return None
    return None

def analizar_con_gemini(texto_base=None, ruta_archivo=None):
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        return {
            "estado": "sin_configuracion",
            "resumen": "No está configurada la clave API de Gemini. Se guardó el documento y se puede analizar cuando se configure la credencial.",
        }

    texto_final = (texto_base or "").strip()
    if ruta_archivo:
        texto_archivo = extraer_texto_archivo(ruta_archivo)
        if texto_archivo:
            texto_final = (texto_final + "\n\n" + texto_archivo).strip()

    if not texto_final:
        return {
            "estado": "sin_texto",
            "resumen": "El documento no tiene texto extraíble para analizar con Gemini. Puedes escribir el contenido o requerimiento manualmente.",
        }

    prompt = (
        "Eres un asistente de gestión documentaria para la Universidad. Analiza el texto e identifica: "
        "1) el contenido principal del documento, 2) el requerimiento o pedido específico, 3) el destinatario o área involucrada y 4) el posible plazo o acción sugerida. "
        "Responde en español, en un formato claro y breve, con un resumen ejecutivo."
    )

    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"text": texto_final[:12000]}
            ]
        }],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 800
        }
    }

    try:
        request_obj = urllib.request.Request(
            f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(request_obj, timeout=30) as response:
            respuesta = json.loads(response.read().decode("utf-8"))
        texto_resumen = respuesta.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text")
        if texto_resumen:
            return {"estado": "ok", "resumen": texto_resumen.strip()}
    except Exception:
        pass

    return {
        "estado": "error_api",
        "resumen": "Se intentó analizar el documento con Gemini, pero la API no respondió correctamente. El archivo quedó registrado y puede revisarse manualmente.",
    }

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("usuario_id"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("usuario_id") or session.get("rol") != "ADMIN":
            flash("Acceso restringido: Se requieren permisos de Administrador.", "danger")
            return redirect(request.referrer or url_for("inicio"))
        return f(*args, **kwargs)
    return decorated_function

PERMISSION_MODULES = {
    "personal": "Personal",
    "prestamos": "Préstamos",
    "actividades": "Actividades TI",
    "documentaria": "Gestión documentaria",
    "sgd": "Monitor SGD",
    "reportes": "Reportes",
    "biblioteca": "Biblioteca Virtual PAE",
}
PERMISSION_ACTIONS = {
    "ver": "Ver",
    "modificar": "Crear / modificar",
    "eliminar": "Eliminar",
    "reportar": "Exportar / reportar",
}

def permisos_completos():
    return {modulo: list(PERMISSION_ACTIONS) for modulo in PERMISSION_MODULES}

def permisos_desde_formulario(formulario):
    permisos = {modulo: [] for modulo in PERMISSION_MODULES}
    for valor in formulario.getlist("permiso"):
        try:
            modulo, accion = valor.split(":", 1)
        except ValueError:
            continue
        if modulo in PERMISSION_MODULES and accion in PERMISSION_ACTIONS:
            permisos[modulo].append(accion)
    return permisos

def cargar_permisos(valor):
    try:
        permisos = json.loads(valor or "{}")
        return permisos if isinstance(permisos, dict) else {}
    except (TypeError, ValueError):
        return {}

def usuario_tiene_permiso(modulo, accion):
    if session.get("rol") == "ADMIN":
        return True
    usuario_id = session.get("usuario_id")
    if not usuario_id:
        return False
    if "permisos" not in session:
        conn = get_db()
        usuario = conn.execute("SELECT permisos FROM usuarios_sistema WHERE id = ?", (usuario_id,)).fetchone()
        conn.close()
        session["permisos"] = cargar_permisos(usuario["permisos"] if usuario else "")
    permisos = session.get("permisos", {})
    return accion in permisos.get(modulo, [])

def filtro_propietario(columna="usuario_id"):
    if session.get("rol") == "ADMIN":
        return "1=1", []
    return f"{columna} = ?", [session.get("usuario_id")]

def permiso_requerido(modulo, accion):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not session.get("usuario_id"):
                return redirect(url_for("login"))
            if not usuario_tiene_permiso(modulo, accion):
                flash("No tiene permiso para realizar esta acción.", "danger")
                return redirect(request.referrer or url_for("inicio"))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

def calcular_duracion(f_ini_str, f_fin_str=None):
    if not f_ini_str:
        return "-"
    try:
        fmt = "%Y-%m-%d %H:%M:%S" if " " in str(f_ini_str) else "%Y-%m-%d"
        dt_ini = datetime.strptime(str(f_ini_str).split(".")[0], fmt)
        
        if f_fin_str:
            fmt_fin = "%Y-%m-%d %H:%M:%S" if " " in str(f_fin_str) else "%Y-%m-%d"
            dt_fin = datetime.strptime(str(f_fin_str).split(".")[0], fmt_fin)
        else:
            dt_fin = datetime.now()
            
        diff = dt_fin - dt_ini
        dias = diff.days
        horas = diff.seconds // 3600
        minutos = (diff.seconds % 3600) // 60
        
        partes = []
        if dias > 0:
            partes.append(f"{dias} día{'s' if dias > 1 else ''}")
        if horas > 0:
            partes.append(f"{horas} h")
        if minutos > 0 or (dias == 0 and horas == 0):
            partes.append(f"{minutos} min")
            
        return ", ".join(partes)
    except Exception:
        return "-"

@app.context_processor
def utility_processor():
    return dict(calcular_duracion=calcular_duracion, usuario_tiene_permiso=usuario_tiene_permiso, csrf_token=obtener_token_csrf)

@app.route("/favicon.ico")
def favicon():
    return redirect(url_for("static", filename="logo-pae.svg", v="2"))

# ================= LOGIN Y AUTENTICACIÓN =================
@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("usuario_id"):
        return redirect(url_for("inicio"))

    if request.method == "POST":
        user_input = (request.form.get("usuario") or "").strip()
        pass_input = (request.form.get("password") or "").strip()

        if not user_input or not pass_input:
            flash("Debe ingresar usuario y contraseña.", "warning")
            return render_template("login.html")

        conn = get_db()
        cursor = conn.cursor()
        try:
            usuario_db = cursor.execute("SELECT * FROM usuarios_sistema WHERE usuario = ?", (user_input,)).fetchone()
        finally:
            conn.close()

        if usuario_db and check_pass(pass_input, usuario_db["password"]):
            if not usuario_db["password"].startswith("pbkdf2:"):
                conn2 = get_db()
                conn2.execute(
                    "UPDATE usuarios_sistema SET password = ? WHERE id = ?",
                    (generate_password_hash(pass_input), usuario_db["id"]),
                )
                conn2.commit()
                conn2.close()
            session["usuario_id"] = usuario_db["id"]
            session["usuario"] = usuario_db["usuario"]
            session["nombre"] = usuario_db["nombre_completo"]
            session["rol"] = usuario_db["rol"]
            session["permisos"] = cargar_permisos(usuario_db["permisos"])
            flash(f"¡Bienvenido(a), {usuario_db['nombre_completo']}!", "success")
            return redirect(url_for("inicio"))

        flash("Usuario o contraseña incorrectos.", "danger")

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    flash("Ha cerrado sesión correctamente.", "info")
    return redirect(url_for("login"))

@app.route("/cambiar-clave", methods=["POST"])
@login_required
def cambiar_clave():
    clave_actual = request.form.get("clave_actual", "").strip()
    clave_nueva = request.form.get("clave_nueva", "").strip()
    clave_confirmar = request.form.get("clave_confirmar", "").strip()

    if not clave_actual or not clave_nueva:
        flash("Debe completar todos los campos.", "warning")
        return redirect(request.referrer or url_for("inicio"))

    if clave_nueva != clave_confirmar:
        flash("La nueva contraseña y la confirmación no coinciden.", "danger")
        return redirect(request.referrer or url_for("inicio"))

    if len(clave_nueva) < 8:
        flash("La nueva contraseña debe tener al menos 8 caracteres.", "warning")
        return redirect(request.referrer or url_for("inicio"))

    conn = get_db()
    usuario = conn.execute("SELECT password FROM usuarios_sistema WHERE id = ?",
                           (session["usuario_id"],)).fetchone()
    if not usuario or not check_pass(clave_actual, usuario["password"]):
        conn.close()
        flash("La contraseña actual es incorrecta.", "danger")
        return redirect(request.referrer or url_for("inicio"))

    conn.execute("UPDATE usuarios_sistema SET password = ? WHERE id = ?",
                 (generate_password_hash(clave_nueva), session["usuario_id"]))
    conn.commit()
    conn.close()
    flash("Contraseña actualizada exitosamente.", "success")
    return redirect(request.referrer or url_for("inicio"))

@app.errorhandler(404)
def pagina_no_encontrada(e):
    return render_template("error.html", codigo=404, mensaje="Página no encontrada",
                           detalle="La página que busca no existe o ha sido movida."), 404

@app.errorhandler(500)
def error_interno(e):
    return render_template("error.html", codigo=500, mensaje="Error interno del servidor",
                           detalle="Ocurrió un error inesperado. Intente nuevamente o contacte al administrador."), 500

# ================= GESTIÓN DE USUARIOS (SOLO ADMINISTRADOR) =================
@app.route("/usuarios")
@admin_required
def ver_usuarios():
    conn = get_db()
    usuarios_db = conn.cursor().execute("SELECT id, usuario, nombre_completo, rol, permisos, creado_el FROM usuarios_sistema ORDER BY id ASC").fetchall()
    trabajadores_activos = conn.cursor().execute("""
        SELECT id, nombres, apellido_paterno, apellido_materno
        FROM trabajadores WHERE activo = 1
        ORDER BY apellido_paterno, apellido_materno, nombres
    """).fetchall()
    conn.close()
    usuarios = []
    for usuario in usuarios_db:
        item = dict(usuario)
        item["permisos"] = cargar_permisos(item["permisos"])
        usuarios.append(item)
    return render_template("usuarios.html", usuarios=usuarios,
                           trabajadores_activos=trabajadores_activos,
                           permission_modules=PERMISSION_MODULES,
                           permission_actions=PERMISSION_ACTIONS)

@app.route("/usuarios/nuevo", methods=["POST"])
@admin_required
def nuevo_usuario():
    trabajador_id = request.form.get("trabajador_id", type=int)
    rol_input = request.form.get("rol", "USUARIO").strip().upper()
    permisos = permisos_desde_formulario(request.form)
    if not any(valor.startswith("actividades:") for valor in request.form.getlist("permiso")):
        permisos["actividades"] = ["ver", "modificar"]
    if not any(valor.startswith("documentaria:") for valor in request.form.getlist("permiso")):
        permisos["documentaria"] = ["ver", "modificar"]

    conn = get_db()
    cursor = conn.cursor()
    trabajador = cursor.execute("""
        SELECT nombres, apellido_paterno, apellido_materno
        FROM trabajadores WHERE id = ? AND activo = 1
    """, (trabajador_id,)).fetchone()
    if not trabajador:
        conn.close()
        flash("Seleccione un trabajador activo.", "warning")
        return redirect(url_for("ver_usuarios"))

    nombre_input = f"{trabajador['nombres']} {trabajador['apellido_paterno']} {trabajador['apellido_materno']}".strip()
    user_input = generar_credencial_trabajador(trabajador["nombres"], trabajador["apellido_paterno"], trabajador["apellido_materno"])
    pass_input = secrets.token_urlsafe(12)

    try:
        cursor.execute("""
            INSERT INTO usuarios_sistema (usuario, password, nombre_completo, rol, permisos)
            VALUES (?, ?, ?, ?, ?)
        """, (user_input, hash_pass(pass_input), nombre_input, rol_input, json.dumps(permisos)))
        conn.commit()
        flash(f"Usuario '{user_input}' creado. Contraseña temporal: {pass_input}. Debe cambiarla al ingresar.", "success")
    except Exception:
        flash(f"Error: El usuario '{user_input}' ya existe.", "danger")
    finally:
        conn.close()

    return redirect(url_for("ver_usuarios"))

@app.route("/usuarios/editar", methods=["POST"])
@admin_required
def editar_usuario():
    user_id = request.form["usuario_id"]
    nombre_input = request.form["nombre_completo"].strip()
    pass_input = request.form.get("password", "").strip()
    rol_input = request.form.get("rol", "USUARIO").strip().upper()
    permisos = permisos_desde_formulario(request.form)

    if pass_input and len(pass_input) < 8:
        flash("La nueva contraseña debe tener al menos 8 caracteres.", "warning")
        return redirect(url_for("ver_usuarios"))

    conn = get_db()
    cursor = conn.cursor()

    if pass_input:
        cursor.execute("""
            UPDATE usuarios_sistema 
            SET nombre_completo = ?, password = ?, rol = ?, permisos = ?
            WHERE id = ?
        """, (nombre_input, hash_pass(pass_input), rol_input, json.dumps(permisos), user_id))
    else:
        cursor.execute("""
            UPDATE usuarios_sistema 
            SET nombre_completo = ?, rol = ?, permisos = ?
            WHERE id = ?
        """, (nombre_input, rol_input, json.dumps(permisos), user_id))

    conn.commit()
    conn.close()
    flash("Usuario actualizado correctamente.", "success")
    return redirect(url_for("ver_usuarios"))

@app.route("/usuarios/eliminar/<int:user_id>", methods=["POST"])
@admin_required
def eliminar_usuario(user_id):
    if user_id == session.get("usuario_id"):
        flash("No puede eliminar su propia cuenta activa.", "warning")
        return redirect(url_for("ver_usuarios"))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM usuarios_sistema WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    flash("Usuario eliminado correctamente.", "success")
    return redirect(url_for("ver_usuarios"))

# ================= PORTAL PRINCIPAL =================
@app.route("/")
@login_required
def inicio():
    conn = get_db()
    cursor = conn.cursor()

    total_moviles = cursor.execute("SELECT COUNT(*) FROM equipos WHERE categoria = 'MOVILES'").fetchone()[0]
    total_tokens = cursor.execute("SELECT COUNT(*) FROM equipos WHERE categoria = 'TOKENS'").fetchone()[0]
    total_informaticos = cursor.execute("SELECT COUNT(*) FROM equipos WHERE categoria = 'INFORMATICOS'").fetchone()[0]
    total_mobiliario = cursor.execute("SELECT COUNT(*) FROM equipos WHERE categoria = 'MOBILIARIO'").fetchone()[0]
    total_equipos = total_moviles + total_tokens + total_informaticos + total_mobiliario
    equipos_prestamo = cursor.execute("SELECT COUNT(*) FROM equipos WHERE estado = 'EN PRESTAMO'").fetchone()[0]
    total_personal = cursor.execute("SELECT COUNT(*) FROM trabajadores WHERE activo = 1").fetchone()[0]
    propietario_sql, propietario_params = filtro_propietario()
    total_documentos = cursor.execute(f"SELECT COUNT(*) FROM documentos_registrados WHERE {propietario_sql}", propietario_params).fetchone()[0]
    tareas_pendientes = cursor.execute(f"SELECT COUNT(*) FROM actividades_soporte WHERE estado != 'COMPLETADA' AND {propietario_sql}", propietario_params).fetchone()[0]
    tareas_hoy = cursor.execute(f"SELECT COUNT(*) FROM actividades_soporte WHERE fecha_programada = CURRENT_DATE AND {propietario_sql}", propietario_params).fetchone()[0]
    conn.close()

    return render_template("inicio.html", 
                           total_equipos=total_equipos,
                           total_moviles=total_moviles,
                           total_tokens=total_tokens,
                           total_informaticos=total_informaticos,
                           total_mobiliario=total_mobiliario,
                           equipos_prestamo=equipos_prestamo,
                           total_personal=total_personal,
                           tareas_pendientes=tareas_pendientes,
                           tareas_hoy=tareas_hoy,
                           total_documentos=total_documentos)

# ================= BIBLIOTECA VIRTUAL PAE UT APURIMAC =================
@app.route("/biblioteca")
@permiso_requerido("biblioteca", "ver")
@login_required
def ver_biblioteca():
    busqueda = request.args.get("q", "").strip()
    conn = get_biblioteca_db()
    if busqueda:
        patron = f"%{busqueda}%"
        registros = conn.execute("""
            SELECT * FROM biblioteca_pae
            WHERE fecha_publicacion LIKE ? OR nombre_evento LIKE ?
               OR link LIKE ? OR observaciones LIKE ?
            ORDER BY id DESC
        """, (patron, patron, patron, patron)).fetchall()
    else:
        registros = conn.execute("SELECT * FROM biblioteca_pae ORDER BY id DESC").fetchall()
    total = conn.execute("SELECT COUNT(*) FROM biblioteca_pae").fetchone()[0]
    conn.close()
    return render_template("biblioteca.html", registros=registros, busqueda=busqueda,
                           total_biblioteca=total)

@app.route("/biblioteca/nuevo", methods=["POST"])
@permiso_requerido("biblioteca", "modificar")
@login_required
def nuevo_registro_biblioteca():
    evento = request.form.get("nombre_evento", "").strip()
    link = request.form.get("link", "").strip()
    if not evento or not link:
        flash("El nombre del evento y el link son obligatorios.", "warning")
        return redirect(url_for("ver_biblioteca"))
    conn = get_biblioteca_db()
    conn.execute("""
        INSERT INTO biblioteca_pae (fecha_publicacion, nombre_evento, link, observaciones)
        VALUES (?, ?, ?, ?)
    """, (request.form.get("fecha_publicacion", "").strip(), evento, link,
          request.form.get("observaciones", "").strip()))
    conn.commit()
    conn.close()
    flash("Registro agregado a la Biblioteca Virtual.", "success")
    return redirect(url_for("ver_biblioteca"))

@app.route("/biblioteca/importar", methods=["POST"])
@permiso_requerido("biblioteca", "modificar")
@login_required
def importar_biblioteca_excel():
    archivo = request.files.get("archivo_biblioteca")
    if not archivo or not archivo.filename.lower().endswith(".xlsx"):
        flash("Seleccione un archivo Excel .xlsx válido.", "danger")
        return redirect(url_for("ver_biblioteca"))
    try:
        ruta_archivo, _ = guardar_archivo_subido(archivo, BIBLIOTECA_FOLDER, {".xlsx"})
        cantidad, omitidos = importar_registros_biblioteca(ruta_archivo)
        flash(f"Se importaron {cantidad} registros. Filas omitidas: {omitidos}.", "success")
    except ValueError as error:
        flash(str(error), "danger")
    except Exception as error:
        flash(f"No se pudo importar el archivo: {error}", "danger")
    return redirect(url_for("ver_biblioteca"))

def normalizar_encabezado(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    return re.sub(r"[^a-z0-9]", "", texto.lower())

def convertir_fecha_biblioteca(valor):
    if valor in (None, ""):
        return ""
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%Y-%m-%d")
    if isinstance(valor, (int, float)):
        try:
            return openpyxl.utils.datetime.from_ISO8601(str(valor)).strftime("%Y-%m-%d")
        except Exception:
            try:
                return openpyxl.utils.datetime.from_excel(valor).strftime("%Y-%m-%d")
            except Exception:
                pass
    return str(valor).strip()

def importar_registros_biblioteca(ruta_archivo):
    libro = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
    hoja = libro.active
    filas = hoja.iter_rows(values_only=True)
    encabezados = None
    fila_encabezados = None
    for numero, fila in enumerate(filas):
        normalizados = [normalizar_encabezado(valor) for valor in fila]
        tiene_fecha = any(valor.startswith("fecha") and ("grabacion" in valor or "publicacion" in valor) for valor in normalizados)
        tiene_evento = any(valor in ("nombredela ponencia", "nombredelaponencia", "nombreevento", "evento", "titulo") for valor in normalizados)
        tiene_link = any(valor in ("link", "enlace", "url", "direccion") for valor in normalizados)
        if tiene_fecha and tiene_evento and tiene_link:
            encabezados = normalizados
            fila_encabezados = numero
            break
    if encabezados is None:
        raise ValueError("No se encontraron encabezados reconocibles. Use fecha de grabación/publicación, nombre del evento o ponencia, link y observaciones.")

    def buscar_indice(opciones):
        for opcion in opciones:
            if opcion in encabezados:
                return encabezados.index(opcion)
        return None

    indice_fecha = buscar_indice(("fechadegrabacionpublicacion", "fechapublicacion", "fecha"))
    indice_evento = buscar_indice(("nombredelaponencia", "nombredel evento", "nombreevento", "evento", "titulo"))
    indice_link = buscar_indice(("link", "enlace", "url", "direccion"))
    indice_observaciones = buscar_indice(("observacion", "observaciones", "comentario", "notas"))
    if indice_evento is None or indice_link is None:
        raise ValueError("Faltan las columnas obligatorias de nombre de evento/ponencia y link.")

    conn = get_biblioteca_db()
    cantidad = 0
    omitidos = 0
    for fila in filas:
        valores = list(fila)
        evento = str(valores[indice_evento] or "").strip() if indice_evento < len(valores) else ""
        link = str(valores[indice_link] or "").strip() if indice_link < len(valores) else ""
        if not evento and not link:
            continue
        if not evento or not link:
            omitidos += 1
            continue
        fecha = convertir_fecha_biblioteca(valores[indice_fecha] if indice_fecha is not None and indice_fecha < len(valores) else "")
        observaciones = str(valores[indice_observaciones] or "").strip() if indice_observaciones is not None and indice_observaciones < len(valores) else ""
        existente = conn.execute("SELECT id FROM biblioteca_pae WHERE nombre_evento = ? AND link = ?", (evento, link)).fetchone()
        if existente:
            conn.execute("UPDATE biblioteca_pae SET fecha_publicacion = ?, observaciones = ? WHERE id = ?", (fecha, observaciones, existente["id"]))
        else:
            conn.execute("INSERT INTO biblioteca_pae (fecha_publicacion, nombre_evento, link, observaciones) VALUES (?, ?, ?, ?)", (fecha, evento, link, observaciones))
        cantidad += 1
    conn.commit()
    conn.close()
    libro.close()
    return cantidad, omitidos

@app.route("/biblioteca/eliminar/<int:registro_id>", methods=["POST"])
@permiso_requerido("biblioteca", "eliminar")
@login_required
def eliminar_registro_biblioteca(registro_id):
    conn = get_biblioteca_db()
    conn.execute("DELETE FROM biblioteca_pae WHERE id = ?", (registro_id,))
    conn.commit()
    conn.close()
    flash("Registro eliminado.", "success")
    return redirect(url_for("ver_biblioteca"))

@app.route("/biblioteca/exportar-excel")
@permiso_requerido("biblioteca", "reportar")
@login_required
def exportar_biblioteca_excel():
    busqueda = request.args.get("q", "").strip()
    conn = get_biblioteca_db()
    if busqueda:
        patron = f"%{busqueda}%"
        filas = conn.execute("SELECT fecha_publicacion, nombre_evento, link, observaciones FROM biblioteca_pae WHERE fecha_publicacion LIKE ? OR nombre_evento LIKE ? OR link LIKE ? OR observaciones LIKE ? ORDER BY id DESC", (patron, patron, patron, patron)).fetchall()
    else:
        filas = conn.execute("SELECT fecha_publicacion, nombre_evento, link, observaciones FROM biblioteca_pae ORDER BY id DESC").fetchall()
    conn.close()
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Biblioteca Virtual PAE"
    encabezados = ["FECHA DE GRABACIÓN/PUBLICACIÓN", "NOMBRE DEL EVENTO", "LINK", "OBSERVACIONES"]
    hoja.append(encabezados)
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="075985", end_color="075985", fill_type="solid")
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for fila in filas:
        hoja.append([fila["fecha_publicacion"] or "", fila["nombre_evento"] or "", fila["link"] or "", fila["observaciones"] or ""])
    for columna in hoja.columns:
        letra = openpyxl.utils.get_column_letter(columna[0].column)
        hoja.column_dimensions[letra].width = min(max(max(len(str(c.value or "")) for c in columna) + 3, 16), 70)
    buffer = BytesIO()
    libro.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Biblioteca_Virtual_PAE_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ================= MÓDULO: MONITOR DE SGD (HISTORIAL POR FECHAS DE CORTE EN BD) =================
@app.route("/sgd")
@permiso_requerido("sgd", "ver")
@login_required
def ver_sgd():
    tipo = request.args.get("tipo", "NO_LEIDOS").upper()
    if tipo not in ["NO_LEIDOS", "RECIBIDOS"]:
        tipo = "NO_LEIDOS"
    
    corte_id = request.args.get("corte_id", type=int)

    conn = get_db()
    cursor = conn.cursor()

    cortes_historial = cursor.execute("""
        SELECT id, tipo_estado, fecha_corte, gran_total, total_trabajadores, archivo_origen, creado_el
        FROM reportes_sgd_cortes
        WHERE tipo_estado = ?
        ORDER BY id DESC
    """, (tipo,)).fetchall()

    resultado = None
    corte_activo = None
    ahora = datetime.now()
    fecha_corte_defecto = ahora.strftime("%d/%m/%Y %H:%M")
    fecha_corte_input = ahora.strftime("%Y-%m-%dT%H:%M")

    if cortes_historial:
        if corte_id:
            corte_db = cursor.execute("SELECT * FROM reportes_sgd_cortes WHERE id = ? AND tipo_estado = ?", (corte_id, tipo)).fetchone()
        else:
            corte_db = cursor.execute("SELECT * FROM reportes_sgd_cortes WHERE id = ?", (cortes_historial[0]["id"],)).fetchone()

        if corte_db:
            corte_activo = corte_db
            try:
                resultado_original = json.loads(corte_db["datos_json"])
                ocultos = json.loads(corte_db["destinatarios_ocultos"] or "[]")
                resultado = aplicar_destinatarios_ocultos(resultado_original, ocultos)
                resultado["destinatarios_ocultos"] = ocultos
                fecha_corte_defecto = corte_db["fecha_corte"]
            except Exception:
                resultado = None

    conn.close()

    return render_template("sgd.html", 
                           tipo_activo=tipo,
                           cortes_historial=cortes_historial,
                           corte_activo=corte_activo,
                           resultado=resultado,
                           fecha_corte=fecha_corte_defecto,
                           fecha_corte_input=fecha_corte_input)

@app.route("/sgd/procesar/<tipo>", methods=["POST"])
@permiso_requerido("sgd", "modificar")
@login_required
def procesar_sgd(tipo):
    tipo_key = tipo.upper()
    if tipo_key not in ["NO_LEIDOS", "RECIBIDOS"]:
        tipo_key = "NO_LEIDOS"

    archivo = request.files.get("archivo_sgd")
    fecha_corte = datetime.now().strftime("%d/%m/%Y %H:%M")
    nombre_estado = "NO LEIDOS" if tipo_key == "NO_LEIDOS" else "RECIBIDOS"

    if not archivo or not archivo.filename:
        flash(f"Debe seleccionar el archivo Excel (.xlsx o .xls) de Documentos {nombre_estado}.", "warning")
        return redirect(f"/sgd?tipo={tipo_key}")

    ruta_guardado, _ = guardar_archivo_subido(archivo, SGD_FOLDER, {".xlsx"})

    resultado = procesar_sgd_excel(ruta_guardado, nombre_estado, fecha_corte)
    if not resultado or resultado["gran_total"] == 0:
        flash(f"No se encontraron registros válidos en el archivo de {nombre_estado}.", "warning")
        return redirect(f"/sgd?tipo={tipo_key}")

    # Guardar permanentemente en la base de datos
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO reportes_sgd_cortes (tipo_estado, fecha_corte, archivo_origen, gran_total, total_trabajadores, datos_json)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (tipo_key, fecha_corte, archivo.filename, resultado["gran_total"], resultado["total_trabajadores"], json.dumps(resultado)))
    nuevo_corte_id = cursor.lastrowid
    conn.commit()
    conn.close()

    flash(f"¡Corte del {fecha_corte} guardado permanentemente en el historial! ({resultado['gran_total']} documentos en {resultado['total_trabajadores']} trabajadores).", "success")
    return redirect(f"/sgd?tipo={tipo_key}&corte_id={nuevo_corte_id}")

@app.route("/sgd/ocultar-destinatario/<int:corte_id>", methods=["POST"])
@permiso_requerido("sgd", "modificar")
@login_required
def ocultar_destinatario_sgd(corte_id):
    destinatario = " ".join(request.form.get("destinatario", "").strip().upper().split())
    conn = get_db()
    corte = conn.execute("SELECT tipo_estado, destinatarios_ocultos FROM reportes_sgd_cortes WHERE id = ?", (corte_id,)).fetchone()
    if corte and destinatario:
        ocultos = json.loads(corte["destinatarios_ocultos"] or "[]")
        if request.form.get("accion") == "mostrar":
            ocultos = [nombre for nombre in ocultos if nombre != destinatario]
        elif destinatario not in ocultos:
            ocultos.append(destinatario)
        conn.execute("UPDATE reportes_sgd_cortes SET destinatarios_ocultos = ? WHERE id = ?", (json.dumps(sorted(ocultos)), corte_id))
        conn.commit()
    conn.close()
    tipo = corte["tipo_estado"] if corte else "NO_LEIDOS"
    return redirect(f"/sgd?tipo={tipo}&corte_id={corte_id}")

@app.route("/sgd/exportar-excel/<tipo>")
@permiso_requerido("sgd", "reportar")
@login_required
def exportar_sgd_excel(tipo):
    tipo_key = tipo.upper()
    if tipo_key not in ["NO_LEIDOS", "RECIBIDOS"]:
        tipo_key = "NO_LEIDOS"

    corte_id = request.args.get("corte_id", type=int)

    conn = get_db()
    cursor = conn.cursor()
    if corte_id:
        corte_db = cursor.execute("SELECT * FROM reportes_sgd_cortes WHERE id = ?", (corte_id,)).fetchone()
    else:
        corte_db = cursor.execute("SELECT * FROM reportes_sgd_cortes WHERE tipo_estado = ? ORDER BY id DESC LIMIT 1", (tipo_key,)).fetchone()
    conn.close()

    if not corte_db:
        flash(f"No hay reporte guardado de {tipo_key} para exportar.", "warning")
        return redirect(f"/sgd?tipo={tipo_key}")

    nombre_estado = "NO LEIDOS" if tipo_key == "NO_LEIDOS" else "RECIBIDOS"
    resultado = json.loads(corte_db["datos_json"])
    ocultos = json.loads(corte_db["destinatarios_ocultos"] or "[]")
    resultado = aplicar_destinatarios_ocultos(resultado, ocultos)
    fecha_corte_str = corte_db["fecha_corte"]
    buffer = generar_excel_sgd(resultado, nombre_estado, fecha_corte_str)
    nombre_descarga = f"Reporte_SGD_{tipo_key}_{corte_db['id']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/sgd/eliminar-corte/<int:corte_id>", methods=["POST"])
@permiso_requerido("sgd", "eliminar")
@login_required
def eliminar_corte_sgd(corte_id):
    conn = get_db()
    cursor = conn.cursor()
    corte = cursor.execute("SELECT tipo_estado, fecha_corte FROM reportes_sgd_cortes WHERE id = ?", (corte_id,)).fetchone()
    if corte:
        cursor.execute("DELETE FROM reportes_sgd_cortes WHERE id = ?", (corte_id,))
        conn.commit()
        flash(f"Corte del {corte['fecha_corte']} eliminado del historial.", "info")
    conn.close()
    tipo_ret = corte["tipo_estado"] if corte else "NO_LEIDOS"
    return redirect(f"/sgd?tipo={tipo_ret}")

# ================= MÓDULO: GESTIÓN DOCUMENTARIA =================
@app.route("/gestion-documentaria")
@permiso_requerido("documentaria", "ver")
@login_required
def gestion_documentaria():
    conn = get_db()
    cursor = conn.cursor()
    propietario_sql, propietario_params = filtro_propietario()
    documentos = cursor.execute(f"""
        SELECT * FROM documentos_registrados
        WHERE {propietario_sql}
        ORDER BY fecha_emision DESC, id DESC
    """, propietario_params).fetchall()
    # Siguiente numeración correlativa dinámica (si se elimina un registro, regresa automáticamente al anterior)
    siguiente_carta = cursor.execute("""
        SELECT COALESCE(MAX(numero_ordinal), 8) + 1 FROM documentos_registrados WHERE tipo_documento = 'CARTA'
    """).fetchone()[0]
    siguiente_informe = cursor.execute("""
        SELECT COALESCE(MAX(numero_ordinal), 8) + 1 FROM documentos_registrados WHERE tipo_documento = 'INFORME'
    """).fetchone()[0]

    conn.close()

    pattern_config = get_document_number_pattern_config()
    drive_configured = bool(os.getenv("GOOGLE_DRIVE_CLIENT_ID") and os.getenv("GOOGLE_DRIVE_CLIENT_SECRET"))

    return render_template(
        "gestion_documentaria.html",
        documentos=documentos,
        gemini_configurada=bool(os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")),
        drive_oauth_configurada=drive_configured,
        document_number_config=pattern_config,
        siguiente_carta=siguiente_carta,
        siguiente_informe=siguiente_informe,
        fecha_hoy=date.today().strftime("%Y-%m-%d")
    )


@app.route("/gestion-documentaria/eliminar/<int:doc_id>", methods=["POST"])
@permiso_requerido("documentaria", "eliminar")
@login_required
def eliminar_documento(doc_id):
    conn = get_db()
    cursor = conn.cursor()
    propietario_sql, propietario_params = filtro_propietario()
    doc = cursor.execute(f"SELECT * FROM documentos_registrados WHERE id = ? AND {propietario_sql}", (doc_id, *propietario_params)).fetchone()
    if doc:
        # Intentar eliminar archivo físico si existe en actas_generadas
        if doc["archivo_ruta"] and os.path.exists(doc["archivo_ruta"]) and "actas_generadas" in doc["archivo_ruta"]:
            try:
                os.remove(doc["archivo_ruta"])
            except Exception:
                pass
        cursor.execute("DELETE FROM documentos_registrados WHERE id = ?", (doc_id,))
        conn.commit()
        flash(f"Documento '{doc['numero'] or doc['titulo']}' eliminado correctamente.", "info")
    conn.close()
    return redirect(url_for("gestion_documentaria"))


@app.route("/gestion-documentaria/generar-informe-docx", methods=["POST"])
@permiso_requerido("documentaria", "modificar")
@login_required
def generar_informe_tdr_docx():
    num_carta = request.form.get("numero_carta", type=int) or 9
    num_informe = request.form.get("numero_informe", type=int) or 9
    fecha_emision = request.form.get("fecha_emision", date.today().strftime("%Y-%m-%d"))
    num_os = request.form.get("numero_os", "0005612").strip()
    periodo_desde = request.form.get("periodo_desde", "")
    periodo_hasta = request.form.get("periodo_hasta", "")
    guardar_en_bd = request.form.get("guardar_en_bd") == "1"

    anio = datetime.now().year
    if fecha_emision and len(fecha_emision) >= 4:
        try:
            anio = int(fecha_emision[:4])
        except Exception:
            pass

    ruta_plantilla = os.path.join("plantillas", "modelo.docx")
    if not os.path.exists(ruta_plantilla):
        ruta_plantilla = os.path.join("plantillas", "ejemplo_final.docx")

    nombre_archivo = f"CARTA_{num_carta:03d}_E_INFORME_{num_informe:03d}_{anio}_VCEM_SI.docx"
    ruta_salida = os.path.join(ACTAS_FOLDER, nombre_archivo)

    conn = get_db()
    cursor = conn.cursor()
    propietario_sql, propietario_params = filtro_propietario()

    # Extraer actividades TI del periodo
    query_act = f"SELECT * FROM actividades_soporte WHERE {propietario_sql}"
    params_act = list(propietario_params)
    if periodo_desde:
        query_act += " AND fecha_programada >= ?"
        params_act.append(periodo_desde)
    if periodo_hasta:
        query_act += " AND fecha_programada <= ?"
        params_act.append(periodo_hasta)
    query_act += " ORDER BY categoria ASC, fecha_programada ASC"

    actividades_db = cursor.execute(query_act, params_act).fetchall()
    actividades_lista = [dict(a) for a in actividades_db]

    try:
        generar_informe_mensual_docx(
            ruta_plantilla=ruta_plantilla,
            num_carta=num_carta,
            num_informe=num_informe,
            fecha_doc_str=fecha_emision,
            num_os=num_os,
            anio=anio,
            actividades_lista=actividades_lista,
            ruta_salida=ruta_salida
        )

        if guardar_en_bd:
            conn = get_db()
            cursor = conn.cursor()
            user_id = session.get("usuario_id")
            
            # Registrar Carta
            num_carta_str = f"CARTA N° {num_carta:03d}-{anio}-VCEM-SI"
            cursor.execute("""
                INSERT INTO documentos_registrados (
                    tipo_documento, numero, numero_ordinal, titulo, descripcion,
                    destinatario, remitente, fecha_emision, requerimiento, uso_documento,
                    archivo_nombre, archivo_ruta, fuente, usuario_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'GENERADO_AUTO', ?)
            """, (
                "CARTA", num_carta_str, num_carta,
                f"Remite Informe de Actividades N° {num_informe:03d}-{anio}-VCEM-SI",
                f"Carta de elevación de informe mensual de actividades TDR correspondiente a la OS N° {num_os}.",
                "LIC. MIGUEL ANGEL CHAMBI CCALLA - JEFE UT APURIMAC",
                "ING. VLADIMIR CLAUDIO ESCOBEDO MIRANDA",
                fecha_emision,
                f"Conformidad de servicio OS N° {num_os}",
                "Trámite de pago de orden de servicio mensual",
                nombre_archivo, ruta_salida, user_id
            ))

            # Registrar Informe
            num_informe_str = f"INFORME N° {num_informe:03d}-{anio}-VCEM-SI"
            cursor.execute("""
                INSERT INTO documentos_registrados (
                    tipo_documento, numero, numero_ordinal, titulo, descripcion,
                    destinatario, remitente, fecha_emision, requerimiento, uso_documento,
                    archivo_nombre, archivo_ruta, fuente, usuario_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'GENERADO_AUTO', ?)
            """, (
                "INFORME", num_informe_str, num_informe,
                f"Informe de Actividades Realizadas como Soporte Informático UTI - Periodo {periodo_desde} al {periodo_hasta}",
                f"Informe mensual de soporte técnico, mantenimiento, inventario y gestión documentaria de la UT Apurímac.",
                "LIC. MIGUEL ANGEL CHAMBI CCALLA - JEFE UT APURIMAC",
                "ING. VLADIMIR CLAUDIO ESCOBEDO MIRANDA",
                fecha_emision,
                f"Cumplimiento de Términos de Referencia OS N° {num_os}",
                "Sustento de actividades realizadas para pago mensual",
                nombre_archivo, ruta_salida, user_id
            ))

            conn.commit()

        conn.close()

        return send_file(
            ruta_salida,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        conn.close()
        flash(f"Error al generar el documento Word: {e}", "danger")
        return redirect(url_for("gestion_documentaria"))


@app.route("/gestion-documentaria/configuracion", methods=["POST"])
@permiso_requerido("documentaria", "modificar")
@login_required
def guardar_configuracion_documentaria():
    patron = request.form.get("patron", "{tipo} N° {numero:03d}-{year}-{sufijo}").strip()
    year = request.form.get("year", datetime.now().year)
    sufijo = request.form.get("sufijo", "VCEM-SI").strip()

    if not patron:
        patron = "{tipo} N° {numero:03d}-{year}-{sufijo}"
    if not sufijo:
        sufijo = "VCEM-SI"

    try:
        year_int = int(year)
    except ValueError:
        year_int = datetime.now().year

    guardar_document_number_pattern_config({
        "patron": patron,
        "year": year_int,
        "sufijo": sufijo
    })

    flash("Se actualizó el patrón de numeración documental correctamente.", "success")
    return redirect(url_for("gestion_documentaria"))


@app.route("/gestion-documentaria/drive/oauth")
@login_required
def iniciar_drive_oauth():
    client_id = os.getenv("GOOGLE_DRIVE_CLIENT_ID")
    client_secret = os.getenv("GOOGLE_DRIVE_CLIENT_SECRET")
    redirect_uri = os.getenv("GOOGLE_DRIVE_REDIRECT_URI", "http://127.0.0.1:5000/gestion-documentaria/drive/oauth/callback")

    if not client_id or not client_secret:
        flash("Falta configurar GOOGLE_DRIVE_CLIENT_ID y GOOGLE_DRIVE_CLIENT_SECRET para habilitar OAuth con Google Drive.", "warning")
        return redirect(url_for("gestion_documentaria"))

    client_config = {
        "web": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "redirect_uris": [redirect_uri]
        }
    }

    flow = Flow.from_client_config(client_config, scopes=["https://www.googleapis.com/auth/drive.file"])
    flow.redirect_uri = redirect_uri
    authorization_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent"
    )
    session["drive_oauth_state"] = state
    return redirect(authorization_url)


@app.route("/gestion-documentaria/drive/oauth/callback")
@login_required
def callback_drive_oauth():
    client_id = os.getenv("GOOGLE_DRIVE_CLIENT_ID")
    client_secret = os.getenv("GOOGLE_DRIVE_CLIENT_SECRET")
    redirect_uri = os.getenv("GOOGLE_DRIVE_REDIRECT_URI", "http://127.0.0.1:5000/gestion-documentaria/drive/oauth/callback")

    if not client_id or not client_secret:
        flash("La configuración de Google Drive no está disponible.", "warning")
        return redirect(url_for("gestion_documentaria"))

    client_config = {
        "web": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "redirect_uris": [redirect_uri]
        }
    }

    state = request.args.get("state")
    if not state or session.get("drive_oauth_state") != state:
        flash("No se pudo validar la sesión de autenticación de Google Drive.", "danger")
        return redirect(url_for("gestion_documentaria"))

    flow = Flow.from_client_config(client_config, scopes=["https://www.googleapis.com/auth/drive.file"], state=state)
    flow.redirect_uri = redirect_uri
    flow.fetch_token(authorization_response=request.url)

    credentials = flow.credentials
    session["google_drive_credentials"] = {
        "token": credentials.token,
        "refresh_token": credentials.refresh_token,
        "token_uri": credentials.token_uri,
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
        "scopes": credentials.scopes
    }
    session.pop("drive_oauth_state", None)
    flash("Conexión con Google Drive autenticada correctamente.", "success")
    return redirect(url_for("gestion_documentaria"))


@app.route("/gestion-documentaria/exportar-pdf")
@permiso_requerido("documentaria", "reportar")
@login_required
def exportar_documentos_pdf():
    conn = get_db()
    propietario_sql, propietario_params = filtro_propietario()
    documentos = conn.cursor().execute(f"SELECT * FROM documentos_registrados WHERE {propietario_sql} ORDER BY fecha_emision DESC, id DESC", propietario_params).fetchall()
    conn.close()

    buffer = BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(buffer, pagesize=A4, title="Historial de documentos")
    story = [
        Paragraph("Historial de documentos", styles['Title']),
        Spacer(1, 18),
        Paragraph("Sistema de gestión documental", styles['Normal'])
    ]

    data = [["Tipo", "Número", "Título", "Fecha", "Uso", "Requerimiento"]]
    for d in documentos:
        data.append([
            d["tipo_documento"],
            d["numero"] or "SIN NUM.",
            d["titulo"] or "-",
            d["fecha_emision"] or "-",
            d["uso_documento"] or "-",
            d["requerimiento"] or "-",
        ])

    table = Table(data, colWidths=[42, 120, 160, 60, 120, 150])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.8, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.Color(1, 1, 1)]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(table)
    doc.build(story)

    buffer.seek(0)
    nombre_descarga = f"Documentos_Gestion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/pdf")


@app.route("/gestion-documentaria/guardar", methods=["POST"])
@permiso_requerido("documentaria", "modificar")
@login_required
def guardar_documento():
    tipo_documento = request.form.get("tipo_documento", "CARTA").strip().upper()
    titulo = request.form.get("titulo", "").strip()
    descripcion = request.form.get("descripcion", "").strip()
    remitente = request.form.get("remitente", "").strip()
    destinatario = request.form.get("destinatario", "").strip()
    fecha_emision = request.form.get("fecha_emision") or datetime.now().strftime("%Y-%m-%d")
    requerimiento = request.form.get("requerimiento", "").strip()
    uso_documento = request.form.get("uso_documento", "").strip()
    texto_manual = request.form.get("contenido_texto", "").strip()

    if not titulo:
        flash("Debe ingresar el título del documento.", "warning")
        return redirect(url_for("gestion_documentaria"))

    conn = get_db()
    cursor = conn.cursor()
    numero = request.form.get("numero", "").strip()
    if not numero:
        numero = generar_numero_documento(tipo_documento, conn)
    numero_ordinal = 0
    if numero:
        match = re.search(r"(\d+)", numero)
        if match:
            numero_ordinal = int(match.group(1))

    archivo = request.files.get("archivo_documento")
    archivo_nombre = None
    archivo_ruta = None
    if archivo and archivo.filename:
        archivo_ruta, archivo_nombre = guardar_archivo_subido(
            archivo, DOCUMENTOS_FOLDER, {".pdf", ".doc", ".docx", ".txt", ".md", ".csv", ".xlsx"}
        )

    resumen = ""
    if texto_manual or archivo_ruta:
        analisis = analizar_con_gemini(texto_manual, archivo_ruta)
        resumen = analisis.get("resumen", "")

    cursor.execute("""
        INSERT INTO documentos_registrados (
            tipo_documento, numero, numero_ordinal, titulo, descripcion, remitente, destinatario,
            fecha_emision, requerimiento, uso_documento, resumen_gemini, archivo_nombre, archivo_ruta,
            fuente, usuario_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        tipo_documento,
        numero,
        numero_ordinal,
        titulo,
        descripcion,
        remitente,
        destinatario,
        fecha_emision,
        requerimiento,
        uso_documento,
        resumen,
        archivo_nombre,
        archivo_ruta,
        "MANUAL",
        session.get("usuario_id")
    ))
    conn.commit()
    conn.close()

    flash(f"Documento registrado correctamente con numeración: {numero or titulo}.", "success")
    return redirect(url_for("gestion_documentaria"))

@app.route("/gestion-documentaria/drive", methods=["POST"])
@permiso_requerido("documentaria", "modificar")
@login_required
def guardar_documento_drive():
    tipo_documento = request.form.get("tipo_documento", "MEMORANDO").strip().upper()
    titulo = request.form.get("titulo", "").strip()
    requisito = request.form.get("requerimiento", "").strip()
    uso_documento = request.form.get("uso_documento", "").strip()
    contenido_texto = request.form.get("contenido_texto", "").strip()
    archivo = request.files.get("archivo_drive")

    if not titulo:
        flash("Debe indicar el título del memorando o indicación.", "warning")
        return redirect(url_for("gestion_documentaria"))

    conn = get_db()
    cursor = conn.cursor()
    numero = generar_numero_documento(tipo_documento, conn)
    numero_ordinal = 0
    if numero:
        match = re.search(r"(\d+)", numero)
        if match:
            numero_ordinal = int(match.group(1))
    archivo_nombre = None
    archivo_ruta = None
    drive_url = ""
    drive_file_id = None

    if archivo and archivo.filename:
        archivo_ruta, archivo_nombre = guardar_archivo_subido(
            archivo, DOCUMENTOS_FOLDER, {".pdf", ".doc", ".docx", ".txt", ".md", ".csv", ".xlsx"}
        )

        if session.get("google_drive_credentials") and os.getenv("GOOGLE_DRIVE_CLIENT_ID") and os.getenv("GOOGLE_DRIVE_CLIENT_SECRET"):
            try:
                cred_data = session.get("google_drive_credentials", {})
                credentials = Credentials(
                    token=cred_data.get("token"),
                    refresh_token=cred_data.get("refresh_token"),
                    token_uri=cred_data.get("token_uri"),
                    client_id=cred_data.get("client_id"),
                    client_secret=cred_data.get("client_secret"),
                    scopes=cred_data.get("scopes")
                )
                service = build("drive", "v3", credentials=credentials, cache_discovery=False)
                file_metadata = {"name": nombre_archivo, "description": f"Documento cargado desde SISGEMOSI: {titulo}"}
                media = MediaIoBaseUpload(BytesIO(archivo.read()), mimetype=archivo.mimetype or "application/octet-stream", resumable=True)
                uploaded = service.files().create(body=file_metadata, media_body=media, fields="id,webViewLink,webContentLink").execute()
                drive_url = uploaded.get("webViewLink") or uploaded.get("webContentLink") or ""
                drive_file_id = uploaded.get("id")
            except Exception as exc:
                flash(f"El archivo se guardó localmente, pero no pudo subirse a Google Drive: {exc}", "warning")

    resumen = ""
    if contenido_texto or archivo_ruta:
        analisis = analizar_con_gemini(contenido_texto, archivo_ruta)
        resumen = analisis.get("resumen", "")

    cursor.execute("""
        INSERT INTO documentos_registrados (
            tipo_documento, numero, numero_ordinal, titulo, descripcion, remitente, destinatario,
            fecha_emision, requerimiento, uso_documento, resumen_gemini, archivo_nombre, archivo_ruta,
            drive_url, drive_file_id, fuente, usuario_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        tipo_documento,
        numero,
        numero_ordinal,
        titulo,
        contenido_texto or requisito,
        "",
        "",
        datetime.now().strftime("%Y-%m-%d"),
        requisito,
        uso_documento,
        resumen,
        archivo_nombre,
        archivo_ruta,
        drive_url,
        drive_file_id,
        "DRIVE",
        session.get("usuario_id")
    ))
    conn.commit()
    conn.close()

    if not (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")):
        flash("El archivo quedó registrado localmente. Para activar el análisis con Gemini, añade la variable GEMINI_API_KEY al entorno del proyecto.", "info")
    else:
        flash("Documento cargado, analizado con Gemini y sincronizado con Drive cuando fue posible.", "success")
    return redirect(url_for("gestion_documentaria"))

@app.route("/gestion-documentaria/exportar-excel")
@permiso_requerido("documentaria", "reportar")
@login_required
def exportar_documentos_excel():
    conn = get_db()
    propietario_sql, propietario_params = filtro_propietario()
    documentos = conn.cursor().execute(f"SELECT * FROM documentos_registrados WHERE {propietario_sql} ORDER BY fecha_emision DESC, id DESC", propietario_params).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Documentos"
    headers = ["ID", "Tipo", "Número", "Título", "Remitente", "Destinatario", "Fecha", "Requerimiento", "Uso", "Resumen Gemini", "Archivo", "Fuente"]
    ws.append(headers)

    for d in documentos:
        ws.append([
            d["id"],
            d["tipo_documento"],
            d["numero"],
            d["titulo"],
            d["remitente"],
            d["destinatario"],
            d["fecha_emision"],
            d["requerimiento"],
            d["uso_documento"],
            d["resumen_gemini"],
            d["archivo_nombre"],
            d["fuente"]
        ])

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for column_cells in ws.columns:
        max_length = 0
        for cell in column_cells:
            try:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[column_cells[0].column_letter].width = min(max_length + 3, 35)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    nombre_descarga = f"Documentos_Gestion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ================= MÓDULO: REGISTRO DE ACTIVIDADES (SOPORTE TI) =================
@app.route("/actividades")
@permiso_requerido("actividades", "ver")
@login_required
def ver_actividades():
    filtro_estado = request.args.get("estado", "TODOS")
    filtro_cat = request.args.get("categoria", "TODAS")
    filtro_periodo = request.args.get("periodo", "TODOS")
    conn = get_db()
    cursor = conn.cursor()

    propietario_sql, propietario_params = filtro_propietario()
    base_query = f"SELECT * FROM actividades_soporte WHERE {propietario_sql}"
    params = list(propietario_params)

    if filtro_estado != "TODOS":
        base_query += " AND estado = ?"
        params.append(filtro_estado)
    if filtro_cat != "TODAS":
        base_query += " AND categoria = ?"
        params.append(filtro_cat)

    hoy_dt = date.today()
    if filtro_periodo == "HOY":
        base_query += " AND fecha_programada = ?"
        params.append(hoy_dt.isoformat())
    elif filtro_periodo == "SEMANA":
        inicio_semana = hoy_dt - timedelta(days=hoy_dt.weekday())
        fin_semana = inicio_semana + timedelta(days=6)
        base_query += " AND fecha_programada BETWEEN ? AND ?"
        params.extend([inicio_semana.isoformat(), fin_semana.isoformat()])
    elif filtro_periodo == "MES":
        inicio_mes = hoy_dt.replace(day=1)
        siguiente_mes = (inicio_mes.replace(day=28) + timedelta(days=4)).replace(day=1)
        fin_mes = siguiente_mes - timedelta(days=1)
        base_query += " AND fecha_programada BETWEEN ? AND ?"
        params.extend([inicio_mes.isoformat(), fin_mes.isoformat()])

    base_query += " ORDER BY CASE estado WHEN 'PENDIENTE' THEN 0 WHEN 'EN PROGRESO' THEN 1 WHEN 'COMPLETADA' THEN 2 ELSE 3 END, fecha_programada DESC, hora_inicio ASC, id DESC"
    actividades_raw = cursor.execute(base_query, params).fetchall()

    actividades = []
    for a in actividades_raw:
        item = dict(a)
        f_prog = item["fecha_programada"] or datetime.now().strftime("%Y-%m-%d")
        h_ini = (item["hora_inicio"] or "09:00").replace(":", "")
        h_fin = (item["hora_fin"] or "10:00").replace(":", "")
        f_clean = str(f_prog).replace("-", "")
        
        start_iso = f"{f_clean}T{h_ini}00"
        end_iso = f"{f_clean}T{h_fin}00"
        
        titulo_enc = urllib.parse.quote(f"Soporte TI: {item['titulo']}")
        detalles_enc = urllib.parse.quote(f"Área: {item['area_solicitante'] or 'UTI'}\\nCategoría: {item['categoria']}\\nPrioridad: {item['prioridad']}\\n\\nDescripción:\\n{item['descripcion'] or ''}")
        
        item["google_calendar_url"] = f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={titulo_enc}&dates={start_iso}/{end_iso}&details={detalles_enc}&location=UTI%20-%20UT%20Apurimac"
        actividades.append(item)

    trabajadores = cursor.execute("""
        SELECT id, dni, (nombres || ' ' || apellido_paterno || ' ' || apellido_materno) AS nombre_completo, cargo 
        FROM trabajadores 
        WHERE activo = 1
        ORDER BY nombres ASC
    """).fetchall()

    total_pendientes = cursor.execute(f"SELECT COUNT(*) FROM actividades_soporte WHERE estado = 'PENDIENTE' AND {propietario_sql}", propietario_params).fetchone()[0]
    total_proceso = cursor.execute(f"SELECT COUNT(*) FROM actividades_soporte WHERE estado = 'EN PROGRESO' AND {propietario_sql}", propietario_params).fetchone()[0]
    total_completadas = cursor.execute(f"SELECT COUNT(*) FROM actividades_soporte WHERE estado = 'COMPLETADA' AND {propietario_sql}", propietario_params).fetchone()[0]
    total_general = cursor.execute(f"SELECT COUNT(*) FROM actividades_soporte WHERE {propietario_sql}", propietario_params).fetchone()[0]
    
    ahora_dt = datetime.now()
    hora_actual = ahora_dt.strftime("%H:%M")
    hora_fin_estimada = (ahora_dt + timedelta(hours=1)).strftime("%H:%M")
    conn.close()

    return render_template("actividades.html",
                           actividades=actividades,
                           trabajadores=trabajadores,
                           filtro_estado=filtro_estado,
                           filtro_cat=filtro_cat,
                           filtro_periodo=filtro_periodo,
                           total_pendientes=total_pendientes,
                           total_proceso=total_proceso,
                           total_completadas=total_completadas,
                           total_general=total_general,
                           hoy=datetime.now().strftime("%Y-%m-%d"),
                           hora_actual=hora_actual,
                           hora_fin_estimada=hora_fin_estimada)

@app.route("/actividades/nueva", methods=["POST"])
@permiso_requerido("actividades", "modificar")
@login_required
def nueva_actividad():
    titulo = request.form["titulo"].strip()
    descripcion = request.form.get("descripcion", "").strip()
    categoria = request.form.get("categoria", "SOPORTE TECNICO").strip()
    area = request.form.get("area_solicitante", "").strip()
    prioridad = request.form.get("prioridad", "MEDIA").strip()
    estado = request.form.get("estado", "PENDIENTE").strip()
    fecha_prog = request.form.get("fecha_programada") or datetime.now().strftime("%Y-%m-%d")
    h_inicio = request.form.get("hora_inicio", "").strip()
    h_fin = request.form.get("hora_fin", "").strip()

    if not h_inicio:
        h_inicio = datetime.now().strftime("%H:%M")
    if not h_fin:
        h_fin = (datetime.now() + timedelta(hours=1)).strftime("%H:%M")

    conn = get_db()
    conn.cursor().execute("""
        INSERT INTO actividades_soporte (titulo, descripcion, categoria, area_solicitante, prioridad, estado, fecha_programada, hora_inicio, hora_fin, usuario_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (titulo, descripcion, categoria, area, prioridad, estado, fecha_prog, h_inicio, h_fin, session.get("usuario_id")))
    conn.commit()
    conn.close()

    flash("Actividad registrada exitosamente.", "success")
    return redirect(url_for("ver_actividades"))

@app.route("/actividades/editar", methods=["POST"])
@permiso_requerido("actividades", "modificar")
@login_required
def editar_actividad():
    actividad_id = request.form["actividad_id"]
    titulo = request.form["titulo"].strip()
    descripcion = request.form.get("descripcion", "").strip()
    categoria = request.form.get("categoria", "SOPORTE TECNICO").strip()
    area = request.form.get("area_solicitante", "").strip()
    prioridad = request.form.get("prioridad", "MEDIA").strip()
    estado = request.form.get("estado", "PENDIENTE").strip()
    fecha_prog = request.form.get("fecha_programada") or datetime.now().strftime("%Y-%m-%d")
    h_inicio = request.form.get("hora_inicio", "").strip()
    h_fin = request.form.get("hora_fin", "").strip()
    solucion = request.form.get("solucion_aplicada", "").strip()

    conn = get_db()
    propietario_sql, propietario_params = filtro_propietario()
    conn.cursor().execute(f"""
        UPDATE actividades_soporte 
        SET titulo = ?, descripcion = ?, categoria = ?, area_solicitante = ?, prioridad = ?,
            estado = ?, fecha_programada = ?, hora_inicio = ?, hora_fin = ?, solucion_aplicada = ?
        WHERE id = ? AND {propietario_sql}
    """, (titulo, descripcion, categoria, area, prioridad, estado, fecha_prog, h_inicio, h_fin, solucion, actividad_id, *propietario_params))
    conn.commit()
    conn.close()

    flash("Actividad actualizada correctamente.", "success")
    return redirect(url_for("ver_actividades"))

@app.route("/actividades/estado/<int:actividad_id>/<nuevo_estado>", methods=["POST"])
@permiso_requerido("actividades", "modificar")
@login_required
def cambiar_estado_actividad(actividad_id, nuevo_estado):
    if nuevo_estado in ["PENDIENTE", "EN PROGRESO", "COMPLETADA"]:
        conn = get_db()
        propietario_sql, propietario_params = filtro_propietario()
        conn.cursor().execute(f"UPDATE actividades_soporte SET estado = ? WHERE id = ? AND {propietario_sql}", (nuevo_estado, actividad_id, *propietario_params))
        conn.commit()
        conn.close()
        flash(f"Estado de la actividad cambiado a {nuevo_estado}.", "info")
    return redirect(request.referrer or url_for("ver_actividades"))

@app.route("/actividades/acciones-realizadas", methods=["POST"])
@permiso_requerido("actividades", "modificar")
@login_required
def guardar_acciones_realizadas():
    actividad_id = request.form["actividad_id"]
    acciones = request.form.get("solucion_aplicada", "").strip()

    if not acciones:
        flash("Debe registrar las acciones realizadas antes de completar la actividad.", "warning")
        return redirect(url_for("ver_actividades"))

    conn = get_db()
    propietario_sql, propietario_params = filtro_propietario()
    conn.cursor().execute(f"""
        UPDATE actividades_soporte
        SET estado = 'COMPLETADA', solucion_aplicada = ?
        WHERE id = ? AND {propietario_sql}
    """, (acciones, actividad_id, *propietario_params))
    conn.commit()
    conn.close()

    flash("Acciones realizadas guardadas correctamente.", "success")
    return redirect(url_for("ver_actividades"))

@app.route("/actividades/eliminar/<int:actividad_id>", methods=["POST"])
@permiso_requerido("actividades", "eliminar")
@login_required
def eliminar_actividad(actividad_id):
    conn = get_db()
    propietario_sql, propietario_params = filtro_propietario()
    conn.cursor().execute(f"DELETE FROM actividades_soporte WHERE id = ? AND {propietario_sql}", (actividad_id, *propietario_params))
    conn.commit()
    conn.close()
    flash("Actividad eliminada de la bitácora.", "success")
    return redirect(url_for("ver_actividades"))

@app.route("/api/actividades-calendario")
@login_required
def api_actividades_calendario():
    conn = get_db()
    propietario_sql, propietario_params = filtro_propietario()
    filas = conn.cursor().execute(f"SELECT * FROM actividades_soporte WHERE {propietario_sql}", propietario_params).fetchall()
    conn.close()

    colores = {
        "PENDIENTE": "#dc3545",
        "EN PROGRESO": "#ffc107",
        "COMPLETADA": "#198754"
    }

    eventos = []
    for f in filas:
        f_prog = f["fecha_programada"] or datetime.now().strftime("%Y-%m-%d")
        h_ini = f["hora_inicio"] or "08:00"
        h_fin = f["hora_fin"] or "09:00"
        
        eventos.append({
            "id": f["id"],
            "title": f"[{f['categoria']}] {f['titulo']}",
            "start": f"{f_prog}T{h_ini}:00",
            "end": f"{f_prog}T{h_fin}:00",
            "backgroundColor": colores.get(f["estado"], "#0d6efd"),
            "borderColor": colores.get(f["estado"], "#0d6efd"),
            "description": f["descripcion"] or "",
            "area": f["area_solicitante"] or "",
            "estado": f["estado"],
            "prioridad": f["prioridad"]
        })
    return jsonify(eventos)

@app.route("/actividades/exportar-excel")
@permiso_requerido("actividades", "reportar")
@login_required
def exportar_actividades_excel():
    conn = get_db()
    cursor = conn.cursor()
    propietario_sql, propietario_params = filtro_propietario()
    filas = cursor.execute(f"""
        SELECT fecha_programada, hora_inicio, hora_fin, area_solicitante, categoria, 
               titulo, descripcion, prioridad, estado, solucion_aplicada
        FROM actividades_soporte
        WHERE {propietario_sql}
        ORDER BY fecha_programada DESC, hora_inicio ASC
    """, propietario_params).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bitácora de Soporte TI"

    headers = [
        "FECHA", "HORA INICIO", "HORA FIN", "ÁREA SOLICITANTE", "CATEGORÍA",
        "ACTIVIDAD / TAREA", "DETALLE / DESCRIPCIÓN", "PRIORIDAD", "ESTADO", "SOLUCIÓN APLICADA / REPORTE"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions.height = 28

    cols_centradas = (1, 2, 3, 8, 9)
    for r_idx, f in enumerate(filas, start=2):
        row_data = [
            sanitizar_texto_excel(formatear_solo_fecha(f["fecha_programada"])),
            sanitizar_texto_excel(f["hora_inicio"]),
            sanitizar_texto_excel(f["hora_fin"]),
            sanitizar_texto_excel(f["area_solicitante"]),
            sanitizar_texto_excel(f["categoria"]),
            sanitizar_texto_excel(f["titulo"]),
            sanitizar_texto_excel(f["descripcion"]),
            sanitizar_texto_excel(f["prioridad"]),
            sanitizar_texto_excel(f["estado"]),
            sanitizar_texto_excel(f["solucion_aplicada"])
        ]
        ws.append(row_data)
        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=r_idx, column=col_num)
            c.border = thin_border
            if col_num in cols_centradas:
                c.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    nombre_descarga = f"Bitacora_Actividades_TI_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/actividades/exportar-docx", methods=["GET", "POST"])
@permiso_requerido("actividades", "reportar")
@login_required
def exportar_actividades_docx():
    periodo_desde = request.values.get("periodo_desde", "").strip()
    periodo_hasta = request.values.get("periodo_hasta", "").strip()
    categoria = request.values.get("categoria", "").strip()
    estado = request.values.get("estado", "").strip()

    conn = get_db()
    cursor = conn.cursor()
    propietario_sql, propietario_params = filtro_propietario()

    query = f"SELECT * FROM actividades_soporte WHERE {propietario_sql}"
    params = list(propietario_params)

    if periodo_desde:
        query += " AND fecha_programada >= ?"
        params.append(periodo_desde)
    if periodo_hasta:
        query += " AND fecha_programada <= ?"
        params.append(periodo_hasta)
    if categoria and categoria != "TODAS":
        query += " AND categoria = ?"
        params.append(categoria)
    if estado and estado != "TODOS":
        query += " AND estado = ?"
        params.append(estado)

    query += " ORDER BY categoria ASC, fecha_programada ASC, hora_inicio ASC"
    filas = cursor.execute(query, params).fetchall()
    conn.close()

    actividades_lista = [dict(f) for f in filas]

    nombre_descarga = f"Reporte_Actividades_TI_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    ruta_salida = os.path.join(ACTAS_FOLDER, nombre_descarga)

    try:
        generar_reporte_actividades_docx(
            actividades_lista=actividades_lista,
            periodo_desde=periodo_desde,
            periodo_hasta=periodo_hasta,
            categoria_filtro=categoria,
            ruta_salida=ruta_salida
        )

        return send_file(
            ruta_salida,
            as_attachment=True,
            download_name=nombre_descarga,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        flash(f"Error al generar el reporte Word de actividades: {e}", "danger")
        return redirect(url_for("ver_actividades"))

# ================= HUB DE PRÉSTAMOS =================
@app.route("/prestamos")
@permiso_requerido("prestamos", "ver")
@login_required
def prestamos_hub():
    conn = get_db()
    cursor = conn.cursor()
    
    moviles = cursor.execute("SELECT COUNT(*) AS total, SUM(CASE WHEN estado='EN PRESTAMO' THEN 1 ELSE 0 END) AS prestados FROM equipos WHERE categoria='MOVILES'").fetchone()
    tokens = cursor.execute("SELECT COUNT(*) AS total, SUM(CASE WHEN estado='EN PRESTAMO' THEN 1 ELSE 0 END) AS prestados FROM equipos WHERE categoria='TOKENS'").fetchone()
    informaticos = cursor.execute("SELECT COUNT(*) AS total, SUM(CASE WHEN estado='EN PRESTAMO' THEN 1 ELSE 0 END) AS prestados FROM equipos WHERE categoria='INFORMATICOS'").fetchone()
    mobiliario = cursor.execute("SELECT COUNT(*) AS total, SUM(CASE WHEN estado='EN PRESTAMO' THEN 1 ELSE 0 END) AS prestados FROM equipos WHERE categoria='MOBILIARIO'").fetchone()
    conn.close()

    return render_template("prestamos_hub.html", moviles=moviles, tokens=tokens, informaticos=informaticos, mobiliario=mobiliario)

# ================= VISTA DE CADA SUBMÓDULO DE PRÉSTAMO =================
@app.route("/prestamos/<categoria>")
@permiso_requerido("prestamos", "ver")
@login_required
def ver_submodulo(categoria):
    cat = categoria.upper()
    if cat not in ["MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"]:
        return redirect(url_for("prestamos_hub"))

    conn = get_db()
    cursor = conn.cursor()

    query = """
    SELECT 
        e.*,
        p.id AS prestamo_id,
        p.fecha_prestamo,
        p.archivo_acta_entrega,
        t.dni AS trabajador_dni,
        (t.nombres || ' ' || t.apellido_paterno || ' ' || t.apellido_materno) AS trabajador_nombre,
        t.cargo AS trabajador_cargo,
        t.telefono AS trabajador_telefono
    FROM equipos e
    LEFT JOIN prestamos p ON e.id = p.equipo_id AND p.estado = 'ACTIVO'
    LEFT JOIN trabajadores t ON p.trabajador_id = t.id
    WHERE e.categoria = ?
    ORDER BY CASE WHEN e.estado = 'EN PRESTAMO' THEN 0 ELSE 1 END, p.fecha_prestamo DESC, e.codigo_margesi ASC
    """
    equipos = cursor.execute(query, (cat,)).fetchall()
    trabajadores = cursor.execute("""
        SELECT id, dni, (nombres || ' ' || apellido_paterno || ' ' || apellido_materno) AS nombre_completo, cargo 
        FROM trabajadores 
        WHERE activo = 1
        ORDER BY nombres ASC
    """).fetchall()

    tipos_raw = cursor.execute("SELECT DISTINCT descripcion FROM equipos WHERE categoria = ? ORDER BY descripcion ASC", (cat,)).fetchall()
    tipos_disponibles = [r[0] for r in tipos_raw if r[0]]

    total_cat = len(equipos)
    prestados_cat = sum(1 for eq in equipos if eq["estado"] == "EN PRESTAMO")
    almacen_cat = total_cat - prestados_cat
    conn.close()

    nombres_titulos = {
        "MOVILES": "Submódulo de Equipos Móviles (Celulares / Tablets)",
        "TOKENS": "Submódulo de Tokens Digitales / Firmas Criptográficas",
        "INFORMATICOS": "Submódulo de Equipos Informáticos y Audiovisuales",
        "MOBILIARIO": "Submódulo de Mobiliario y Bienes Patrimoniales"
    }

    return render_template("dashboard_categoria.html", 
                           equipos=equipos, 
                           trabajadores=trabajadores, 
                           categoria=cat,
                           titulo_categoria=nombres_titulos.get(cat, cat),
                           tipos_disponibles=tipos_disponibles,
                           total_cat=total_cat,
                           prestados_cat=prestados_cat,
                           almacen_cat=almacen_cat)

@app.route("/api/equipos/buscar-codigo/<codigo>")
@login_required
def api_buscar_equipo_codigo(codigo):
    codigo_limpio = (codigo or "").strip()
    if not codigo_limpio:
        return jsonify({"error": "Código vacío"}), 400

    conn = get_db()
    equipo = conn.execute(
        """
        SELECT e.id, e.codigo_margesi, e.codigo_interno, e.categoria,
               e.descripcion, e.tipo_equipo, e.marca, e.modelo,
               e.serie_medidas, e.numero_serie, e.estado, p.id AS prestamo_id
        FROM equipos e
        LEFT JOIN prestamos p ON e.id = p.equipo_id AND p.estado = 'ACTIVO'
        WHERE e.codigo_margesi = ? OR e.codigo_interno = ?
        LIMIT 1
        """,
        (codigo_limpio, codigo_limpio)
    ).fetchone()
    conn.close()

    if not equipo:
        return jsonify({"error": f"No se encontró un equipo con el código {codigo_limpio}"}), 404

    return jsonify({
        "id": equipo["id"],
        "codigo": equipo["codigo_margesi"] or equipo["codigo_interno"],
        "categoria": equipo["categoria"],
        "descripcion": equipo["descripcion"] or equipo["tipo_equipo"] or "Equipo",
        "marca": equipo["marca"] or "",
        "modelo": equipo["modelo"] or "",
        "serie": equipo["serie_medidas"] or equipo["numero_serie"] or "",
        "estado": equipo["estado"],
        "prestamo_id": equipo["prestamo_id"]
    })

@app.route("/api/historial-equipo/<int:equipo_id>")
@login_required
def api_historial_equipo(equipo_id):
    conn = get_db()
    cursor = conn.cursor()

    eq = cursor.execute("SELECT * FROM equipos WHERE id = ?", (equipo_id,)).fetchone()
    if not eq:
        conn.close()
        return jsonify({"error": "Equipo no encontrado"}), 404

    query = """
    SELECT 
        p.id,
        p.fecha_prestamo,
        p.fecha_devolucion,
        p.estado,
        p.observaciones_entrega,
        p.observaciones_devolucion,
        p.archivo_acta_entrega,
        t.dni,
        (t.nombres || ' ' || t.apellido_paterno || ' ' || t.apellido_materno) AS custodio,
        t.cargo,
        t.telefono
    FROM prestamos p
    JOIN trabajadores t ON p.trabajador_id = t.id
    WHERE p.equipo_id = ?
    ORDER BY p.fecha_prestamo DESC
    """
    filas = cursor.execute(query, (equipo_id,)).fetchall()
    conn.close()

    historial = []
    for f in filas:
        historial.append({
            "id": f["id"],
            "fecha_prestamo": str(f["fecha_prestamo"] or ""),
            "fecha_devolucion": str(f["fecha_devolucion"] or "Aún en custodia"),
            "duracion": calcular_duracion(f["fecha_prestamo"], f["fecha_devolucion"]),
            "estado": f["estado"],
            "custodio": f["custodio"],
            "dni": f["dni"],
            "cargo": f["cargo"],
            "telefono": f["telefono"],
            "observaciones_entrega": f["observaciones_entrega"] or "-",
            "observaciones_devolucion": f["observaciones_devolucion"] or "-",
            "archivo_acta": f["archivo_acta_entrega"]
        })

    return jsonify({
        "codigo_margesi": eq["codigo_margesi"] or eq["codigo_interno"],
        "descripcion": eq["descripcion"] or eq["tipo_equipo"],
        "marca": eq["marca"],
        "modelo": eq["modelo"],
        "serie": eq["serie_medidas"] or eq["numero_serie"],
        "estado_actual": eq["estado"],
        "total_prestamos": len(historial),
        "historial": historial
    })

@app.route("/equipos/nuevo", methods=["GET", "POST"])
@permiso_requerido("prestamos", "modificar")
@login_required
def registrar_equipo():
    if request.method == "GET":
        return redirect(url_for("prestamos_hub"))

    cat = request.form["categoria"].strip().upper()
    cod = request.form["codigo_margesi"].strip()
    desc = request.form["descripcion"].strip()
    serie = request.form.get("serie_medidas", "").strip()

    conn = get_db()
    conn.cursor().execute("""
        INSERT INTO equipos (codigo_margesi, codigo_interno, categoria, tipo_equipo, descripcion, 
                             marca, modelo, serie_medidas, numero_serie,
                             estado_conservacion, color, sede, valor_inicial, valor_neto,
                             observaciones, celular_institucional, estado)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'EN ALMACEN')
    """, (
        cod,
        cod,
        cat,
        desc,
        desc,
        request.form.get("marca", "").strip(),
        request.form.get("modelo", "").strip(),
        serie,
        serie,
        request.form.get("estado_conservacion", "BUENO").strip(),
        request.form.get("color", "").strip(),
        request.form.get("sede", "ABANCAY").strip(),
        request.form.get("valor_inicial", "").strip(),
        request.form.get("valor_neto", "").strip(),
        request.form.get("observaciones", "").strip(),
        request.form.get("celular_institucional", "").strip()
    ))
    conn.commit()
    conn.close()
    flash("Equipo registrado correctamente en Almacén.", "success")
    return redirect(f"/prestamos/{cat.lower()}")

@app.route("/equipos/editar", methods=["GET", "POST"])
@permiso_requerido("prestamos", "modificar")
@login_required
def editar_equipo():
    if request.method == "GET":
        return redirect(url_for("prestamos_hub"))

    equipo_id = request.form["equipo_id"]
    cat = request.form.get("categoria", "INFORMATICOS").strip().upper()
    cod = request.form["codigo_margesi"].strip()
    desc = request.form["descripcion"].strip()
    serie = request.form.get("serie_medidas", "").strip()

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE equipos 
        SET codigo_margesi = ?, codigo_interno = ?, descripcion = ?, tipo_equipo = ?,
            marca = ?, modelo = ?, serie_medidas = ?, numero_serie = ?,
            estado_conservacion = ?, color = ?, sede = ?, observaciones = ?,
            celular_institucional = ?
        WHERE id = ?
    """, (cod, cod, desc, desc,
          request.form.get("marca", "").strip(),
          request.form.get("modelo", "").strip(),
          serie, serie,
          request.form.get("estado_conservacion", "BUENO").strip(),
          request.form.get("color", "").strip(),
          request.form.get("sede", "ABANCAY").strip(),
          request.form.get("observaciones", "").strip(),
          request.form.get("celular_institucional", "").strip(),
          equipo_id))
    conn.commit()
    conn.close()
    flash("Datos del bien actualizados correctamente.", "success")
    return redirect(f"/prestamos/{cat.lower()}")

@app.route("/equipos/eliminar/<int:equipo_id>", methods=["POST"])
@permiso_requerido("prestamos", "eliminar")
@login_required
def eliminar_equipo(equipo_id):
    cat = request.form.get("categoria", "informaticos").lower()
    conn = get_db()
    cursor = conn.cursor()
    
    eq = cursor.execute("SELECT estado FROM equipos WHERE id = ?", (equipo_id,)).fetchone()
    if eq and eq["estado"] == "EN PRESTAMO":
        flash("No se puede eliminar un bien que se encuentra actualmente en préstamo.", "danger")
    else:
        cursor.execute("DELETE FROM equipos WHERE id = ?", (equipo_id,))
        conn.commit()
        flash("Bien eliminado del inventario correctamente.", "success")
    conn.close()
    return redirect(f"/prestamos/{cat}")

@app.route("/importar-equipos-excel", methods=["POST"])
@permiso_requerido("prestamos", "modificar")
@login_required
def subir_equipos_excel():
    archivo = request.files.get("archivo_equipos")
    categoria_forzada = request.form.get("categoria_forzada", None)
    if archivo and archivo.filename.endswith((".xlsx", ".xls")):
        ruta = os.path.join(UPLOAD_FOLDER, "margesi_equipos.xlsx")
        archivo.save(ruta)
        resultado = importar_equipos_excel(ruta, categoria_forzada)
        flash(f"Clasificación Exitosa: {resultado['MOVILES']} Móviles, {resultado['TOKENS']} Tokens, {resultado['INFORMATICOS']} Informáticos.", "success")
    else:
        flash("Formato inválido. Debe ser un archivo Excel (.xlsx).", "danger")
    return redirect(request.referrer or url_for("prestamos_hub"))

@app.route("/prestar", methods=["POST"])
@app.route("/prestamos/nuevo", methods=["POST"])
@permiso_requerido("prestamos", "modificar")
@login_required
def prestar_equipo():
    equipo_id = request.form.get("equipo_id", "").strip()
    trabajador_id = request.form.get("trabajador_id", "").strip()
    observaciones = request.form.get("observaciones_entrega", request.form.get("observaciones", "")).strip()

    if not equipo_id or not trabajador_id:
        flash("Debe seleccionar el bien y el trabajador antes de confirmar el préstamo.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))

    conn = get_db()
    cursor = conn.cursor()

    eq = cursor.execute("SELECT * FROM equipos WHERE id = ?", (equipo_id,)).fetchone()
    tr = cursor.execute("SELECT * FROM trabajadores WHERE id = ?", (trabajador_id,)).fetchone()

    if not eq or not tr:
        conn.close()
        flash("No se pudo encontrar el bien o el trabajador seleccionado. Intente nuevamente.", "danger")
        return redirect(request.referrer or url_for("prestamos_hub"))

    if eq["estado"] != "EN ALMACEN":
        conn.close()
        flash("El bien seleccionado ya no está disponible en almacén.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))

    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    cod_mostrar = eq["codigo_margesi"] or eq["codigo_interno"] or "EQ"
    nombre_archivo = secure_filename(f"Acta_{cod_mostrar}_{tr['dni']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx")
    ruta_salida = os.path.join(ACTAS_FOLDER, nombre_archivo)

    plantilla = "entrega_celular.docx" if eq["categoria"] == "MOVILES" else "entrega_general.docx"
    ruta_plantilla = os.path.join(PLANTILLAS_FOLDER, plantilla)

    nombre_completo = f"{tr['nombres']} {tr['apellido_paterno']} {tr['apellido_materno']}".strip()

    datos_word = {
        "{{DNI}}": tr["dni"],
        "{{dni}}": tr["dni"],
        "{{NOMBRES_COMPLETOS}}": nombre_completo,
        "{{NOMBRE_COMPLETO}}": nombre_completo,
        "{{NOMBRES}}": tr["nombres"],
        "{{CARGO}}": tr["cargo"],
        "{{TELEFONO}}": tr["telefono"] or "",
        "{{CELULAR_TRABAJADOR}}": tr["celular_institucional"] or "",
        "{{CORREO_INSTITUCIONAL}}": tr["correo_institucional"],
        "{{TIPO_CONTRATO}}": tr["tipo_contrato"],
        "{{FECHA_NACIMIENTO}}": formatear_solo_fecha(tr["fecha_nacimiento"]),
        "{{CODIGO_MARGESI}}": cod_mostrar,
        "{{CODIGO}}": cod_mostrar,
        "{{DESCRIPCION}}": eq["descripcion"] or eq["tipo_equipo"],
        "{{TIPO_EQUIPO}}": eq["descripcion"] or eq["tipo_equipo"],
        "{{MARCA}}": eq["marca"] or "",
        "{{MODELO}}": eq["modelo"] or "",
        "{{SERIE}}": eq["serie_medidas"] or eq["numero_serie"] or "",
        "{{IMEI_1}}": eq["serie_medidas"] or eq["numero_serie"] or "",
        "{{COLOR}}": eq["color"] or "",
        "{{SEDE}}": eq["sede"] or "",
        "{{CELULAR_INSTITUCIONAL}}": eq["celular_institucional"] or "",
        "{{N_LINEA}}": eq["celular_institucional"] or "",
        "{{NUMERO_LINEA}}": eq["celular_institucional"] or "",
        "{{SIM_CARD}}": "",
        "{{ESTADO_CONSERVACION}}": eq["estado_conservacion"] or "",
        "{{OBSERVACIONES_PATRIMONIO}}": eq["observaciones"] or "",
        "{{FECHA_PRESTAMO}}": fecha_actual,
        "{{FECHA}}": fecha_actual,
        "{{OBSERVACIONES}}": observaciones,
        "{{OBSERVACION}}": observaciones
    }

    acta_generada = False
    if os.path.exists(ruta_plantilla):
        try:
            rellenar_plantilla_docx(ruta_plantilla, datos_word, ruta_salida)
            acta_generada = True
        except Exception:
            nombre_archivo = None

    cursor.execute("""
        INSERT INTO prestamos (equipo_id, trabajador_id, fecha_prestamo, estado, observaciones_entrega, archivo_acta_entrega)
        VALUES (?, ?, CURRENT_TIMESTAMP, 'ACTIVO', ?, ?)
    """, (equipo_id, trabajador_id, observaciones, nombre_archivo))

    cursor.execute("UPDATE equipos SET estado = 'EN PRESTAMO' WHERE id = ?", (equipo_id,))
    if eq["categoria"] == "MOVILES" and (eq["celular_institucional"] or "").strip():
        cursor.execute(
            "UPDATE trabajadores SET celular_institucional = ? WHERE id = ?",
            (eq["celular_institucional"].strip(), trabajador_id)
        )
    conn.commit()
    conn.close()

    if acta_generada:
        flash(f"Préstamo registrado para {nombre_completo}.", "success")
    else:
        flash(f"Préstamo registrado para {nombre_completo}, pero no se pudo generar el acta Word.", "warning")
    return redirect(request.referrer or url_for("prestamos_hub"))

@app.route("/devolver", methods=["GET", "POST"])
@app.route("/prestamos/devolver", methods=["GET", "POST"])
@permiso_requerido("prestamos", "modificar")
@login_required
def devolver_equipo():
    if request.method == "GET":
        return redirect(request.referrer or url_for("prestamos_hub"))

    prestamo_id = request.form.get("prestamo_id", "").strip()
    if not prestamo_id:
        equipo_id = request.form.get("equipo_id", "").strip()
        if equipo_id:
            conn = get_db()
            prestamo = conn.execute(
                "SELECT id FROM prestamos WHERE equipo_id = ? AND estado = 'ACTIVO'",
                (equipo_id,)
            ).fetchone()
            conn.close()
            prestamo_id = str(prestamo["id"]) if prestamo else ""

    if not prestamo_id:
        flash("No se encontró un préstamo activo para devolver.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))

    obs_dev = request.form.get("observaciones_devolucion", "").strip()

    conn = get_db()
    cursor = conn.cursor()

    prestamo = cursor.execute("SELECT equipo_id FROM prestamos WHERE id = ?", (prestamo_id,)).fetchone()
    if prestamo:
        cursor.execute("""
            UPDATE prestamos 
            SET fecha_devolucion = CURRENT_TIMESTAMP, estado = 'DEVUELTO', observaciones_devolucion = ?
            WHERE id = ?
        """, (obs_dev, prestamo_id))

        cursor.execute("UPDATE equipos SET estado = 'EN ALMACEN' WHERE id = ?", (prestamo["equipo_id"],))
        conn.commit()
    else:
        flash("El préstamo seleccionado no existe o ya fue devuelto.", "warning")
        conn.close()
        return redirect(request.referrer or url_for("prestamos_hub"))

    conn.close()
    flash("Equipo retornado exitosamente a EN ALMACÉN.", "success")
    return redirect(request.referrer or url_for("prestamos_hub"))

@app.route("/descargar-acta/<nombre_archivo>")
@app.route("/actas/descargar/<nombre_archivo>")
@login_required
def descargar_acta(nombre_archivo):
    nombre_archivo = os.path.basename(nombre_archivo)
    ruta = os.path.join(ACTAS_FOLDER, nombre_archivo)
    if os.path.exists(ruta):
        return send_file(ruta, as_attachment=True)
    flash("El archivo no se encuentra en el servidor.", "warning")
    return redirect(request.referrer or url_for("prestamos_hub"))

# ================= PERSONAL / TRABAJADORES =================
@app.route("/trabajadores")
@permiso_requerido("personal", "ver")
@login_required
def ver_trabajadores():
    conn = get_db()
    cursor = conn.cursor()

    trabajadores = cursor.execute("""
        SELECT * FROM trabajadores 
        WHERE activo = 1
        ORDER BY apellido_paterno ASC, apellido_materno ASC, nombres ASC
    """).fetchall()
    trabajadores_antiguos = cursor.execute("""
        SELECT * FROM trabajadores
        WHERE activo = 0
        ORDER BY apellido_paterno ASC, apellido_materno ASC, nombres ASC
    """).fetchall()
    contratos = cursor.execute("""
        SELECT id, trabajador_id, tipo_documento, nombre_archivo
        FROM contratos_trabajadores
        ORDER BY creado_el DESC, id DESC
    """).fetchall()
    contratos_por_trabajador = {}
    for contrato in contratos:
        contratos_por_trabajador.setdefault(contrato["trabajador_id"], {})[contrato["tipo_documento"]] = contrato
    conn.close()
    return render_template("trabajadores.html", trabajadores=trabajadores,
                           trabajadores_antiguos=trabajadores_antiguos,
                           contratos_por_trabajador=contratos_por_trabajador)

@app.route("/trabajadores/nuevo", methods=["POST"])
@permiso_requerido("personal", "modificar")
@login_required
def registrar_trabajador():
    dni = request.form["dni"].strip()
    nombres = request.form["nombres"].strip()
    apellido_paterno = request.form["apellido_paterno"].strip()
    apellido_materno = request.form["apellido_materno"].strip()
    cargo = request.form.get("cargo", "").strip()
    correo_personal = request.form.get("correo_personal", "").strip()
    correo_institucional = request.form.get("correo_institucional", "").strip()
    tipo_contrato = request.form.get("tipo_contrato", "").strip()
    telefono = request.form.get("telefono", "").strip()
    celular_institucional = request.form.get("celular_institucional", "").strip()
    fecha_nacimiento = formatear_solo_fecha(request.form.get("fecha_nacimiento", "").strip())

    conn = get_db()
    cursor = conn.cursor()
    try:
        existente = cursor.execute("SELECT id FROM trabajadores WHERE dni = ?", (dni,)).fetchone()
        datos = (nombres, apellido_paterno, apellido_materno, cargo, correo_personal,
                 correo_institucional, tipo_contrato, telefono, celular_institucional,
                 fecha_nacimiento)
        if existente:
            cursor.execute("""
                UPDATE trabajadores
                SET nombres = ?, apellido_paterno = ?, apellido_materno = ?, cargo = ?,
                    correo_personal = ?, correo_institucional = ?, tipo_contrato = ?, telefono = ?,
                    celular_institucional = ?, fecha_nacimiento = ?, activo = 1
                WHERE id = ?
            """, datos + (existente["id"],))
        else:
            cursor.execute("""
                INSERT INTO trabajadores (dni, nombres, apellido_paterno, apellido_materno, cargo,
                                          correo_personal, correo_institucional, tipo_contrato, telefono,
                                          celular_institucional, fecha_nacimiento, activo)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
            """, (dni,) + datos)
        conn.commit()
        flash(f"Trabajador {nombres} {apellido_paterno} registrado o reactivado correctamente.", "success")
    except Exception:
        flash(f"Error: El DNI {dni} ya se encuentra registrado.", "danger")
    finally:
        conn.close()

    return redirect(request.referrer or url_for("ver_trabajadores"))

@app.route("/trabajadores/editar", methods=["POST"])
@permiso_requerido("personal", "modificar")
@login_required
def editar_trabajador():
    trabajador_id = request.form["trabajador_id"]
    dni = request.form["dni"].strip()
    nombres = request.form["nombres"].strip()
    apellido_paterno = request.form["apellido_paterno"].strip()
    apellido_materno = request.form["apellido_materno"].strip()
    cargo = request.form.get("cargo", "").strip()
    correo_personal = request.form.get("correo_personal", "").strip()
    correo_institucional = request.form.get("correo_institucional", "").strip()
    tipo_contrato = request.form.get("tipo_contrato", "").strip()
    telefono = request.form.get("telefono", "").strip()
    celular_institucional = request.form.get("celular_institucional", "").strip()
    fecha_nacimiento = formatear_solo_fecha(request.form.get("fecha_nacimiento", "").strip())

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            UPDATE trabajadores 
            SET dni = ?, nombres = ?, apellido_paterno = ?, apellido_materno = ?, cargo = ?,
                                correo_personal = ?, correo_institucional = ?, tipo_contrato = ?, telefono = ?,
                                celular_institucional = ?, fecha_nacimiento = ?
            WHERE id = ?
                """, (dni, nombres, apellido_paterno, apellido_materno, cargo, correo_personal,
                            correo_institucional, tipo_contrato, telefono, celular_institucional,
                            fecha_nacimiento, trabajador_id))
        conn.commit()
        flash(f"Datos de {nombres} {apellido_paterno} actualizados correctamente.", "success")
    except Exception:
        flash(f"Error al actualizar: Verifique que el DNI no esté duplicado.", "danger")
    finally:
        conn.close()

    return redirect(url_for("ver_trabajadores"))

@app.route("/trabajadores/eliminar/<int:trabajador_id>", methods=["POST"])
@permiso_requerido("personal", "eliminar")
@login_required
def eliminar_trabajador(trabajador_id):
    conn = get_db()
    cursor = conn.cursor()

    prestamos_activos = cursor.execute("""
        SELECT COUNT(*) FROM prestamos 
        WHERE trabajador_id = ? AND estado = 'ACTIVO'
    """, (trabajador_id,)).fetchone()[0]

    if prestamos_activos > 0:
        flash("No se puede eliminar al trabajador porque tiene equipos actualmente en préstamo.", "danger")
    else:
        cursor.execute("UPDATE trabajadores SET activo = 0 WHERE id = ?", (trabajador_id,))
        conn.commit()
        flash("Trabajador trasladado a Trabajadores Antiguos.", "success")

    conn.close()
    return redirect(url_for("ver_trabajadores"))

@app.route("/trabajadores/cambiar-estado/<int:trabajador_id>", methods=["POST"])
@permiso_requerido("personal", "eliminar")
@login_required
def cambiar_estado_trabajador(trabajador_id):
    conn = get_db()
    trabajador = conn.execute("SELECT activo FROM trabajadores WHERE id = ?", (trabajador_id,)).fetchone()
    if trabajador:
        nuevo_estado = 0 if trabajador["activo"] else 1
        conn.execute("UPDATE trabajadores SET activo = ? WHERE id = ?", (nuevo_estado, trabajador_id))
        conn.commit()
        flash("Trabajador activado correctamente." if nuevo_estado else "Trabajador trasladado a Trabajadores Antiguos.", "success")
    conn.close()
    return redirect(url_for("ver_trabajadores"))

@app.route("/trabajadores/<int:trabajador_id>/contrato", methods=["POST"])
@login_required
def subir_contrato_trabajador(trabajador_id):
    tipo_documento = request.form.get("tipo_documento", "").strip().upper()
    archivo = request.files.get("archivo_contrato")
    if tipo_documento not in ("OS", "CAS") or not archivo or not archivo.filename:
        flash("Seleccione un archivo válido y el tipo OS o CAS.", "danger")
        return redirect(url_for("ver_trabajadores"))

    ruta_archivo, nombre_original = guardar_archivo_subido(
        archivo, CONTRATOS_FOLDER, {".pdf", ".doc", ".docx"}
    )
    ruta_relativa = os.path.relpath(ruta_archivo, UPLOAD_FOLDER)
    conn = get_db()
    conn.execute("DELETE FROM contratos_trabajadores WHERE trabajador_id = ? AND tipo_documento = ?", (trabajador_id, tipo_documento))
    conn.execute("INSERT INTO contratos_trabajadores (trabajador_id, tipo_documento, nombre_archivo, ruta_archivo) VALUES (?, ?, ?, ?)",
                 (trabajador_id, tipo_documento, nombre_original, ruta_relativa))
    conn.commit()
    conn.close()
    flash(f"Documento {tipo_documento} guardado correctamente.", "success")
    return redirect(url_for("ver_trabajadores"))

@app.route("/trabajadores/contrato/<int:contrato_id>")
@login_required
def descargar_contrato_trabajador(contrato_id):
    conn = get_db()
    contrato = conn.execute("SELECT nombre_archivo, ruta_archivo FROM contratos_trabajadores WHERE id = ?", (contrato_id,)).fetchone()
    conn.close()
    if not contrato:
        flash("El contrato no existe.", "warning")
        return redirect(url_for("ver_trabajadores"))
    ruta = os.path.join(UPLOAD_FOLDER, contrato["ruta_archivo"])
    if not os.path.isfile(ruta):
        flash("El archivo del contrato no se encuentra en el servidor.", "warning")
        return redirect(url_for("ver_trabajadores"))
    return send_file(ruta, as_attachment=True, download_name=contrato["nombre_archivo"])

@app.route("/trabajadores/exportar-excel")
@permiso_requerido("personal", "reportar")
@login_required
def exportar_trabajadores_excel():
    campos, etiquetas, filas = obtener_datos_exportacion_trabajadores()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Personal"
    ws.append(etiquetas)

    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    thin_border = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
    )
    for col_num in range(1, len(etiquetas) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions.height = 28

    for r_idx, fila in enumerate(filas, start=2):
        ws.append([valor_exportacion_trabajador(fila, campo) for campo in campos])
        for col_num in range(1, len(etiquetas) + 1):
            celda = ws.cell(row=r_idx, column=col_num)
            celda.border = thin_border
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 34)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    nombre_descarga = f"Lista_Trabajadores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


TRABAJADORES_EXPORTABLES = {
    "dni": ("DNI", "dni"),
    "nombre": ("APELLIDOS Y NOMBRES", "nombre"),
    "cargo": ("CARGO", "cargo"),
    "tipo_contrato": ("TIPO DE CONTRATO", "tipo_contrato"),
    "telefono": ("TELÉFONO", "telefono"),
    "celular_institucional": ("CELULAR INSTITUCIONAL", "celular_institucional"),
    "correo_institucional": ("CORREO INSTITUCIONAL", "correo_institucional"),
    "correo_personal": ("CORREO PERSONAL", "correo_personal"),
    "fecha_nacimiento": ("FECHA DE NACIMIENTO", "fecha_nacimiento"),
}


def valor_exportacion_trabajador(fila, campo):
    if campo == "nombre":
        return " ".join(filter(None, [fila["apellido_paterno"], fila["apellido_materno"], fila["nombres"]]))
    if campo == "fecha_nacimiento":
        return formatear_solo_fecha(fila[campo])
    return str(fila[campo] or "")


def obtener_datos_exportacion_trabajadores():
    solicitados = request.args.getlist("campos")
    campos = [campo for campo in solicitados if campo in TRABAJADORES_EXPORTABLES]
    if not campos:
        campos = list(TRABAJADORES_EXPORTABLES)
    estado = request.args.get("estado", "activos")
    condicion = "" if estado == "todos" else " AND activo = ?"
    parametros = () if estado == "todos" else (1 if estado == "activos" else 0,)

    conn = get_db()
    cursor = conn.cursor()
    filas = cursor.execute("""
        SELECT nombres, apellido_paterno, apellido_materno, dni, cargo,
             correo_personal, correo_institucional, tipo_contrato, telefono,
             celular_institucional, fecha_nacimiento, activo
        FROM trabajadores
        WHERE 1 = 1""" + condicion + """
        ORDER BY apellido_paterno ASC, apellido_materno ASC, nombres ASC
    """, parametros).fetchall()
    conn.close()
    etiquetas = [TRABAJADORES_EXPORTABLES[campo][0] for campo in campos]
    return campos, etiquetas, filas


@app.route("/trabajadores/exportar-pdf")
@permiso_requerido("personal", "reportar")
@login_required
def exportar_trabajadores_pdf():
    campos, etiquetas, filas = obtener_datos_exportacion_trabajadores()
    orientacion = request.args.get("orientacion", "auto")
    if orientacion not in ("auto", "vertical", "horizontal"):
        orientacion = "auto"
    es_horizontal = orientacion == "horizontal" or (orientacion == "auto" and len(campos) > 5)
    pagina = landscape(A4) if es_horizontal else A4
    margen = 24
    buffer = BytesIO()
    documento = SimpleDocTemplate(buffer, pagesize=pagina, rightMargin=margen,
                                  leftMargin=margen, topMargin=28, bottomMargin=28,
                                  title="Reporte de Personal")
    estilos = getSampleStyleSheet()
    estilo_celda = estilos["Normal"].clone("celda_exportacion")
    estilo_celda.fontSize = 7.2 if es_horizontal else 6.8
    estilo_celda.leading = estilo_celda.fontSize + 1.5
    estilo_celda.alignment = 1
    datos = [[Paragraph(f"<b>{etiqueta}</b>", estilo_celda) for etiqueta in etiquetas]]
    for fila in filas:
        datos.append([Paragraph(escape_xml(valor_exportacion_trabajador(fila, campo)), estilo_celda) for campo in campos])
    ancho_util = pagina[0] - (margen * 2)
    pesos_columnas = {
        "nombre": 1.55,
        "cargo": 1.25,
        "correo_institucional": 2.25,
        "correo_personal": 2.25,
        "tipo_contrato": 1.35,
        "celular_institucional": 1.3,
        "fecha_nacimiento": 1.25,
        "dni": 0.9,
        "telefono": 1.1,
    }
    pesos = [pesos_columnas.get(campo, 1) for campo in campos]
    suma_pesos = sum(pesos)
    tabla = Table(datos, colWidths=[ancho_util * peso / suma_pesos for peso in pesos], repeatRows=1, hAlign="CENTER")
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#075985")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#64748b")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f9")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    documento.build([
        Paragraph("<b>PADRÓN DE PERSONAL</b>", estilos["Title"]),
        Spacer(1, 6),
        Paragraph(f"Reporte generado el {datetime.now().strftime('%d/%m/%Y %H:%M')} | Registros: {len(filas)}", estilos["Normal"]),
        Spacer(1, 12), tabla
    ])
    buffer.seek(0)
    nombre_descarga = f"Reporte_Personal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/pdf")

@app.route("/importar-excel", methods=["POST"])
@permiso_requerido("personal", "modificar")
@login_required
def subir_excel():
    archivo = request.files.get("archivo_excel")
    if archivo and archivo.filename.lower().endswith(".xlsx"):
        ruta, _ = guardar_archivo_subido(archivo, UPLOAD_FOLDER, {".xlsx"})
        resultado = importar_trabajadores_excel(ruta)
        if isinstance(resultado, dict):
            total = resultado.get("total", 0)
            nuevos = resultado.get("nuevos", 0)
            actualizados = resultado.get("actualizados", 0)
            diferencias = resultado.get("mismatches", [])
            mensaje = f"Se procesaron {total} trabajadores ({nuevos} nuevos, {actualizados} actualizados)."
            if diferencias:
                detalle = "; ".join(
                    f"{d['dni']} ({len(d['cambios'])} cambios)" for d in diferencias[:3]
                )
                if len(diferencias) > 3:
                    detalle += f" ... y {len(diferencias) - 3} más"
                flash(f"{mensaje} Hay diferencias en: {detalle}", "warning")
            else:
                flash(mensaje, "success")
        else:
            flash(f"Se importaron / actualizaron {resultado} trabajadores exitosamente.", "success")
    else:
        flash("Formato inválido. Seleccione un archivo Excel (.xlsx).", "danger")
    return redirect(request.referrer or url_for("ver_trabajadores"))

@app.route("/trabajadores/importar-excel", methods=["POST"])
@permiso_requerido("personal", "modificar")
@login_required
def subir_excel_trabajadores():
    archivo = request.files.get("archivo_excel_trabajadores")
    if archivo and archivo.filename.lower().endswith(".xlsx"):
        ruta, _ = guardar_archivo_subido(archivo, UPLOAD_FOLDER, {".xlsx"})
        resultado = importar_trabajadores_excel(ruta)
        if isinstance(resultado, dict):
            total = resultado.get("total", 0)
            nuevos = resultado.get("nuevos", 0)
            actualizados = resultado.get("actualizados", 0)
            diferencias = resultado.get("mismatches", [])
            mensaje = f"Se procesaron {total} trabajadores ({nuevos} nuevos, {actualizados} actualizados)."
            if diferencias:
                detalle = "; ".join(
                    f"{d['dni']} ({', '.join(d['cambios'][:2])}{'...' if len(d['cambios']) > 2 else ''})" for d in diferencias[:3]
                )
                if len(diferencias) > 3:
                    detalle += f" ... y {len(diferencias) - 3} más"
                flash(f"{mensaje} Diferencias detectadas: {detalle}", "warning")
            else:
                flash(mensaje, "success")
        else:
            flash(f"Se importaron / actualizaron {resultado} trabajadores exitosamente.", "success")
    else:
        flash("Formato inválido. Seleccione un archivo Excel (.xlsx) para personal.", "danger")
    return redirect(url_for("ver_trabajadores"))

# ================= MÓDULO DE REPORTES Y TRAZABILIDAD =================
@app.route("/reportes")
@permiso_requerido("reportes", "ver")
@login_required
def ver_reportes():
    filtro_estado = request.args.get("estado", "TODOS")
    categoria = request.args.get("categoria", "TODAS").upper()
    if categoria not in ["TODAS", "MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"]:
        categoria = "TODAS"
    periodo = request.args.get("periodo", "TODOS").upper()
    f_desde = request.args.get("desde", "")
    f_hasta = request.args.get("hasta", "")
    
    conn = get_db()
    cursor = conn.cursor()

    categoria_where = "" if categoria == "TODAS" else " AND categoria = ?"
    categoria_param = [] if categoria == "TODAS" else [categoria]

    total_equipos = cursor.execute(
        "SELECT COUNT(*) FROM equipos WHERE categoria IN ('MOVILES', 'TOKENS', 'INFORMATICOS', 'MOBILIARIO')" + categoria_where,
        categoria_param
    ).fetchone()[0]
    total_almacen = cursor.execute(
        "SELECT COUNT(*) FROM equipos WHERE estado = 'EN ALMACEN' AND categoria IN ('MOVILES', 'TOKENS', 'INFORMATICOS', 'MOBILIARIO')" + categoria_where,
        categoria_param
    ).fetchone()[0]
    total_prestamo = cursor.execute(
        "SELECT COUNT(*) FROM equipos WHERE estado = 'EN PRESTAMO' AND categoria IN ('MOVILES', 'TOKENS', 'INFORMATICOS', 'MOBILIARIO')" + categoria_where,
        categoria_param
    ).fetchone()[0]
    total_historico_query = "SELECT COUNT(*) FROM prestamos p JOIN equipos e ON p.equipo_id = e.id"
    if categoria != "TODAS":
        total_historico_query += " WHERE e.categoria = ?"
    total_historico = cursor.execute(total_historico_query, categoria_param).fetchone()[0]

    base_query = """
    SELECT 
        p.id AS prestamo_id,
        p.fecha_prestamo,
        p.fecha_devolucion,
        p.estado AS estado_prestamo,
        p.observaciones_entrega,
        p.observaciones_devolucion,
        p.archivo_acta_entrega,
        e.codigo_margesi,
        e.categoria,
        e.descripcion,
        e.marca,
        e.modelo,
        e.serie_medidas,
        t.dni AS trabajador_dni,
        (t.nombres || ' ' || t.apellido_paterno || ' ' || t.apellido_materno) AS trabajador_nombre,
        t.cargo AS trabajador_cargo,
        t.telefono AS trabajador_telefono
    FROM prestamos p
    JOIN equipos e ON p.equipo_id = e.id
    JOIN trabajadores t ON p.trabajador_id = t.id
    WHERE 1=1
    """
    params = []

    if categoria != "TODAS":
        base_query += " AND e.categoria = ?"
        params.append(categoria)

    if filtro_estado == "EN_PRESTAMO":
        base_query += " AND p.estado = 'ACTIVO'"
    elif filtro_estado == "DEVUELTOS":
        base_query += " AND p.estado = 'DEVUELTO'"

    if periodo == "HOY":
        base_query += " AND (date(p.fecha_prestamo) = date('now', 'localtime') OR date(p.fecha_devolucion) = date('now', 'localtime'))"
    elif periodo == "SEMANA":
        base_query += " AND (date(p.fecha_prestamo) >= date('now', '-7 days', 'localtime') OR date(p.fecha_devolucion) >= date('now', '-7 days', 'localtime'))"
    elif periodo == "MES":
        base_query += " AND (strftime('%Y-%m', p.fecha_prestamo) = strftime('%Y-%m', 'now', 'localtime') OR strftime('%Y-%m', p.fecha_devolucion) = strftime('%Y-%m', 'now', 'localtime'))"
    elif periodo == "RANGO" and f_desde and f_hasta:
        base_query += " AND date(p.fecha_prestamo) BETWEEN ? AND ?"
        params.extend([f_desde, f_hasta])

    base_query += " ORDER BY p.fecha_prestamo DESC, p.id DESC"
    movimientos = cursor.execute(base_query, params).fetchall()

    equipos_almacen = cursor.execute("""
        SELECT * FROM equipos 
        WHERE estado = 'EN ALMACEN' AND categoria IN ('MOVILES', 'TOKENS', 'INFORMATICOS', 'MOBILIARIO')
    """ + categoria_where + " ORDER BY codigo_margesi ASC", categoria_param).fetchall()
    conn.close()

    return render_template("reportes.html", 
                           movimientos=movimientos, 
                           equipos_almacen=equipos_almacen,
                           filtro_estado=filtro_estado,
                           categoria=categoria,
                           periodo=periodo,
                           f_desde=f_desde,
                           f_hasta=f_hasta,
                           total_equipos=total_equipos,
                           total_almacen=total_almacen,
                           total_prestamo=total_prestamo,
                           total_historico=total_historico)

@app.route("/reportes/exportar-excel")
@app.route("/reportes/exportar-excel/")
@permiso_requerido("reportes", "reportar")
@login_required
def exportar_reporte_excel():
    filtro_estado = request.args.get("estado", "TODOS")
    categoria = request.args.get("categoria", "TODAS").upper()
    if categoria not in ["TODAS", "MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"]:
        categoria = "TODAS"
    periodo = request.args.get("periodo", "TODOS").upper()
    f_desde = request.args.get("desde", "")
    f_hasta = request.args.get("hasta", "")

    conn = get_db()
    cursor = conn.cursor()

    base_query = """
    SELECT 
        e.codigo_margesi,
        e.categoria,
        e.descripcion,
        e.marca,
        e.modelo,
        e.serie_medidas,
        e.celular_institucional,
        t.dni,
        (t.nombres || ' ' || t.apellido_paterno || ' ' || t.apellido_materno) AS trabajador,
        t.cargo,
        t.telefono,
        p.fecha_prestamo,
        p.fecha_devolucion,
        p.estado,
        p.observaciones_entrega,
        p.observaciones_devolucion
    FROM prestamos p
    JOIN equipos e ON p.equipo_id = e.id
    JOIN trabajadores t ON p.trabajador_id = t.id
    WHERE 1=1
    """
    params = []

    if categoria != "TODAS":
        base_query += " AND e.categoria = ?"
        params.append(categoria)

    if filtro_estado == "EN_PRESTAMO":
        base_query += " AND p.estado = 'ACTIVO'"
    elif filtro_estado == "DEVUELTOS":
        base_query += " AND p.estado = 'DEVUELTO'"

    if periodo == "HOY":
        base_query += " AND (date(p.fecha_prestamo) = date('now', 'localtime') OR date(p.fecha_devolucion) = date('now', 'localtime'))"
    elif periodo == "SEMANA":
        base_query += " AND (date(p.fecha_prestamo) >= date('now', '-7 days', 'localtime') OR date(p.fecha_devolucion) >= date('now', '-7 days', 'localtime'))"
    elif periodo == "MES":
        base_query += " AND (strftime('%Y-%m', p.fecha_prestamo) = strftime('%Y-%m', 'now', 'localtime') OR strftime('%Y-%m', p.fecha_devolucion) = strftime('%Y-%m', 'now', 'localtime'))"
    elif periodo == "RANGO" and f_desde and f_hasta:
        base_query += " AND date(p.fecha_prestamo) BETWEEN ? AND ?"
        params.extend([f_desde, f_hasta])

    base_query += " ORDER BY p.fecha_prestamo DESC"
    filas = cursor.execute(base_query, params).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Historial y Trazabilidad"

    headers = [
        "Cód. Margesí", "Categoría", "Descripción", "Marca", "Modelo", "N° Serie/Medidas",
        "Celular Institucional",
        "DNI Custodio", "Trabajador", "Cargo", "Teléfono",
        "Fecha Préstamo (Salida)", "Fecha Devolución (Retorno)", "Tiempo de Uso / Duración",
        "Estado", "Obs. Entrega", "Obs. Devolución"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions.height = 28

    cols_reportes = (1, 2, 7, 10, 11, 12, 13, 14)
    for r_idx, f in enumerate(filas, start=2):
        ws.append([
            f["codigo_margesi"], f["categoria"], f["descripcion"], f["marca"], f["modelo"], f["serie_medidas"],
            f["celular_institucional"],
            f["dni"], f["trabajador"], f["cargo"], f["telefono"],
            str(f["fecha_prestamo"] or ""), str(f["fecha_devolucion"] or "Aún en custodia"),
            calcular_duracion(f["fecha_prestamo"], f["fecha_devolucion"]),
            "EN PRÉSTAMO" if f["estado"] == "ACTIVO" else "DEVUELTO",
            f["observaciones_entrega"] or "", f["observaciones_devolucion"] or ""
        ])
        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=r_idx, column=col_num)
            c.border = thin_border
            if col_num in cols_reportes:
                c.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    nombre_descarga = f"Historial_Prestamos_{periodo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================================================================
# FUNCIÓN: EXPORTAR REPORTE DE EQUIPOS A EXCEL (MÓVILES / TOKENS / ETC.)
# =========================================================================
@app.route("/prestamos/exportar-excel/<categoria>")
@permiso_requerido("prestamos", "reportar")
@login_required
def exportar_equipos_excel(categoria):
    cat = categoria.upper()
    filtro_estado = request.args.get("estado", "TODOS").upper()
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Consulta a la base de datos con los datos del equipo y del trabajador/custodio
    base_query = """
        SELECT e.codigo_margesi, e.descripcion, e.marca, e.modelo, 
             e.serie_medidas, e.celular_institucional, e.estado_conservacion, e.sede, e.estado,
               t.dni AS trabajador_dni,
               (t.nombres || ' ' || t.apellido_paterno || ' ' || t.apellido_materno) AS trabajador_nombre,
               t.cargo AS trabajador_cargo,
               t.telefono AS trabajador_telefono,
               p.fecha_prestamo,
               p.observaciones_entrega
        FROM equipos e
        LEFT JOIN prestamos p ON e.id = p.equipo_id AND p.estado = 'ACTIVO'
        LEFT JOIN trabajadores t ON p.trabajador_id = t.id
        WHERE e.categoria = ?
    """
    params = [cat]
    
    # Filtro según la opción elegida
    if filtro_estado == "PRESTADOS":
        base_query += " AND e.estado = 'EN PRESTAMO'"
    elif filtro_estado == "ALMACEN":
        base_query += " AND e.estado = 'EN ALMACEN'"
        
    base_query += " ORDER BY CASE WHEN e.estado = 'EN PRESTAMO' THEN 0 ELSE 1 END, p.fecha_prestamo DESC, e.codigo_margesi ASC"
    
    filas = cursor.execute(base_query, params).fetchall()
    conn.close()
    
    # Crear libro de Excel con formato institucional
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Reporte {cat}"
    
    headers = [
        "N°", "CÓD. MARGESÍ", "DESCRIPCIÓN / TIPO", "MARCA", "MODELO",
        "SERIE / IMEI", "CELULAR INSTITUCIONAL", "ESTADO FÍSICO", "SEDE", "ESTADO DE CONTROL",
        "CUSTODIO ACTUAL", "DNI", "CARGO", "TELÉFONO", "FECHA SALIDA / PRÉSTAMO", "OBSERVACIONES"
    ]
    ws.append(headers)
    
    # Estilos de encabezado azul institucional
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[1].height = 28
    
    cols_centradas = (1, 2, 4, 5, 7, 8, 9, 11, 13, 14)
    for r_idx, f in enumerate(filas, start=2):
        row_data = [
            r_idx - 1,
            f["codigo_margesi"] or "",
            f["descripcion"] or "",
            f["marca"] or "-",
            f["modelo"] or "-",
            f["serie_medidas"] or "-",
            f["celular_institucional"] or "-",
            f["estado_conservacion"] or "BUENO",
            f["sede"] or "ABANCAY",
            "EN PRÉSTAMO" if f["estado"] == "EN PRESTAMO" else "EN ALMACÉN",
            f["trabajador_nombre"] if f["estado"] == "EN PRESTAMO" else "En Almacén",
            f["trabajador_dni"] if f["estado"] == "EN PRESTAMO" else "-",
            f["trabajador_cargo"] if f["estado"] == "EN PRESTAMO" else "-",
            f["trabajador_telefono"] if f["estado"] == "EN PRESTAMO" else "-",
            str(f["fecha_prestamo"]) if f["estado"] == "EN PRESTAMO" and f["fecha_prestamo"] else "-",
            f["observaciones_entrega"] or "-"
        ]
        ws.append(row_data)
        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=r_idx, column=col_num)
            c.border = thin_border
            if col_num in cols_centradas:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    # Ajustar ancho de columnas automáticamente
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    nombre_descarga = f"Reporte_{cat}_{filtro_estado}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=nombre_descarga,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    init_db()
    init_biblioteca_db()
    certificado = os.path.join("certificados", "sisgemosi-cert.pem")
    clave = os.path.join("certificados", "sisgemosi-key.pem")
    usar_https = os.getenv("SISGEMOSI_HTTPS", "1").strip().lower() in ("1", "true", "si", "yes")
    ssl_config = None
    if usar_https:
        certificados_disponibles = os.path.exists(certificado) and os.path.exists(clave)
        if PRODUCTION_MODE and not certificados_disponibles:
            raise RuntimeError("Producción requiere un certificado HTTPS real en certificados/.")
        ssl_config = (certificado, clave) if certificados_disponibles else "adhoc"
    host = "127.0.0.1" if PRODUCTION_MODE else "0.0.0.0"
    app.run(host=host, port=5000, debug=False, ssl_context=ssl_config)
