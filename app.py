import os
import json
import hashlib
import hmac
import math
import re
import secrets
import unicodedata
import urllib.parse
import urllib.request
import time
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
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from database import (
    get_db, init_db, get_biblioteca_db, init_biblioteca_db,
    ensure_apowerrec_schema, ensure_access_matrix_schema, normalizar_numero_celular,
)
from excel_loader import importar_trabajadores_excel, formatear_solo_fecha
from equipos_loader import importar_equipos_excel
from docx_generator import rellenar_plantilla_docx, generar_informe_mensual_docx
from docx_generator import rellenar_plantilla_docx, generar_informe_mensual_docx, generar_reporte_actividades_docx
from sgd_processor import (
    procesar_sgd_excel, generar_excel_sgd, aplicar_destinatarios_ocultos,
    aplicar_umbral_antiguedad,
)

app = Flask(__name__)
PRODUCTION_MODE = os.environ.get("SISGEMOSI_ENV", "development").lower() == "production"
secret_key = os.environ.get("SISGEMOSI_SECRET_KEY")
SISGEMOSI_SECRET_KEY_CONFIGURED = bool(secret_key)
if PRODUCTION_MODE and not secret_key:
    raise RuntimeError("Defina SISGEMOSI_SECRET_KEY antes de iniciar en producción.")
app.secret_key = secret_key or "sisgemosi_dev_secret_key_default_2026_fixed"
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


def buscar_trabajador_con_celular(cursor, numero, trabajador_id_excluir=None):
    celular_normalizado = normalizar_numero_celular(numero)
    if not celular_normalizado:
        return None

    consulta = """
        SELECT id, dni, nombres, apellido_paterno, apellido_materno
        FROM trabajadores
        WHERE celular_institucional_normalizado = ?
    """
    parametros = [celular_normalizado]
    if trabajador_id_excluir is not None:
        consulta += " AND id != ?"
        parametros.append(trabajador_id_excluir)
    return cursor.execute(consulta, parametros).fetchone()


def celular_ya_registrado_en_otro_equipo(cursor, numero, equipo_id_excluir=None):
    celular_normalizado = normalizar_numero_celular(numero)
    if not celular_normalizado:
        return False

    consulta = "SELECT celular_institucional FROM equipos WHERE categoria = 'MOVILES'"
    parametros = []
    if equipo_id_excluir is not None:
        consulta += " AND id != ?"
        parametros.append(equipo_id_excluir)
    return any(
        normalizar_numero_celular(fila["celular_institucional"]) == celular_normalizado
        for fila in cursor.execute(consulta, parametros).fetchall()
    )


def normalizar_identificador_equipo(identificador):
    return re.sub(r"[^A-Z0-9]", "", str(identificador or "").upper())


def buscar_movil_con_imei(cursor, imei, equipo_id_excluir=None, codigo=None):
    imei_normalizado = normalizar_identificador_equipo(imei)
    codigo_normalizado = normalizar_identificador_equipo(codigo)
    if not imei_normalizado and not codigo_normalizado:
        return None

    consulta = """
        SELECT id, categoria, codigo_margesi, codigo_interno, descripcion,
               serie_medidas, numero_serie
        FROM equipos
    """
    parametros = []
    if equipo_id_excluir is not None:
        consulta += " AND id != ?"
        parametros.append(equipo_id_excluir)

    for equipo in cursor.execute(consulta, parametros).fetchall():
        codigos_existentes = {
            normalizar_identificador_equipo(valor)
            for valor in (equipo["codigo_margesi"], equipo["codigo_interno"])
            if normalizar_identificador_equipo(valor)
        }
        identificadores_existentes = {
            normalizar_identificador_equipo(valor)
            for valor in (
                equipo["codigo_margesi"],
                equipo["codigo_interno"],
                equipo["serie_medidas"],
                equipo["numero_serie"],
            )
            if normalizar_identificador_equipo(valor)
        }
        if codigo_normalizado and codigo_normalizado in codigos_existentes:
            return equipo
        if equipo["categoria"] == "MOVILES" and (
            imei_normalizado in identificadores_existentes
            or codigo_normalizado in identificadores_existentes
        ):
            return equipo
    return None


def buscar_equipo_por_codigo(cursor, codigo, equipo_id_excluir=None):
    codigo_limpio = (codigo or "").strip()
    if not codigo_limpio:
        return None

    consulta = """
        SELECT id, codigo_margesi, codigo_interno, categoria, descripcion
        FROM equipos
        WHERE (
            TRIM(COALESCE(codigo_margesi, '')) = ? COLLATE NOCASE
            OR TRIM(COALESCE(codigo_interno, '')) = ? COLLATE NOCASE
        )
    """
    parametros = [codigo_limpio, codigo_limpio]
    if equipo_id_excluir is not None:
        consulta += " AND id != ?"
        parametros.append(equipo_id_excluir)
    return cursor.execute(consulta + " LIMIT 1", parametros).fetchone()


DIAS_ALERTA_CERTIFICADO = 30


def validar_dni(valor):
    dni = str(valor or "").strip()
    if not re.fullmatch(r"[0-9]{8}", dni):
        raise ValueError("El DNI debe contener exactamente 8 números.")
    return dni


def validar_id_formulario(valor, etiqueta):
    texto = str(valor or "").strip()
    if not re.fullmatch(r"[0-9]{1,10}", texto) or int(texto) <= 0:
        raise ValueError(f"El identificador de {etiqueta} no es válido.")
    return int(texto)


def validar_documento_numerico(valor):
    documento = str(valor or "").strip()
    if not re.fullmatch(r"[0-9]{1,30}", documento):
        raise ValueError("El DNI/documento debe contener solo números (máximo 30 dígitos).")
    return documento


def validar_numero_contacto(valor, etiqueta):
    numero = str(valor or "").strip()
    if numero and not re.fullmatch(r"[0-9]{1,15}", numero):
        raise ValueError(f"{etiqueta} debe contener solo números (máximo 15 dígitos).")
    return numero


def validar_correo(valor, etiqueta):
    correo = str(valor or "").strip()
    if correo and (
        len(correo) > 254
        or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", correo)
    ):
        raise ValueError(f"{etiqueta} no tiene un formato válido.")
    return correo


def validar_fecha_formulario(valor, etiqueta, obligatoria=False):
    texto = str(valor or "").strip()
    if not texto:
        if obligatoria:
            raise ValueError(f"Debe ingresar {etiqueta}.")
        return ""
    if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", texto):
        raise ValueError(f"{etiqueta} no es válida. Seleccione una fecha del calendario.")
    try:
        return date.fromisoformat(texto).isoformat()
    except ValueError as error:
        raise ValueError(f"{etiqueta} no es una fecha real.") from error


def validar_hora_formulario(valor, etiqueta):
    texto = str(valor or "").strip()
    if texto and not re.fullmatch(r"(?:[01][0-9]|2[0-3]):[0-5][0-9]", texto):
        raise ValueError(f"{etiqueta} debe tener el formato HH:MM.")
    return texto


def validar_rango_fechas(desde, hasta, etiqueta="periodo"):
    desde_validada = validar_fecha_formulario(desde, f"la fecha inicial de {etiqueta}")
    hasta_validada = validar_fecha_formulario(hasta, f"la fecha final de {etiqueta}")
    if desde_validada and hasta_validada and hasta_validada < desde_validada:
        raise ValueError(f"La fecha final de {etiqueta} no puede ser anterior a la inicial.")
    return desde_validada, hasta_validada


def validar_campos_actividad(titulo, descripcion, categoria, area, prioridad, estado):
    categorias = {
        "HARDWARE", "SOFTWARE", "REDES", "TOKENS", "EQUIPOS MOVILES",
        "CAPACITACION", "GRABAR SESION", "TRAMITE UTI", "SOPORTE TECNICO",
    }
    if not titulo or len(titulo) > 300:
        raise ValueError("El título es obligatorio y no puede superar los 300 caracteres.")
    if len(descripcion) > 5000 or len(area) > 200:
        raise ValueError("La descripción o el área solicitante superan la longitud permitida.")
    if categoria not in categorias:
        raise ValueError("La categoría de actividad no es válida.")
    if prioridad not in {"ALTA", "MEDIA", "BAJA"}:
        raise ValueError("La prioridad seleccionada no es válida.")
    if estado not in {"PENDIENTE", "EN PROGRESO", "COMPLETADA"}:
        raise ValueError("El estado seleccionado no es válido.")


def normalizar_encabezado_certificado(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    return re.sub(r"[^a-z0-9]+", "_", texto.lower()).strip("_")


def fecha_certificado_iso(valor):
    if valor is None or str(valor).strip() == "":
        raise ValueError("Falta una fecha obligatoria.")
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    if isinstance(valor, (int, float)):
        try:
            return openpyxl.utils.datetime.from_excel(valor).date().isoformat()
        except (TypeError, ValueError, OverflowError) as error:
            raise ValueError("La fecha del certificado no es válida.") from error

    texto = str(valor).strip()
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(texto, formato).date().isoformat()
        except ValueError:
            continue
    raise ValueError(f"Fecha inválida: {texto}. Use YYYY-MM-DD o DD/MM/YYYY.")


def datos_certificado_desde_formulario(formulario, cursor=None, exigir_titular=False):
    titular_formulario = normalizar_encabezado_certificado(
        formulario.get("tipo_titular", "")
    )
    tipos_titular = {
        "personal": "PERSONAL",
        "trabajador": "PERSONAL",
        "comite": "COMITE",
        "miembro_del_comite": "COMITE",
        "comite_de_gestion_de_compras": "COMITE",
        "registro_previo": "",
    }
    tipo_titular = tipos_titular.get(titular_formulario, titular_formulario.upper())
    if exigir_titular and tipo_titular not in {"PERSONAL", "COMITE"}:
        raise ValueError("Indique si el titular pertenece al personal o al comité de gestión de compras.")
    trabajador = None
    if tipo_titular == "PERSONAL":
        if cursor is None:
            raise ValueError("No se pudo consultar el padrón actualizado de personal.")
        personal_id = formulario.get("personal_id", "").strip()
        if personal_id:
            if not re.fullmatch(r"[0-9]+", personal_id):
                raise ValueError("Seleccione un trabajador válido del padrón.")
            trabajador = cursor.execute(
                "SELECT * FROM trabajadores WHERE id = ? AND activo = 1",
                (int(personal_id),),
            ).fetchone()
        else:
            dni = validar_documento_numerico(formulario.get("dni"))
            trabajador = cursor.execute(
                "SELECT * FROM trabajadores WHERE dni = ? AND activo = 1", (dni,)
            ).fetchone()
        if not trabajador:
            raise ValueError("Seleccione un trabajador activo del padrón actualizado.")
    elif tipo_titular not in ("", "COMITE"):
        raise ValueError("Seleccione personal del padrón o miembro del comité.")

    if trabajador:
        dni = trabajador["dni"]
        apellido_paterno = trabajador["apellido_paterno"]
        apellido_materno = trabajador["apellido_materno"]
        nombres = trabajador["nombres"]
        cargo = trabajador["cargo"] or ""
        correo = trabajador["correo_institucional"] or trabajador["correo_personal"] or ""
        trabajador_id = trabajador["id"]
    else:
        dni = validar_documento_numerico(formulario.get("dni"))
        apellido_paterno = formulario.get("apellido_paterno", "").strip()
        apellido_materno = formulario.get("apellido_materno", "").strip()
        nombres = formulario.get("nombres", "").strip()
        cargo = formulario.get("cargo", "").strip()
        correo = formulario.get("correo", "").strip()
        trabajador_id = None

    datos = {
        "dni": dni,
        "apellido_paterno": apellido_paterno,
        "apellido_materno": apellido_materno,
        "nombres": nombres,
        "cargo": cargo,
        "oficina": formulario.get("oficina", "").strip(),
        "correo": correo,
        "tipo_certificado": formulario.get("tipo_certificado", "").strip().upper(),
        "trabajador_id": trabajador_id,
        "numero_resolucion": formulario.get("numero_resolucion", "").strip(),
        "fecha_inicio": validar_fecha_formulario(
            formulario.get("fecha_inicio"), "la fecha de inicio", obligatoria=True
        ),
        "fecha_caducidad": validar_fecha_formulario(
            formulario.get("fecha_caducidad"), "la fecha de caducidad", obligatoria=True
        ),
    }
    obligatorios = (
        "dni", "apellido_paterno", "apellido_materno", "nombres", "tipo_certificado"
    )
    if any(not datos[campo] for campo in obligatorios):
        raise ValueError("Complete el DNI/documento, nombres, ambos apellidos y tipo de certificado.")
    tipo_certificado = normalizar_encabezado_certificado(datos["tipo_certificado"])
    tipos_certificado = {
        "token": "TOKEN",
        "certificado_digital_token": "TOKEN",
        "para_pc": "PARA PC",
        "certificado_digital_para_pc": "PARA PC",
    }
    datos["tipo_certificado"] = tipos_certificado.get(tipo_certificado, "")
    if not datos["tipo_certificado"]:
        raise ValueError("El tipo de certificado debe ser TOKEN o PARA PC.")
    if tipo_titular == "COMITE" and not datos["numero_resolucion"]:
        raise ValueError("Ingrese el número de resolución del miembro del comité.")
    if tipo_titular == "COMITE" and not datos["cargo"]:
        raise ValueError("Ingrese el cargo del miembro del comité.")
    limites = {
        "apellido_paterno": 100,
        "apellido_materno": 100,
        "nombres": 120,
        "cargo": 150,
        "oficina": 150,
        "correo": 254,
        "tipo_certificado": 100,
        "numero_resolucion": 100,
    }
    if any(len(datos[campo]) > limite for campo, limite in limites.items()):
        raise ValueError("Uno o más campos del certificado superan la longitud permitida.")
    datos["correo"] = validar_correo(datos["correo"], "El correo")
    if datos["fecha_caducidad"] < datos["fecha_inicio"]:
        raise ValueError("La fecha de caducidad no puede ser anterior a la fecha de inicio.")
    return datos


def buscar_certificado_duplicado(cursor, datos, certificado_id=None):
    consulta = """
        SELECT id FROM certificados_digitales
        WHERE dni = ? AND tipo_certificado = ?
          AND fecha_inicio = ? AND fecha_caducidad = ?
    """
    parametros = [
        datos["dni"], datos["tipo_certificado"],
        datos["fecha_inicio"], datos["fecha_caducidad"],
    ]
    if certificado_id is not None:
        consulta += " AND id != ?"
        parametros.append(certificado_id)
    return cursor.execute(consulta, parametros).fetchone()


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
    if request.method != "POST" or request.endpoint in {"static", "login", "sincronizar_grabacion_apowerrec"}:
        return None

    token_formulario = (
        request.form.get("csrf_token", "")
        or request.headers.get("X-CSRFToken", "")
        or request.headers.get("X-CSRF-Token", "")
    )
    token_sesion = session.get("csrf_token", "")
    if not token_sesion or not token_formulario or not hmac.compare_digest(token_formulario, token_sesion):
        if request.is_json or request.headers.get("X-Requested-With") == "XMLHttpRequest":
            return jsonify(error="Solicitud no válida o sesión expirada."), 400

        obtener_token_csrf()
        flash("La sesión o solicitud de formulario expiró. Por favor vuelva a enviar el formulario.", "warning")
        return redirect(request.referrer or url_for("inicio"))
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
    "tokens": "Gestión de certificados digitales",
    "actividades": "Actividades TI",
    "documentaria": "Gestión documentaria",
    "sgd": "Monitor SGD",
    "reportes": "Reportes",
    "biblioteca": "Biblioteca Virtual PAE",
    "cargos_asignaciones": "Cargos y sistemas asignados",
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

def codigo_control_actividad(actividad_id):
    return f"ACT-{int(actividad_id):06d}"

@app.context_processor
def utility_processor():
    return dict(
        calcular_duracion=calcular_duracion,
        codigo_control_actividad=codigo_control_actividad,
        usuario_tiene_permiso=usuario_tiene_permiso,
        csrf_token=obtener_token_csrf,
    )

NIVELES_ACCESO_UT = {
    "CUENTA": "Solo cuenta (sin acceso funcional)",
    "LECTURA": "Consulta",
    "REGISTRO": "Registro",
    "EDICION": "Edición",
    "ADMINISTRACION": "Administración",
}
NIVELES_SOLICITUD_UT = {
    "CONSULTA": "Consulta",
    "REGISTRO": "Registro",
    "MODIFICACION": "Modificación",
    "APROBACION": "Aprobación",
    "REPORTES": "Reportes",
}
ACCIONES_SOLICITUD_UT = {
    "MANTENER": "Mantener",
    "MODIFICAR": "Modificar",
    "NUEVO": "Nuevo",
    "RETIRAR": "Retirar",
}

def _obtener_conexion_matriz_ut():
    conn = get_db()
    ensure_access_matrix_schema(conn)
    return conn

ORDEN_SISTEMAS_UTI = {
    nombre: indice
    for indice, nombre in enumerate((
        "SIGO", "SIGDEL", "SGD", "SADE", "CORREO ELECTRONICO", "MESA DE AYUDA",
        "APK - APLICATIVOS MOVILES", "APK - APLICATIVOS MOBILES",
        "SIGA", "SIAF CONSULTA"
    ))
}

INFO_SISTEMAS_CATALOGO_UT = {
    "SIGO": {
        "icono": "bi-diagram-3-fill",
        "icono_bg": "ut-bg-sigo",
        "color": "#2563eb",
        "descripcion": "Sistema Integrado de Gestión Operativa",
        "categoria": "Operaciones y Gestión",
    },
    "SIGDEL": {
        "icono": "bi-truck",
        "icono_bg": "ut-bg-sigdel",
        "color": "#7c3aed",
        "descripcion": "Sistema de Gestión de Distribución",
        "categoria": "Distribución y Entregas",
    },
    "SGD": {
        "icono": "bi-file-earmark-text-fill",
        "icono_bg": "ut-bg-sgd",
        "color": "#059669",
        "descripcion": "Sistema de Gestión Documental",
        "categoria": "Trámite Documentario",
    },
    "SADE": {
        "icono": "bi-building-fill-gear",
        "icono_bg": "ut-bg-sade",
        "color": "#d97706",
        "descripcion": "Sistema de Administración",
        "categoria": "Administración",
    },
    "CORREO ELECTRONICO": {
        "icono": "bi-envelope-at-fill",
        "icono_bg": "ut-bg-correo",
        "color": "#dc2626",
        "descripcion": "Correo Institucional",
        "categoria": "Comunicación Oficial",
    },
    "MESA DE AYUDA": {
        "icono": "bi-headset",
        "icono_bg": "ut-bg-mesa",
        "color": "#0891b2",
        "descripcion": "Mesa de Ayuda y Soporte TI",
        "categoria": "Soporte Técnico UTI",
    },
    "APK - APLICATIVOS MOVILES": {
        "icono": "bi-phone-fill",
        "icono_bg": "ut-bg-apk",
        "color": "#16a34a",
        "descripcion": "Aplicativos móviles institucionales del PAE",
        "categoria": "Aplicativos Móviles (APK)",
    },
    "APK - APLICATIVOS MOBILES": {
        "icono": "bi-phone-fill",
        "icono_bg": "ut-bg-apk",
        "color": "#16a34a",
        "descripcion": "Aplicativos móviles institucionales del PAE",
        "categoria": "Aplicativos Móviles (APK)",
    },
    "APLICATIVO MOVIL: PAE IIEE": {
        "icono": "bi-phone-fill",
        "icono_bg": "ut-bg-apk",
        "color": "#16a34a",
        "descripcion": "Supervisión PAE en IIEE",
        "categoria": "Aplicativo Móvil APK",
    },
    "APLICATIVO MOVIL: PAE ESTABLECIMIENTO": {
        "icono": "bi-phone-fill",
        "icono_bg": "ut-bg-apk-est",
        "color": "#4f46e5",
        "descripcion": "Supervisión en Establecimiento",
        "categoria": "Aplicativo Móvil APK",
    },
    "APLICATIVO MOVIL: ALMACEN EN IE": {
        "icono": "bi-phone-fill",
        "icono_bg": "ut-bg-apk-alm",
        "color": "#ea580c",
        "descripcion": "Control de Almacén en IE",
        "categoria": "Aplicativo Móvil APK",
    },
    "APLICATIVO MOVIL: IIEE": {
        "icono": "bi-phone-fill",
        "icono_bg": "ut-bg-apk-iiee",
        "color": "#0d9488",
        "descripcion": "Aplicativo Móvil IIEE",
        "categoria": "Aplicativo Móvil APK",
    },
    "SIGA": {
        "icono": "bi-journal-check",
        "icono_bg": "ut-bg-siga",
        "color": "#9333ea",
        "descripcion": "Sistema Integrado de Gestión Administrativa",
        "categoria": "Gestión MEF",
    },
    "SIAF CONSULTA": {
        "icono": "bi-bank2",
        "icono_bg": "ut-bg-siaf",
        "color": "#0284c7",
        "descripcion": "Sistema de Administración Financiera (Consulta)",
        "categoria": "Finanzas MEF",
    },
}

def _decorar_sistema_ut(sistema):
    nombre_limpio = sistema["nombre"].upper().strip()
    info = INFO_SISTEMAS_CATALOGO_UT.get(nombre_limpio, {
        "icono": "bi-app-indicator",
        "icono_bg": "ut-bg-default",
        "color": "#475569",
        "descripcion": "Sistema / Aplicativo UTI",
        "categoria": sistema.get("tipo", "SISTEMA"),
    })
    sistema["icono"] = info["icono"]
    sistema["icono_bg"] = info["icono_bg"]
    sistema["color"] = info["color"]
    sistema["descripcion"] = info["descripcion"]
    sistema["categoria"] = info["categoria"]
    return sistema

def _construir_arbol_modulos_ut(modulos):
    nodos = {modulo["id"]: {**modulo, "hijos": []} for modulo in modulos}
    raices = []
    for modulo in modulos:
        nodo = nodos[modulo["id"]]
        padre = nodos.get(modulo["padre_id"])
        if padre:
            padre["hijos"].append(nodo)
        else:
            raices.append(nodo)

    orden_tipos = {"SISTEMA": 0, "SERVICIO": 1, "APK": 2, "MODULO": 3, "SUBMODULO": 4}

    def contar_descendientes(nodo):
        nodo["cantidad_descendientes"] = sum(
            1 + contar_descendientes(hijo) for hijo in nodo["hijos"]
        )
        return nodo["cantidad_descendientes"]

    for nodo in nodos.values():
        nodo["hijos"].sort(key=lambda hijo: hijo["id"])
        contar_descendientes(nodo)
    raices.sort(key=lambda nodo: (
        ORDEN_SISTEMAS_UTI.get(nodo["nombre"].upper(), 100 + orden_tipos.get(nodo["tipo"], 4)),
        nodo["nombre"].casefold(),
    ))
    return raices

def _catalogo_solicitud_acceso_ut(cursor):
    modulos = [
        dict(row) for row in cursor.execute("""
            SELECT id, nombre, tipo, padre_id, obligatorio_todos
            FROM sistemas_modulos_ut
            WHERE activo = 1
        """).fetchall()
    ]
    por_id = {modulo["id"]: modulo for modulo in modulos}
    for modulo in modulos:
        partes = [modulo["nombre"]]
        padre_id = modulo["padre_id"]
        visitados = {modulo["id"]}
        while padre_id in por_id and padre_id not in visitados:
            padre = por_id[padre_id]
            partes.append(padre["nombre"])
            visitados.add(padre_id)
            padre_id = padre["padre_id"]
        modulo["ruta"] = " / ".join(reversed(partes))
    return sorted(modulos, key=lambda modulo: modulo["ruta"].casefold())

def _etiqueta_nivel_acceso_ut(nivel):
    return NIVELES_ACCESO_UT.get(
        nivel,
        NIVELES_SOLICITUD_UT.get(nivel, nivel),
    )

def _obtener_niveles_perfil_cargo_ut(cursor, perfil_id):
    niveles = {
        row["id"]: "CUENTA"
        for row in cursor.execute("""
            SELECT id
            FROM sistemas_modulos_ut
            WHERE activo = 1 AND obligatorio_todos = 1
        """).fetchall()
    }
    niveles.update({
        row["sistema_modulo_id"]: row["nivel_acceso"]
        for row in cursor.execute("""
            SELECT a.sistema_modulo_id, a.nivel_acceso
            FROM accesos_perfil_cargo_ut a
            JOIN sistemas_modulos_ut m
              ON m.id = a.sistema_modulo_id AND m.activo = 1
            WHERE a.perfil_id = ?
        """, (perfil_id,)).fetchall()
    })
    return niveles

def _registrar_snapshot_accesos_trabajador_ut(cursor, trabajador_id, perfil, usuario_id):
    niveles = _obtener_niveles_perfil_cargo_ut(cursor, perfil["id"])
    if not niveles:
        raise ValueError(
            f"El perfil del cargo «{perfil['cargo']}» no tiene sistemas activos asignados."
        )
    cursor.execute("""
        INSERT INTO solicitudes_acceso_trabajador_ut
            (trabajador_id, funciones, nivel_acceso, reportes_requeridos,
             accion, autorizador, usuario_id)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        trabajador_id,
        perfil["funciones"],
        "CONSULTA",
        perfil["reportes_requeridos"],
        perfil["accion_acceso"],
        perfil["autorizador"],
        usuario_id,
    ))
    solicitud_id = cursor.lastrowid
    cursor.executemany("""
        INSERT INTO modulos_solicitud_acceso_trabajador_ut
            (solicitud_id, sistema_modulo_id, nivel_acceso)
        VALUES (?, ?, ?)
    """, [
        (solicitud_id, modulo_id, nivel)
        for modulo_id, nivel in sorted(niveles.items())
    ])
    return solicitud_id

def _obtener_perfiles_cargo_validados(cursor):
    modulos = _catalogo_solicitud_acceso_ut(cursor)
    modulos_activos = {modulo["id"]: modulo for modulo in modulos}
    accesos_por_perfil = {}
    for acceso in cursor.execute("""
        SELECT a.perfil_id, a.sistema_modulo_id, a.nivel_acceso
        FROM accesos_perfil_cargo_ut a
        JOIN sistemas_modulos_ut m
          ON m.id = a.sistema_modulo_id AND m.activo = 1
    """).fetchall():
        accesos_por_perfil.setdefault(acceso["perfil_id"], {})[
            acceso["sistema_modulo_id"]
        ] = acceso["nivel_acceso"]

    perfiles = []
    for fila in cursor.execute("""
        SELECT id, cargo, funciones, autorizador, reportes_requeridos, accion_acceso
        FROM perfiles_cargo_ut
        WHERE activo = 1 AND validado = 1
          AND TRIM(funciones) != '' AND TRIM(autorizador) != ''
          AND TRIM(reportes_requeridos) != ''
        ORDER BY cargo COLLATE NOCASE
    """).fetchall():
        niveles = {
            modulo["id"]: "CUENTA"
            for modulo in modulos
            if modulo["obligatorio_todos"]
        }
        niveles.update(accesos_por_perfil.get(fila["id"], {}))
        asignaciones = [
            {
                "id": modulo_id,
                "ruta": modulos_activos[modulo_id]["ruta"],
                "tipo": modulos_activos[modulo_id]["tipo"],
                "nivel": nivel,
                "nivel_nombre": _etiqueta_nivel_acceso_ut(nivel),
            }
            for modulo_id, nivel in niveles.items()
            if modulo_id in modulos_activos
        ]
        asignaciones.sort(key=lambda item: item["ruta"].casefold())
        perfiles.append({
            "id": fila["id"],
            "cargo": fila["cargo"],
            "funciones": fila["funciones"],
            "autorizador": fila["autorizador"],
            "reportes_requeridos": fila["reportes_requeridos"],
            "accion_acceso": fila["accion_acceso"],
            "asignaciones": asignaciones,
        })
    return perfiles

@app.route("/cargos-asignaciones")
@permiso_requerido("cargos_asignaciones", "ver")
@login_required
def cargos_asignaciones_ut():
    conn = _obtener_conexion_matriz_ut()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR IGNORE INTO perfiles_cargo_ut (cargo)
        SELECT DISTINCT TRIM(cargo)
        FROM trabajadores
        WHERE cargo IS NOT NULL AND TRIM(cargo) != ''
    """)

    modulos_catalogo = [
        dict(row) for row in cursor.execute("""
            SELECT m.*, p.nombre AS padre_nombre
            FROM sistemas_modulos_ut m
            LEFT JOIN sistemas_modulos_ut p ON p.id = m.padre_id
            ORDER BY m.activo DESC, m.nombre COLLATE NOCASE
        """).fetchall()
    ]
    modulos = [item for item in modulos_catalogo if item["activo"]]
    modulos_por_id = {item["id"]: item for item in modulos_catalogo}
    for item in modulos_catalogo:
        camino = [item["nombre"]]
        padre_id = item["padre_id"]
        visitados = {item["id"]}
        while padre_id in modulos_por_id and padre_id not in visitados:
            padre = modulos_por_id[padre_id]
            camino.append(padre["nombre"])
            visitados.add(padre_id)
            padre_id = padre["padre_id"]
        item["ruta"] = " / ".join(reversed(camino))
    arbol_catalogo = _construir_arbol_modulos_ut(modulos_catalogo)
    arbol_modulos = _construir_arbol_modulos_ut(modulos)

    perfiles = [
        dict(row) for row in cursor.execute("""
            SELECT p.*,
                   COUNT(DISTINCT t.id) AS trabajadores_asignados,
                   COUNT(DISTINCT m.id) AS modulos_asignados
            FROM perfiles_cargo_ut p
            LEFT JOIN trabajadores t
                ON LOWER(TRIM(COALESCE(t.cargo, ''))) = LOWER(TRIM(p.cargo))
               AND t.activo = 1
            LEFT JOIN accesos_perfil_cargo_ut a ON a.perfil_id = p.id
            LEFT JOIN sistemas_modulos_ut m
                ON m.id = a.sistema_modulo_id AND m.activo = 1
            WHERE p.activo = 1
            GROUP BY p.id
            ORDER BY p.cargo COLLATE NOCASE
        """).fetchall()
    ]
    accesos_por_perfil = {}
    for acceso in cursor.execute("""
        SELECT perfil_id, sistema_modulo_id, nivel_acceso
        FROM accesos_perfil_cargo_ut
    """).fetchall():
        accesos_por_perfil.setdefault(acceso["perfil_id"], {})[
            acceso["sistema_modulo_id"]
        ] = acceso["nivel_acceso"]

    reporte_trabajadores = _obtener_reporte_accesos_ut(cursor)
    historial = cursor.execute("""
        SELECT h.*, u.nombre_completo AS usuario_nombre
        FROM historial_perfiles_cargo_ut h
        LEFT JOIN usuarios_sistema u ON u.id = h.usuario_id
        ORDER BY h.registrado_el DESC, h.id DESC
        LIMIT 30
    """).fetchall()

    sistemas_catalogo = []
    for item in arbol_catalogo:
        if item.get("padre_id") is None and item.get("activo") and item.get("tipo") in {"SISTEMA", "SERVICIO"}:
            if not item.get("nombre", "").upper().startswith("APLICATIVO MOVIL:"):
                sistemas_catalogo.append(_decorar_sistema_ut(dict(item)))

    sistemas_catalogo.sort(key=lambda s: (
        ORDEN_SISTEMAS_UTI.get(s["nombre"].upper(), 999),
        s["nombre"].casefold()
    ))

    conn.commit()
    conn.close()
    return render_template(
        "cargos_asignaciones.html",
        modulos=modulos,
        modulos_catalogo=modulos_catalogo,
        arbol_catalogo=arbol_catalogo,
        arbol_modulos=arbol_modulos,
        sistemas_catalogo=sistemas_catalogo,
        tab_activo=request.args.get("tab", "catalogo"),
        perfiles=perfiles,
        accesos_por_perfil=accesos_por_perfil,
        reporte_trabajadores=reporte_trabajadores,
        historial=historial,
        niveles_acceso=NIVELES_ACCESO_UT,
        total_trabajadores=len(reporte_trabajadores),
    )

@app.route("/cargos-asignaciones/catalogo")
@permiso_requerido("cargos_asignaciones", "ver")
@login_required
def catalogo_sistemas_ut():
    return redirect(url_for("cargos_asignaciones_ut", tab="catalogo"))

def _obtener_reporte_accesos_ut(cursor):
    modulos = {
        row["id"]: dict(row)
        for row in cursor.execute("""
            SELECT id, nombre, tipo, padre_id, obligatorio_todos
            FROM sistemas_modulos_ut
            WHERE activo = 1
        """).fetchall()
    }
    accesos_perfil = {}
    for row in cursor.execute("""
        SELECT a.perfil_id, a.sistema_modulo_id, a.nivel_acceso
        FROM accesos_perfil_cargo_ut a
        JOIN sistemas_modulos_ut m ON m.id = a.sistema_modulo_id AND m.activo = 1
    """).fetchall():
        accesos_perfil.setdefault(row["perfil_id"], {})[
            row["sistema_modulo_id"]
        ] = row["nivel_acceso"]

    todos_modulos = {
        row["id"]: dict(row)
        for row in cursor.execute("""
            SELECT id, nombre, tipo, padre_id, activo
            FROM sistemas_modulos_ut
        """).fetchall()
    }
    solicitudes = {
        row["trabajador_id"]: dict(row)
        for row in cursor.execute("""
            SELECT s.*
            FROM solicitudes_acceso_trabajador_ut s
            WHERE s.id = (
                SELECT MAX(s2.id)
                FROM solicitudes_acceso_trabajador_ut s2
                WHERE s2.trabajador_id = s.trabajador_id
            )
        """).fetchall()
    }
    if solicitudes:
        solicitudes_por_id = {
            solicitud["id"]: solicitud for solicitud in solicitudes.values()
        }
        for row in cursor.execute("""
            SELECT a.solicitud_id, m.id, m.nombre, m.tipo, m.padre_id, m.activo,
                   a.nivel_acceso
            FROM modulos_solicitud_acceso_trabajador_ut a
            JOIN sistemas_modulos_ut m ON m.id = a.sistema_modulo_id
            WHERE a.solicitud_id IN (
                SELECT MAX(s.id)
                FROM solicitudes_acceso_trabajador_ut s
                GROUP BY s.trabajador_id
            )
        """).fetchall():
            solicitud = solicitudes_por_id[row["solicitud_id"]]
            solicitud.setdefault("modulos", []).append(dict(row))

    trabajadores = cursor.execute("""
        SELECT t.id, t.nombres, t.apellido_paterno, t.apellido_materno, t.cargo,
               p.id AS perfil_id, p.funciones, p.autorizador, p.validado
        FROM trabajadores t
        LEFT JOIN perfiles_cargo_ut p
            ON LOWER(TRIM(COALESCE(t.cargo, ''))) = LOWER(TRIM(p.cargo))
           AND p.activo = 1
        WHERE t.activo = 1
        ORDER BY t.apellido_paterno COLLATE NOCASE,
                 t.apellido_materno COLLATE NOCASE, t.nombres COLLATE NOCASE
    """).fetchall()

    def ruta_modulo(modulo_id):
        partes = []
        visitados = set()
        actual_id = modulo_id
        while actual_id in todos_modulos and actual_id not in visitados:
            actual = todos_modulos[actual_id]
            partes.append(actual["nombre"])
            visitados.add(actual_id)
            actual_id = actual["padre_id"]
        return " / ".join(reversed(partes))

    resultado = []
    for trabajador in trabajadores:
        niveles = {
            item["id"]: "CUENTA"
            for item in modulos.values()
            if item["obligatorio_todos"]
        }
        perfil_id = trabajador["perfil_id"]
        if perfil_id:
            niveles.update(accesos_perfil.get(perfil_id, {}))
        accesos = [
            {
                "ruta": ruta_modulo(modulo_id),
                "nivel": nivel,
                "nivel_nombre": NIVELES_ACCESO_UT[nivel],
            }
            for modulo_id, nivel in niveles.items()
            if modulo_id in modulos
        ]
        accesos.sort(key=lambda acceso: acceso["ruta"].casefold())
        solicitud = solicitudes.get(trabajador["id"])
        if solicitud:
            solicitud["nivel_nombre"] = NIVELES_SOLICITUD_UT[solicitud["nivel_acceso"]]
            solicitud["accion_nombre"] = ACCIONES_SOLICITUD_UT[solicitud["accion"]]
            solicitud.setdefault("modulos", [])
            solicitud["modulos"].sort(
                key=lambda modulo: ruta_modulo(modulo["id"]).casefold()
            )
            for modulo in solicitud["modulos"]:
                modulo["ruta"] = ruta_modulo(modulo["id"])
                modulo["nivel_nombre"] = _etiqueta_nivel_acceso_ut(
                    modulo["nivel_acceso"]
                )
                modulo["estado_catalogo"] = (
                    "ACTIVO" if modulo["activo"] else "INACTIVO"
                )
        resultado.append({
            "id": trabajador["id"],
            "nombre": " ".join(
                valor for valor in (
                    trabajador["nombres"],
                    trabajador["apellido_paterno"],
                    trabajador["apellido_materno"],
                ) if valor
            ),
            "cargo": trabajador["cargo"] or "SIN CARGO REGISTRADO",
            "funciones": trabajador["funciones"] or "",
            "autorizador": trabajador["autorizador"] or "",
            "estado_perfil": (
                "SIN PERFIL DE CARGO" if not perfil_id else
                "CONFIGURADO" if trabajador["validado"] else
                "PENDIENTE DE VALIDACIÓN"
            ),
            "accesos": accesos,
            "solicitud": solicitud,
        })
    return resultado

def _redir_modulo_ut(form):
    redirigir_sistema = form.get("redirigir_sistema", "").strip()
    if re.fullmatch(r"[0-9]+", redirigir_sistema):
        return redirect(url_for("sistema_detalle_ut", sistema_id=int(redirigir_sistema)))
    if form.get("redirigir") == "catalogo":
        return redirect(url_for("cargos_asignaciones_ut", tab="catalogo"))
    return redirect(url_for("cargos_asignaciones_ut"))

@app.route("/cargos-asignaciones/sistemas/<int:sistema_id>")
@permiso_requerido("cargos_asignaciones", "ver")
@login_required
def sistema_detalle_ut(sistema_id):
    conn = _obtener_conexion_matriz_ut()
    cursor = conn.cursor()
    sistema_row = cursor.execute("SELECT * FROM sistemas_modulos_ut WHERE id = ?", (sistema_id,)).fetchone()
    if not sistema_row:
        conn.close()
        flash("El sistema solicitado no existe en el catálogo.", "warning")
        return redirect(url_for("cargos_asignaciones_ut"))

    sistema = _decorar_sistema_ut(dict(sistema_row))

    modulos_raw = [
        dict(row) for row in cursor.execute("""
            SELECT * FROM sistemas_modulos_ut ORDER BY activo DESC, nombre COLLATE NOCASE
        """).fetchall()
    ]
    modulos_por_id = {item["id"]: item for item in modulos_raw}
    for item in modulos_raw:
        camino = [item["nombre"]]
        p_id = item["padre_id"]
        visitados = {item["id"]}
        while p_id in modulos_por_id and p_id not in visitados:
            padre = modulos_por_id[p_id]
            camino.append(padre["nombre"])
            visitados.add(p_id)
            p_id = padre["padre_id"]
        item["ruta"] = " / ".join(reversed(camino))

    # Identificar todos los descendientes de este sistema
    def es_descendiente(item_id, root_id):
        curr_id = item_id
        visitados = set()
        while curr_id in modulos_por_id and curr_id not in visitados:
            visitados.add(curr_id)
            curr = modulos_por_id[curr_id]
            if curr["padre_id"] == root_id:
                return True
            curr_id = curr["padre_id"]
        return False

    descendientes = [_decorar_sistema_ut(dict(m)) for m in modulos_raw if es_descendiente(m["id"], sistema_id)]

    nodos = {m["id"]: {**m, "hijos": []} for m in descendientes}
    arbol_hijos = []
    for m in descendientes:
        nodo = nodos[m["id"]]
        if m["padre_id"] == sistema_id:
            arbol_hijos.append(nodo)
        elif m["padre_id"] in nodos:
            nodos[m["padre_id"]]["hijos"].append(nodo)

    def contar(n):
        n["cantidad_descendientes"] = sum(1 + contar(h) for h in n["hijos"])
        return n["cantidad_descendientes"]

    for n in nodos.values():
        n["hijos"].sort(key=lambda h: h["id"])
        contar(n)
    arbol_hijos.sort(key=lambda n: n["nombre"].casefold())

    modulos_del_sistema = [
        m for m in descendientes if m["activo"] and m["tipo"] in {"MODULO", "SUBMODULO", "APK"}
    ]
    modulos_del_sistema.sort(key=lambda m: m["ruta"].casefold())

    total_descendientes = len(descendientes)
    conn.close()

    return render_template(
        "sistema_detalle.html",
        sistema=sistema,
        arbol_hijos=arbol_hijos,
        modulos_del_sistema=modulos_del_sistema,
        total_descendientes=total_descendientes,
    )

@app.route("/cargos-asignaciones/modulos/nuevo", methods=["POST"])
@permiso_requerido("cargos_asignaciones", "modificar")
@login_required
def agregar_modulo_ut():
    nombre = request.form.get("nombre", "").strip()
    tipo = request.form.get("tipo", "").strip().upper()
    padre_raw = request.form.get("padre_id", "").strip()
    obligatorio_todos = 1 if request.form.get("obligatorio_todos") == "on" else 0
    if not nombre or len(nombre) > 120 or tipo not in {"SISTEMA", "MODULO", "SUBMODULO", "APK", "SERVICIO"}:
        flash("Ingrese un nombre válido y un tipo de sistema, módulo, submódulo, APK o servicio.", "warning")
        return _redir_modulo_ut(request.form)
    padre_id = None
    if padre_raw:
        try:
            padre_id = validar_id_formulario(padre_raw, "sistema padre")
        except ValueError:
            flash("El sistema padre seleccionado no es válido.", "warning")
            return _redir_modulo_ut(request.form)
        conn = _obtener_conexion_matriz_ut()
        existe_padre = conn.execute(
            "SELECT tipo FROM sistemas_modulos_ut WHERE id = ? AND activo = 1",
            (padre_id,),
        ).fetchone()
        if not existe_padre:
            conn.close()
            flash("El sistema padre seleccionado no existe o está inactivo.", "warning")
            return _redir_modulo_ut(request.form)
        if existe_padre["tipo"] not in {"SISTEMA", "MODULO", "SUBMODULO"} or tipo == "SISTEMA":
            conn.close()
            flash("Seleccione un sistema, módulo o submódulo como padre y un tipo de elemento válido.", "warning")
            return _redir_modulo_ut(request.form)
    else:
        conn = _obtener_conexion_matriz_ut()

    cursor = conn.cursor()
    duplicado = cursor.execute(
        "SELECT 1 FROM sistemas_modulos_ut WHERE nombre = ? COLLATE NOCASE",
        (nombre,),
    ).fetchone()
    if duplicado:
        conn.close()
        flash("Ya existe un sistema o módulo con ese nombre.", "warning")
        return _redir_modulo_ut(request.form)
    cursor.execute("""
        INSERT INTO sistemas_modulos_ut (nombre, tipo, padre_id, obligatorio_todos)
        VALUES (?, ?, ?, ?)
    """, (nombre, tipo, padre_id, obligatorio_todos))
    conn.commit()
    conn.close()
    flash("Elemento agregado al catálogo.", "success")
    return _redir_modulo_ut(request.form)

@app.route("/cargos-asignaciones/modulos/<int:modulo_id>/editar", methods=["POST"])
@permiso_requerido("cargos_asignaciones", "modificar")
@login_required
def editar_modulo_ut(modulo_id):
    nombre = request.form.get("nombre", "").strip()
    tipo = request.form.get("tipo", "").strip().upper()
    padre_raw = request.form.get("padre_id", "").strip()
    obligatorio_todos = 1 if request.form.get("obligatorio_todos") == "on" else 0
    if not nombre or len(nombre) > 120 or tipo not in {"SISTEMA", "MODULO", "SUBMODULO", "APK", "SERVICIO"}:
        flash("Ingrese un nombre válido y un tipo de sistema, módulo, submódulo, APK o servicio.", "warning")
        return _redir_modulo_ut(request.form)

    conn = _obtener_conexion_matriz_ut()
    cursor = conn.cursor()
    modulo = cursor.execute(
        "SELECT id, tipo FROM sistemas_modulos_ut WHERE id = ?",
        (modulo_id,),
    ).fetchone()
    if not modulo:
        conn.close()
        flash("El elemento que intenta modificar ya no existe.", "warning")
        return _redir_modulo_ut(request.form)
    if cursor.execute(
        "SELECT 1 FROM sistemas_modulos_ut WHERE nombre = ? COLLATE NOCASE AND id != ?",
        (nombre, modulo_id),
    ).fetchone():
        conn.close()
        flash("Ya existe otro elemento del catálogo con ese nombre.", "warning")
        return _redir_modulo_ut(request.form)
    if obligatorio_todos and not modulo["tipo"] == "SERVICIO" and tipo not in {"SISTEMA", "SERVICIO"}:
        conn.close()
        flash("Solo un sistema o servicio puede marcarse como requerido para todo el personal.", "warning")
        return _redir_modulo_ut(request.form)

    padre_id = None
    if padre_raw:
        try:
            padre_id = validar_id_formulario(padre_raw, "sistema o módulo padre")
        except ValueError:
            conn.close()
            flash("El sistema o módulo padre seleccionado no es válido.", "warning")
            return _redir_modulo_ut(request.form)
        padre = cursor.execute(
            "SELECT tipo, activo FROM sistemas_modulos_ut WHERE id = ?",
            (padre_id,),
        ).fetchone()
        if not padre or not padre["activo"] or padre["tipo"] not in {"SISTEMA", "MODULO", "SUBMODULO"}:
            conn.close()
            flash("El elemento padre debe ser un sistema, módulo o submódulo activo.", "warning")
            return _redir_modulo_ut(request.form)
        if padre_id == modulo_id or tipo == "SISTEMA":
            conn.close()
            flash("La jerarquía seleccionada no es válida.", "warning")
            return _redir_modulo_ut(request.form)

        ascendientes = {modulo_id}
        actual_id = padre_id
        while actual_id:
            if actual_id in ascendientes:
                conn.close()
                flash("No se puede mover el elemento dentro de su propia jerarquía.", "warning")
                return _redir_modulo_ut(request.form)
            ascendientes.add(actual_id)
            actual = cursor.execute(
                "SELECT padre_id FROM sistemas_modulos_ut WHERE id = ?",
                (actual_id,),
            ).fetchone()
            actual_id = actual["padre_id"] if actual else None

    tiene_hijos = cursor.execute(
        "SELECT 1 FROM sistemas_modulos_ut WHERE padre_id = ? LIMIT 1",
        (modulo_id,),
    ).fetchone()
    if tiene_hijos and tipo not in {"SISTEMA", "MODULO", "SUBMODULO"}:
        conn.close()
        flash("Un elemento con módulos dependientes debe conservar el tipo Sistema, Módulo o Submódulo.", "warning")
        return _redir_modulo_ut(request.form)

    cursor.execute("""
        UPDATE sistemas_modulos_ut
        SET nombre = ?, tipo = ?, padre_id = ?, obligatorio_todos = ?
        WHERE id = ?
    """, (nombre, tipo, padre_id, obligatorio_todos, modulo_id))
    conn.commit()
    conn.close()
    flash("Elemento del catálogo actualizado.", "success")
    return _redir_modulo_ut(request.form)

@app.route("/cargos-asignaciones/modulos/<int:modulo_id>/eliminar", methods=["POST"])
@permiso_requerido("cargos_asignaciones", "modificar")
@login_required
def eliminar_modulo_ut(modulo_id):
    conn = _obtener_conexion_matriz_ut()
    cursor = conn.cursor()
    modulo = cursor.execute(
        "SELECT id, nombre, obligatorio_todos FROM sistemas_modulos_ut WHERE id = ?",
        (modulo_id,),
    ).fetchone()
    if not modulo:
        conn.close()
        flash("El elemento que intenta eliminar ya no existe.", "warning")
        return _redir_modulo_ut(request.form)
    if modulo["obligatorio_todos"]:
        conn.close()
        flash("No se puede eliminar un elemento obligatorio para todo el personal. Primero edítelo para quitar la obligatoriedad si desea retirarlo.", "warning")
        return _redir_modulo_ut(request.form)

    # Obtener el elemento y todos sus descendientes recursivamente
    ids_a_eliminar = [modulo_id]
    por_revisar = [modulo_id]
    while por_revisar:
        actual_id = por_revisar.pop(0)
        hijos = [r["id"] for r in cursor.execute("SELECT id FROM sistemas_modulos_ut WHERE padre_id = ?", (actual_id,)).fetchall()]
        ids_a_eliminar.extend(hijos)
        por_revisar.extend(hijos)

    placeholders = ",".join("?" * len(ids_a_eliminar))
    tiene_asignaciones = cursor.execute(f"""
        SELECT 1 FROM accesos_perfil_cargo_ut WHERE sistema_modulo_id IN ({placeholders}) LIMIT 1
    """, ids_a_eliminar).fetchone() or cursor.execute(f"""
        SELECT 1 FROM modulos_solicitud_acceso_trabajador_ut WHERE sistema_modulo_id IN ({placeholders}) LIMIT 1
    """, ids_a_eliminar).fetchone()

    if tiene_asignaciones:
        conn.close()
        flash("Este elemento (o sus submódulos) tiene asignaciones en perfiles de cargo o historial asociado. Desactívelo para conservar los registros históricos.", "warning")
        return _redir_modulo_ut(request.form)

    # Eliminar en orden inverso (hojas primero, luego padre) para respetar integridad referencial
    for item_id in reversed(ids_a_eliminar):
        cursor.execute("DELETE FROM sistemas_modulos_ut WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()
    flash(f"Se eliminó «{modulo['nombre']}» del catálogo.", "success")
    return _redir_modulo_ut(request.form)

@app.route("/cargos-asignaciones/modulos/<int:modulo_id>/estado", methods=["POST"])
@permiso_requerido("cargos_asignaciones", "modificar")
@login_required
def cambiar_estado_modulo_ut(modulo_id):
    conn = _obtener_conexion_matriz_ut()
    cursor = conn.cursor()
    modulo = cursor.execute(
        "SELECT activo, obligatorio_todos FROM sistemas_modulos_ut WHERE id = ?",
        (modulo_id,),
    ).fetchone()
    if not modulo:
        conn.close()
        flash("El sistema o módulo solicitado no existe.", "warning")
        return _redir_modulo_ut(request.form)
    if modulo["activo"] and modulo["obligatorio_todos"]:
        conn.close()
        flash("No se pueden desactivar los accesos requeridos para todo el personal.", "warning")
        return _redir_modulo_ut(request.form)
    if modulo["activo"] and cursor.execute(
        "SELECT 1 FROM sistemas_modulos_ut WHERE padre_id = ? AND activo = 1 LIMIT 1",
        (modulo_id,),
    ).fetchone():
        conn.close()
        flash("Desactive primero los módulos y aplicativos dependientes.", "warning")
        return _redir_modulo_ut(request.form)
    nuevo_estado = 0 if modulo["activo"] else 1
    cursor.execute(
        "UPDATE sistemas_modulos_ut SET activo = ? WHERE id = ?",
        (nuevo_estado, modulo_id),
    )
    conn.commit()
    conn.close()
    flash(
        "Elemento del catálogo activado." if nuevo_estado else
        "Elemento del catálogo desactivado; se conserva el historial y las asignaciones.",
        "success",
    )
    return _redir_modulo_ut(request.form)

@app.route("/cargos-asignaciones/perfiles/nuevo", methods=["POST"])
@permiso_requerido("cargos_asignaciones", "modificar")
@login_required
def agregar_perfil_cargo_ut():
    cargo = request.form.get("cargo", "").strip()
    funciones = request.form.get("funciones", "").strip()
    autorizador = request.form.get("autorizador", "").strip()
    justificacion = request.form.get("justificacion", "").strip()
    if not cargo or len(cargo) > 160 or not funciones or not autorizador or not justificacion:
        flash("Complete el cargo, sus funciones, el responsable que autoriza y la justificación.", "warning")
        return redirect(url_for("cargos_asignaciones_ut"))

    conn = _obtener_conexion_matriz_ut()
    cursor = conn.cursor()
    duplicado = cursor.execute(
        "SELECT 1 FROM perfiles_cargo_ut WHERE cargo = ? COLLATE NOCASE",
        (cargo,),
    ).fetchone()
    if duplicado:
        conn.close()
        flash("Ya existe un perfil para ese cargo.", "warning")
        return redirect(url_for("cargos_asignaciones_ut"))
    cursor.execute("""
        INSERT INTO perfiles_cargo_ut (cargo, funciones, autorizador)
        VALUES (?, ?, ?)
    """, (cargo, funciones, autorizador))
    perfil_id = cursor.lastrowid
    cursor.execute("""
        INSERT INTO historial_perfiles_cargo_ut
            (perfil_id, cargo, cambios, justificacion, autorizado_por, usuario_id)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        perfil_id,
        cargo,
        json.dumps({"perfil": "creado", "funciones": funciones}, ensure_ascii=False),
        justificacion,
        autorizador,
        session.get("usuario_id"),
    ))
    conn.commit()
    conn.close()
    flash("Perfil de cargo creado. Configure y valide sus accesos requeridos.", "success")
    return redirect(url_for("cargos_asignaciones_ut"))

@app.route("/cargos-asignaciones/perfiles/<int:perfil_id>/guardar", methods=["POST"])
@permiso_requerido("cargos_asignaciones", "modificar")
@login_required
def guardar_perfil_cargo_ut(perfil_id):
    funciones = request.form.get("funciones", "").strip()
    autorizador = request.form.get("autorizador", "").strip()
    reportes_requeridos = request.form.get("reportes_requeridos", "").strip()
    accion_acceso = request.form.get("accion_acceso", "").strip().upper()
    justificacion = request.form.get("justificacion", "").strip()
    modulo_ids = request.form.getlist("modulo_id")
    if (
        not funciones or len(funciones) > 2000
        or not autorizador or len(autorizador) > 200
        or not reportes_requeridos or len(reportes_requeridos) > 2000
        or accion_acceso not in ACCIONES_SOLICITUD_UT
        or not justificacion or len(justificacion) > 2000
    ):
        flash(
            "Complete las funciones, reportes requeridos, acción del acceso, "
            "responsable que autoriza y justificación.",
            "warning",
        )
        return redirect(url_for("cargos_asignaciones_ut"))
    if any(not re.fullmatch(r"[0-9]+", valor) for valor in modulo_ids):
        flash("La selección de sistemas o módulos contiene un identificador inválido.", "warning")
        return redirect(url_for("cargos_asignaciones_ut"))

    conn = _obtener_conexion_matriz_ut()
    cursor = conn.cursor()
    perfil = cursor.execute(
        "SELECT * FROM perfiles_cargo_ut WHERE id = ? AND activo = 1",
        (perfil_id,),
    ).fetchone()
    if not perfil:
        conn.close()
        flash("El perfil seleccionado no existe o está inactivo.", "warning")
        return redirect(url_for("cargos_asignaciones_ut"))

    modulos_seleccionados = {int(valor) for valor in modulo_ids}
    modulos_activos = {
        row["id"]: row for row in cursor.execute("""
            SELECT id, nombre, obligatorio_todos
            FROM sistemas_modulos_ut
            WHERE activo = 1
        """).fetchall()
    }
    if not modulos_seleccionados.issubset(modulos_activos):
        conn.close()
        flash("La selección incluye un sistema o módulo inexistente o inactivo.", "warning")
        return redirect(url_for("cargos_asignaciones_ut"))

    modulos_seleccionados.update(
        modulo_id
        for modulo_id, modulo in modulos_activos.items()
        if modulo["obligatorio_todos"]
    )
    niveles_nuevos = {}
    for modulo_id in modulos_seleccionados:
        predeterminado = (
            "CUENTA" if modulos_activos[modulo_id]["obligatorio_todos"] else "LECTURA"
        )
        nivel = request.form.get(f"nivel_{modulo_id}", predeterminado).strip().upper()
        if nivel not in NIVELES_ACCESO_UT:
            conn.close()
            flash("Uno de los niveles de acceso seleccionados no es válido.", "warning")
            return redirect(url_for("cargos_asignaciones_ut"))
        niveles_nuevos[modulo_id] = nivel

    niveles_anteriores = {
        row["sistema_modulo_id"]: row["nivel_acceso"]
        for row in cursor.execute("""
            SELECT sistema_modulo_id, nivel_acceso
            FROM accesos_perfil_cargo_ut
            WHERE perfil_id = ?
        """, (perfil_id,)).fetchall()
    }
    cambios = []
    for modulo_id in sorted(set(niveles_anteriores) | set(niveles_nuevos)):
        nivel_anterior = niveles_anteriores.get(modulo_id)
        nivel_nuevo = niveles_nuevos.get(modulo_id)
        if nivel_anterior != nivel_nuevo:
            modulo = modulos_activos.get(modulo_id)
            cambios.append({
                "modulo": modulo["nombre"] if modulo else f"ID {modulo_id}",
                "antes": nivel_anterior,
                "despues": nivel_nuevo,
            })
    if perfil["funciones"] != funciones:
        cambios.append({"campo": "funciones", "antes": perfil["funciones"], "despues": funciones})
    if perfil["autorizador"] != autorizador:
        cambios.append({"campo": "autorizador", "antes": perfil["autorizador"], "despues": autorizador})
    if perfil["reportes_requeridos"] != reportes_requeridos:
        cambios.append({
            "campo": "reportes_requeridos",
            "antes": perfil["reportes_requeridos"],
            "despues": reportes_requeridos,
        })
    if perfil["accion_acceso"] != accion_acceso:
        cambios.append({
            "campo": "accion_acceso",
            "antes": perfil["accion_acceso"],
            "despues": accion_acceso,
        })
    if not perfil["validado"]:
        cambios.append({"campo": "perfil", "antes": "PENDIENTE", "despues": "VALIDADO"})
    if not cambios:
        conn.close()
        flash("No hubo cambios en la matriz del cargo.", "info")
        return redirect(url_for("cargos_asignaciones_ut"))

    cursor.execute("""
        UPDATE perfiles_cargo_ut
        SET funciones = ?, autorizador = ?, reportes_requeridos = ?,
            accion_acceso = ?, validado = 1, actualizado_el = CURRENT_TIMESTAMP
        WHERE id = ?
    """, (funciones, autorizador, reportes_requeridos, accion_acceso, perfil_id))
    cursor.execute("DELETE FROM accesos_perfil_cargo_ut WHERE perfil_id = ?", (perfil_id,))
    cursor.executemany("""
        INSERT INTO accesos_perfil_cargo_ut (perfil_id, sistema_modulo_id, nivel_acceso)
        VALUES (?, ?, ?)
    """, [
        (perfil_id, modulo_id, nivel)
        for modulo_id, nivel in sorted(niveles_nuevos.items())
    ])
    cursor.execute("""
        INSERT INTO historial_perfiles_cargo_ut
            (perfil_id, cargo, cambios, justificacion, autorizado_por, usuario_id)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        perfil_id,
        perfil["cargo"],
        json.dumps(cambios, ensure_ascii=False),
        justificacion,
        autorizador,
        session.get("usuario_id"),
    ))
    conn.commit()
    conn.close()
    flash(f"Matriz de accesos del cargo «{perfil['cargo']}» guardada y validada.", "success")
    return redirect(url_for("cargos_asignaciones_ut"))

@app.route("/cargos-asignaciones/exportar-excel")
@permiso_requerido("cargos_asignaciones", "reportar")
@login_required
def exportar_cargos_asignaciones_ut():
    conn = _obtener_conexion_matriz_ut()
    reporte = _obtener_reporte_accesos_ut(conn.cursor())
    conn.close()

    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Accesos requeridos UT"
    encabezados = [
        "NOMBRE DEL TRABAJADOR", "CARGO", "SISTEMA / MÓDULO / APK",
        "NIVEL DE ACCESO", "ORIGEN DEL REQUERIMIENTO", "ESTADO DEL PERFIL",
        "FUNCIONES DEL CARGO", "RESPONSABLE DEL CARGO", "ACCIÓN SOLICITADA",
        "ESTADO DE SOLICITUD", "REPORTES REQUERIDOS", "FUNCIONES PRINCIPALES",
        "RESPONSABLE QUE VALIDA",
    ]
    hoja.append(encabezados)
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="17365D")
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for trabajador in reporte:
        if trabajador["accesos"]:
            for acceso in trabajador["accesos"]:
                hoja.append([
                    trabajador["nombre"], trabajador["cargo"], acceso["ruta"],
                    acceso["nivel_nombre"], "General o perfil del cargo",
                    trabajador["estado_perfil"], trabajador["funciones"],
                    trabajador["autorizador"], "", "", "", "", "",
                ])
        solicitud = trabajador["solicitud"]
        if solicitud:
            for modulo in solicitud["modulos"]:
                hoja.append([
                    trabajador["nombre"], trabajador["cargo"], modulo["ruta"],
                    modulo["nivel_nombre"], "Matriz validada por cargo",
                    trabajador["estado_perfil"], trabajador["funciones"],
                    trabajador["autorizador"], solicitud["accion_nombre"],
                    solicitud["estado"], solicitud["reportes_requeridos"],
                    solicitud["funciones"], solicitud["autorizador"],
                ])
        elif not trabajador["accesos"]:
            hoja.append([
                trabajador["nombre"], trabajador["cargo"], "Sin sistemas asociados",
                "", "", trabajador["estado_perfil"], trabajador["funciones"],
                trabajador["autorizador"], "", "", "", "", "",
            ])
    for fila in hoja.iter_rows(min_row=2):
        for celda in fila:
            celda.data_type = "s"
            celda.alignment = Alignment(vertical="top", wrap_text=True)
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for columna, ancho in {
        "A": 34, "B": 28, "C": 48, "D": 22, "E": 26, "F": 26,
        "G": 55, "H": 32, "I": 20, "J": 24, "K": 48, "L": 55, "M": 32,
    }.items():
        hoja.column_dimensions[columna].width = ancho

    buffer = BytesIO()
    libro.save(buffer)
    buffer.seek(0)
    nombre = f"Accesos_Requeridos_UT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

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

    if len(clave_nueva) < 12:
        flash("La nueva contraseña debe tener al menos 12 caracteres.", "warning")
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
    try:
        trabajador_id = validar_id_formulario(
            request.form.get("trabajador_id"), "trabajador"
        )
    except ValueError:
        trabajador_id = None
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
    try:
        user_id = validar_id_formulario(request.form.get("usuario_id"), "usuario")
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_usuarios"))
    nombre_input = request.form["nombre_completo"].strip()
    pass_input = request.form.get("password", "").strip()
    rol_input = request.form.get("rol", "USUARIO").strip().upper()
    permisos = permisos_desde_formulario(request.form)

    if pass_input and len(pass_input) < 12:
        flash("La nueva contraseña debe tener al menos 12 caracteres.", "warning")
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
    hoy_iso = date.today().isoformat()
    tokens_vigentes = cursor.execute("""
        SELECT COUNT(*) FROM certificados_digitales
        WHERE estado = 'VIGENTE' AND fecha_inicio <= ? AND fecha_caducidad >= ?
    """, (hoy_iso, hoy_iso)).fetchone()[0]
    fecha_limite_tokens = (date.today() + timedelta(days=DIAS_ALERTA_CERTIFICADO)).isoformat()
    tokens_por_vencer = cursor.execute("""
        SELECT COUNT(*) FROM certificados_digitales
        WHERE estado = 'VIGENTE' AND fecha_caducidad <= ?
    """, (fecha_limite_tokens,)).fetchone()[0]
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
                           tokens_vigentes=tokens_vigentes,
                           tokens_por_vencer=tokens_por_vencer,
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
    try:
        fecha_publicacion = validar_fecha_formulario(
            request.form.get("fecha_publicacion"), "la fecha de publicación"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_biblioteca"))
    conn = get_biblioteca_db()
    conn.execute("""
        INSERT INTO biblioteca_pae (fecha_publicacion, nombre_evento, link, observaciones)
        VALUES (?, ?, ?, ?)
    """, (fecha_publicacion, evento, link,
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
            return openpyxl.utils.datetime.from_excel(valor).date().isoformat()
        except (TypeError, ValueError, OverflowError) as error:
            raise ValueError("La fecha de publicación del Excel no es válida.") from error
    texto = str(valor).strip()
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(texto, formato).date().isoformat()
        except ValueError:
            continue
    raise ValueError("La fecha de publicación del Excel no es válida.")

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
        try:
            fecha = convertir_fecha_biblioteca(
                valores[indice_fecha] if indice_fecha is not None and indice_fecha < len(valores) else ""
            )
        except ValueError:
            omitidos += 1
            continue
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

    configuracion_umbral = cursor.execute(
        "SELECT valor FROM configuraciones_sistema WHERE clave = ?",
        ("sgd_umbral_antiguedad_dias",)
    ).fetchone()
    umbral_dias = int(configuracion_umbral["valor"]) if configuracion_umbral else 100

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
                resultado = aplicar_umbral_antiguedad(
                    resultado, corte_db["fecha_corte"], umbral_dias
                )
                resultado["destinatarios_ocultos"] = ocultos
                fecha_corte_defecto = corte_db["fecha_corte"]
            except (json.JSONDecodeError, KeyError, TypeError, ValueError):
                app.logger.exception("No se pudo analizar el corte SGD %s.", corte_db["id"])
                flash("No se pudo analizar este corte del Monitor SGD. Revise el archivo del corte o contacte al administrador.", "danger")
                resultado = None

    conn.close()

    return render_template("sgd.html", 
                           tipo_activo=tipo,
                           cortes_historial=cortes_historial,
                           corte_activo=corte_activo,
                           resultado=resultado,
                           umbral_dias=umbral_dias,
                           puede_modificar_sgd=usuario_tiene_permiso("sgd", "modificar"),
                           fecha_corte=fecha_corte_defecto,
                           fecha_corte_input=fecha_corte_input)

@app.route("/sgd/configurar-antiguedad", methods=["POST"])
@permiso_requerido("sgd", "modificar")
@login_required
def configurar_umbral_antiguedad_sgd():
    umbral_dias = request.form.get("umbral_dias", type=int)
    tipo = request.form.get("tipo", "NO_LEIDOS").upper()
    if tipo not in ["NO_LEIDOS", "RECIBIDOS"]:
        tipo = "NO_LEIDOS"
    corte_id = request.form.get("corte_id", type=int)

    if umbral_dias is None or not 1 <= umbral_dias <= 3650:
        flash("El rango debe ser un número entero entre 1 y 3650 días.", "warning")
        return redirect(f"/sgd?tipo={tipo}" + (f"&corte_id={corte_id}" if corte_id else ""))

    conn = get_db()
    conn.execute("""
        INSERT INTO configuraciones_sistema (clave, valor)
        VALUES (?, ?)
        ON CONFLICT(clave) DO UPDATE
        SET valor = excluded.valor, actualizado_el = CURRENT_TIMESTAMP
    """, ("sgd_umbral_antiguedad_dias", str(umbral_dias)))
    conn.commit()
    conn.close()

    flash(f"El rango de antigüedad del Monitor SGD se actualizó a {umbral_dias} días.", "success")
    return redirect(f"/sgd?tipo={tipo}" + (f"&corte_id={corte_id}" if corte_id else ""))

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
    num_carta_raw = request.form.get("numero_carta", "").strip()
    num_informe_raw = request.form.get("numero_informe", "").strip()
    num_os = request.form.get("numero_os", "0005612").strip()
    periodo_desde = request.form.get("periodo_desde", "")
    periodo_hasta = request.form.get("periodo_hasta", "")
    guardar_en_bd = request.form.get("guardar_en_bd") == "1"

    try:
        fecha_emision = validar_fecha_formulario(
            request.form.get("fecha_emision", date.today().isoformat()),
            "la fecha de emisión",
            obligatoria=True,
        )
        periodo_desde, periodo_hasta = validar_rango_fechas(
            periodo_desde, periodo_hasta, "periodo"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("gestion_documentaria"))
    if num_os and not re.fullmatch(r"[0-9]{1,20}", num_os):
        flash("El número de orden de servicio debe contener solo números.", "warning")
        return redirect(url_for("gestion_documentaria"))
    if any(
        valor and not re.fullmatch(r"[0-9]{1,9}", valor)
        for valor in (num_carta_raw, num_informe_raw)
    ):
        flash("Los números de carta e informe deben ser enteros positivos.", "warning")
        return redirect(url_for("gestion_documentaria"))
    num_carta = int(num_carta_raw) if num_carta_raw else 9
    num_informe = int(num_informe_raw) if num_informe_raw else 9

    if num_carta <= 0 or num_informe <= 0:
        flash("Los números de carta e informe deben ser enteros positivos.", "warning")
        return redirect(url_for("gestion_documentaria"))
    anio = date.fromisoformat(fecha_emision).year

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

@app.route("/api/apowerrec/status", methods=["GET"])
def estado_sincronizacion_apowerrec():
    token_configurado = os.getenv("APOWERREC_SYNC_TOKEN", "")
    token_recibido = request.headers.get("Authorization", "")
    if not token_configurado:
        return jsonify(error="La sincronización ApowerREC no está configurada en el servidor."), 503
    if not hmac.compare_digest(token_recibido, f"Bearer {token_configurado}"):
        return jsonify(error="No autorizado."), 401

    conn = get_db()
    try:
        ensure_apowerrec_schema(conn)
        tareas = conn.execute("""
            SELECT id, titulo, estado_iniciado_el
            FROM actividades_soporte
            WHERE categoria = 'GRABAR SESION' AND estado = 'EN PROGRESO'
            ORDER BY id
        """).fetchall()
        if len(tareas) != 1:
            return jsonify(ready=False, tareas_en_progreso=len(tareas))
        if tareas[0]["estado_iniciado_el"] is None:
            return jsonify(
                ready=False,
                tareas_en_progreso=1,
                motivo="Vuelva a marcar la actividad GRABAR SESION como EN PROGRESO."
            )
        return jsonify(
            ready=True,
            actividad_id=tareas[0]["id"],
            titulo=tareas[0]["titulo"],
            iniciada_el=tareas[0]["estado_iniciado_el"]
        )
    finally:
        conn.close()

@app.route("/api/apowerrec/recordings", methods=["POST"])
def sincronizar_grabacion_apowerrec():
    token_configurado = os.getenv("APOWERREC_SYNC_TOKEN", "")
    token_recibido = request.headers.get("Authorization", "")
    if not token_configurado:
        return jsonify(error="La sincronización ApowerREC no está configurada en el servidor."), 503
    if not hmac.compare_digest(token_recibido, f"Bearer {token_configurado}"):
        return jsonify(error="No autorizado."), 401

    datos = request.get_json(silent=True)
    if not isinstance(datos, dict):
        return jsonify(error="El cuerpo JSON de la grabación no es válido."), 400

    nombre_archivo = secure_filename(str(datos.get("filename", "")))
    if not nombre_archivo or os.path.splitext(nombre_archivo)[1].lower() != ".mp4":
        return jsonify(error="Solo se aceptan grabaciones MP4 con nombre válido."), 400

    hash_declarado = str(datos.get("sha256", "")).lower()
    if not re.fullmatch(r"[a-f0-9]{64}", hash_declarado):
        return jsonify(error="Falta una huella SHA-256 válida del video."), 400
    valor_actividad_id = datos.get("actividad_id")
    if isinstance(valor_actividad_id, bool) or not isinstance(valor_actividad_id, (int, str)):
        return jsonify(error="El identificador de actividad no es válido."), 400
    try:
        actividad_id_esperada = int(valor_actividad_id)
    except (TypeError, ValueError):
        return jsonify(error="El identificador de actividad no es válido."), 400
    if actividad_id_esperada <= 0:
        return jsonify(error="El identificador de actividad no es válido."), 400

    drive_file_id = str(datos.get("drive_file_id", "")).strip()
    drive_url = str(datos.get("drive_url", "")).strip()
    parsed_drive_url = urllib.parse.urlsplit(drive_url)
    if (
        not re.fullmatch(r"[A-Za-z0-9_-]{10,200}", drive_file_id)
        or parsed_drive_url.scheme != "https"
        or parsed_drive_url.hostname not in {"drive.google.com", "docs.google.com"}
        or (
            drive_file_id not in parsed_drive_url.path
            and drive_file_id not in urllib.parse.parse_qs(parsed_drive_url.query).get("id", [])
        )
    ):
        return jsonify(error="El identificador o enlace de Google Drive no es válido."), 400

    try:
        fecha_grabacion = date.fromisoformat(str(datos.get("fecha_grabacion", ""))).isoformat()
        mtime_archivo = float(datos.get("mtime", ""))
    except (TypeError, ValueError):
        return jsonify(error="La fecha o marca de tiempo de la grabación no es válida."), 400
    if not (mtime_archivo > 0 and math.isfinite(mtime_archivo)):
        return jsonify(error="La marca de tiempo de la grabación no es válida."), 400

    conn = get_db()
    biblioteca_conn = None
    try:
        ensure_apowerrec_schema(conn)
        sincronizacion = conn.execute("""
            SELECT g.*, a.titulo AS nombre_evento, a.estado AS estado_actividad
            FROM grabaciones_apowerrec g
            LEFT JOIN actividades_soporte a ON a.id = g.actividad_id
            WHERE g.sha256 = ?
        """, (hash_declarado,)).fetchone()
        if sincronizacion and sincronizacion["estado"] == "SINCRONIZADA":
            return jsonify(
                estado="SINCRONIZADA",
                repetida=True,
                drive_url=sincronizacion["drive_url"],
                nombre_evento=sincronizacion["nombre_evento"] or sincronizacion["archivo_nombre"],
                actividad_completada=sincronizacion["estado_actividad"] == "COMPLETADA"
            ), 200
        ahora = time.time()
        if sincronizacion and sincronizacion["estado"] == "PROCESANDO" and ahora - sincronizacion["actualizado_el"] < 1800:
            return jsonify(error="Este video ya está siendo procesado; se reintentará automáticamente."), 409

        tareas_en_progreso = conn.execute("""
            SELECT id, titulo, solucion_aplicada, estado_iniciado_el
            FROM actividades_soporte
            WHERE categoria = 'GRABAR SESION' AND estado = 'EN PROGRESO'
            ORDER BY id
        """).fetchall()
        if len(tareas_en_progreso) != 1:
            return jsonify(error="Debe existir exactamente una actividad GRABAR SESION en progreso para vincular el video."), 409

        tarea = tareas_en_progreso[0]
        if tarea["id"] != actividad_id_esperada:
            return jsonify(error="La actividad en progreso cambió durante la transferencia; el video no se asociará a otra tarea."), 409
        if tarea["estado_iniciado_el"] is None or mtime_archivo < float(tarea["estado_iniciado_el"]) - 2:
            return jsonify(error="El video debe haberse grabado después de iniciar la actividad GRABAR SESION."), 409

        if sincronizacion:
            conn.execute("""
                UPDATE grabaciones_apowerrec
                SET archivo_nombre = ?, actividad_id = ?, fecha_grabacion = ?,
                    estado = 'PROCESANDO', error = NULL, actualizado_el = ?
                WHERE sha256 = ?
            """, (nombre_archivo, tarea["id"], fecha_grabacion, ahora, hash_declarado))
        else:
            conn.execute("""
                INSERT INTO grabaciones_apowerrec
                    (sha256, archivo_nombre, actividad_id, fecha_grabacion, estado, actualizado_el)
                VALUES (?, ?, ?, ?, 'PROCESANDO', ?)
            """, (hash_declarado, nombre_archivo, tarea["id"], fecha_grabacion, ahora))
        conn.commit()

        observaciones = f"Grabación ApowerREC sincronizada automáticamente desde la actividad TI #{tarea['id']}."
        biblioteca_conn = get_biblioteca_db()
        registro = biblioteca_conn.execute(
            "SELECT id FROM biblioteca_pae WHERE link = ?", (drive_url,)
        ).fetchone()
        if registro:
            biblioteca_id = registro["id"]
        else:
            cursor_biblioteca = biblioteca_conn.execute("""
                INSERT INTO biblioteca_pae (fecha_publicacion, nombre_evento, link, observaciones)
                VALUES (?, ?, ?, ?)
            """, (fecha_grabacion, tarea["titulo"], drive_url, observaciones))
            biblioteca_id = cursor_biblioteca.lastrowid
        biblioteca_conn.commit()

        solucion_actualizada = (
            f"{tarea['solucion_aplicada'].strip()}\n{observaciones} {drive_url}"
            if (tarea["solucion_aplicada"] or "").strip()
            else f"{observaciones} {drive_url}"
        )
        cursor_actividad = conn.execute("""
            UPDATE actividades_soporte
            SET estado = 'COMPLETADA', estado_iniciado_el = NULL, solucion_aplicada = ?
            WHERE id = ? AND estado = 'EN PROGRESO' AND estado_iniciado_el = ?
        """, (solucion_actualizada, tarea["id"], tarea["estado_iniciado_el"]))
        actividad_completada = cursor_actividad.rowcount == 1
        conn.execute("""
            UPDATE grabaciones_apowerrec
            SET estado = 'SINCRONIZADA', drive_file_id = ?, drive_url = ?,
                biblioteca_id = ?, error = NULL, actualizado_el = ?
            WHERE sha256 = ?
        """, (drive_file_id, drive_url, biblioteca_id, time.time(), hash_declarado))
        conn.commit()
        return jsonify(
            estado="SINCRONIZADA",
            repetida=False,
            drive_url=drive_url,
            nombre_evento=tarea["titulo"],
            actividad_completada=actividad_completada,
            advertencia=None if actividad_completada else "El video quedó en Drive y en Biblioteca PAE, pero la actividad cambió de estado durante la carga."
        ), 200
    except Exception:
        app.logger.exception("Falló la sincronización de una grabación ApowerREC.")
        try:
            conn.execute("""
                UPDATE grabaciones_apowerrec
                SET estado = 'ERROR', error = 'Revise el registro del servidor.', actualizado_el = ?
                WHERE sha256 = ?
            """, (time.time(), hash_declarado))
            conn.commit()
        except Exception:
            app.logger.exception("No se pudo guardar el estado de error de la grabación ApowerREC.")
        return jsonify(error="No se pudo completar la sincronización. Revise el registro del servidor; el archivo se podrá reintentar."), 502
    finally:
        if biblioteca_conn:
            biblioteca_conn.close()
        conn.close()


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
    try:
        fecha_emision = validar_fecha_formulario(
            request.form.get("fecha_emision") or date.today().isoformat(),
            "la fecha de emisión",
            obligatoria=True,
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("gestion_documentaria"))
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
                file_metadata = {"name": archivo_nombre, "description": f"Documento cargado desde SISGEMOSI: {titulo}"}
                with open(archivo_ruta, "rb") as contenido:
                    media = MediaIoBaseUpload(contenido, mimetype=archivo.mimetype or "application/octet-stream", resumable=True)
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
    actividad_nueva_id = request.args.get("nueva", type=int)
    mostrar_tareas_vencidas = session.pop("abrir_tareas_vencidas", False)
    conn = get_db()
    ensure_apowerrec_schema(conn)
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
        item["codigo_control"] = codigo_control_actividad(item["id"])
        actividades.append(item)

    trabajadores = cursor.execute("""
        SELECT id, dni, (nombres || ' ' || apellido_paterno || ' ' || apellido_materno) AS nombre_completo, cargo 
        FROM trabajadores 
        WHERE activo = 1
        ORDER BY nombres ASC
    """).fetchall()
    actividades_vencidas = cursor.execute(f"""
        SELECT * FROM actividades_soporte
        WHERE {propietario_sql}
          AND fecha_programada < ?
          AND estado = 'PENDIENTE'
        ORDER BY fecha_programada ASC, hora_inicio ASC, id ASC
    """, (*propietario_params, hoy_dt.isoformat())).fetchall()

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
                           actividades_vencidas=actividades_vencidas,
                           trabajadores=trabajadores,
                           filtro_estado=filtro_estado,
                           filtro_cat=filtro_cat,
                           filtro_periodo=filtro_periodo,
                           actividad_nueva_id=actividad_nueva_id,
                           mostrar_tareas_vencidas=mostrar_tareas_vencidas,
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
    titulo = request.form.get("titulo", "").strip()
    descripcion = request.form.get("descripcion", "").strip()
    categoria = request.form.get("categoria", "SOPORTE TECNICO").strip()
    area = request.form.get("area_solicitante", "").strip()
    prioridad = request.form.get("prioridad", "MEDIA").strip()
    estado = request.form.get("estado", "PENDIENTE").strip()
    try:
        validar_campos_actividad(titulo, descripcion, categoria, area, prioridad, estado)
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))
    try:
        fecha_prog = validar_fecha_formulario(
            request.form.get("fecha_programada") or date.today().isoformat(),
            "la fecha programada",
            obligatoria=True,
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))
    h_inicio = request.form.get("hora_inicio", "").strip()
    h_fin = request.form.get("hora_fin", "").strip()
    try:
        whatsapp_telefono = validar_numero_contacto(
            request.form.get("whatsapp_telefono"), "El número de WhatsApp"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))

    if not h_inicio:
        h_inicio = datetime.now().strftime("%H:%M")
    if not h_fin:
        h_fin = (datetime.now() + timedelta(hours=1)).strftime("%H:%M")
    try:
        h_inicio = validar_hora_formulario(h_inicio, "La hora de inicio")
        h_fin = validar_hora_formulario(h_fin, "La hora de fin")
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))

    conn = get_db()
    cursor = conn.cursor()
    estado_iniciado_el = time.time() if estado == "EN PROGRESO" else None
    cursor.execute("""
        INSERT INTO actividades_soporte (
            titulo, descripcion, categoria, area_solicitante, prioridad, estado,
            fecha_programada, hora_inicio, hora_fin, whatsapp_telefono, usuario_id, estado_iniciado_el
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (titulo, descripcion, categoria, area, prioridad, estado, fecha_prog, h_inicio, h_fin,
          whatsapp_telefono,
          session.get("usuario_id"), estado_iniciado_el))
    actividad_id = cursor.lastrowid
    conn.commit()
    conn.close()

    flash("Actividad registrada exitosamente.", "success")
    return redirect(url_for("ver_actividades", estado="TODOS", nueva=actividad_id))

@app.route("/actividades/editar", methods=["POST"])
@permiso_requerido("actividades", "modificar")
@login_required
def editar_actividad():
    try:
        actividad_id = validar_id_formulario(
            request.form.get("actividad_id"), "actividad"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))
    titulo = request.form.get("titulo", "").strip()
    descripcion = request.form.get("descripcion", "").strip()
    categoria = request.form.get("categoria", "SOPORTE TECNICO").strip()
    area = request.form.get("area_solicitante", "").strip()
    prioridad = request.form.get("prioridad", "MEDIA").strip()
    estado = request.form.get("estado", "PENDIENTE").strip()
    try:
        validar_campos_actividad(titulo, descripcion, categoria, area, prioridad, estado)
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))
    try:
        fecha_prog = validar_fecha_formulario(
            request.form.get("fecha_programada") or date.today().isoformat(),
            "la fecha programada",
            obligatoria=True,
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))
    h_inicio = request.form.get("hora_inicio", "").strip()
    h_fin = request.form.get("hora_fin", "").strip()
    solucion = request.form.get("solucion_aplicada", "").strip()
    try:
        whatsapp_telefono = validar_numero_contacto(
            request.form.get("whatsapp_telefono"), "El número de WhatsApp"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))
    if len(solucion) > 5000:
        flash("La solución aplicada no puede superar los 5000 caracteres.", "warning")
        return redirect(url_for("ver_actividades"))
    try:
        h_inicio = validar_hora_formulario(h_inicio, "La hora de inicio")
        h_fin = validar_hora_formulario(h_fin, "La hora de fin")
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))

    conn = get_db()
    propietario_sql, propietario_params = filtro_propietario()
    conn.cursor().execute(f"""
        UPDATE actividades_soporte 
        SET titulo = ?, descripcion = ?, categoria = ?, area_solicitante = ?, prioridad = ?,
            estado = ?, fecha_programada = ?, hora_inicio = ?, hora_fin = ?, whatsapp_telefono = ?, solucion_aplicada = ?,
            estado_iniciado_el = CASE
                WHEN estado = 'EN PROGRESO' AND estado_iniciado_el IS NOT NULL THEN estado_iniciado_el
                WHEN ? = 'EN PROGRESO' THEN ?
                ELSE NULL
            END
        WHERE id = ? AND {propietario_sql}
    """, (titulo, descripcion, categoria, area, prioridad, estado, fecha_prog, h_inicio, h_fin,
          whatsapp_telefono,
          solucion, estado, time.time(), actividad_id, *propietario_params))
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
        cursor = conn.cursor()
        cursor.execute(f"""
            UPDATE actividades_soporte
            SET estado = ?,
                estado_iniciado_el = CASE
                    WHEN ? = 'EN PROGRESO' THEN
                        CASE
                            WHEN estado = 'EN PROGRESO' AND estado_iniciado_el IS NOT NULL THEN estado_iniciado_el
                            ELSE ?
                        END
                    ELSE NULL
                END
            WHERE id = ? AND {propietario_sql}
        """, (nuevo_estado, nuevo_estado, time.time(), actividad_id, *propietario_params))
        actualizado = cursor.rowcount > 0
        conn.commit()
        conn.close()
        if actualizado:
            session["abrir_tareas_vencidas"] = True
            if nuevo_estado == "PENDIENTE":
                flash("La tarea se mantiene como PENDIENTE porque no se completó; seguirá en TAREAS NO REALIZADAS.", "warning")
            else:
                flash(f"La tarea se marcó correctamente como {nuevo_estado}.", "success")
        else:
            flash("No se encontró la tarea o no tienes permiso para modificarla.", "danger")
    else:
        flash("El estado solicitado no es válido.", "danger")
    return redirect(request.referrer or url_for("ver_actividades"))

@app.route("/actividades/acciones-realizadas", methods=["POST"])
@permiso_requerido("actividades", "modificar")
@login_required
def guardar_acciones_realizadas():
    try:
        actividad_id = validar_id_formulario(
            request.form.get("actividad_id"), "actividad"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_actividades"))
    acciones = request.form.get("solucion_aplicada", "").strip()

    if not acciones:
        flash("Debe registrar las acciones realizadas antes de completar la actividad.", "warning")
        return redirect(url_for("ver_actividades"))

    conn = get_db()
    propietario_sql, propietario_params = filtro_propietario()
    conn.cursor().execute(f"""
        UPDATE actividades_soporte
        SET estado = 'COMPLETADA', solucion_aplicada = ?, estado_iniciado_el = NULL
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
        codigo_control = codigo_control_actividad(f["id"])
        
        eventos.append({
            "id": f["id"],
            "title": f"[{codigo_control}] [{f['categoria']}] {f['titulo']}",
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

@app.route("/api/actividades-alertas")
@permiso_requerido("actividades", "ver")
@login_required
def api_actividades_alertas():
    ahora = datetime.now()
    limite = ahora + timedelta(minutes=5)
    propietario_sql, propietario_params = filtro_propietario()
    conn = get_db()
    filas = conn.cursor().execute(f"""
        SELECT id, titulo, area_solicitante, fecha_programada, hora_inicio, whatsapp_telefono
        FROM actividades_soporte
        WHERE {propietario_sql}
          AND estado IN ('PENDIENTE', 'EN PROGRESO')
          AND fecha_programada IS NOT NULL AND fecha_programada != ''
          AND hora_inicio IS NOT NULL AND hora_inicio != ''
          AND datetime(fecha_programada || ' ' || hora_inicio) BETWEEN datetime(?) AND datetime(?)
        ORDER BY fecha_programada, hora_inicio, id
    """, (*propietario_params, ahora.isoformat(sep=" ", timespec="seconds"),
          limite.isoformat(sep=" ", timespec="seconds"))).fetchall()
    conn.close()
    return jsonify([
        {
            "id": fila["id"],
            "titulo": fila["titulo"],
            "area_solicitante": fila["area_solicitante"] or "",
            "fecha_programada": fila["fecha_programada"],
            "hora_inicio": fila["hora_inicio"],
            "whatsapp_telefono": fila["whatsapp_telefono"] or "",
        }
        for fila in filas
    ])

@app.route("/actividades/exportar-excel")
@permiso_requerido("actividades", "reportar")
@login_required
def exportar_actividades_excel():
    conn = get_db()
    cursor = conn.cursor()
    propietario_sql, propietario_params = filtro_propietario()
    filas = cursor.execute(f"""
        SELECT id, fecha_programada, hora_inicio, hora_fin, area_solicitante, categoria,
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
        "ID DE CONTROL", "FECHA", "HORA INICIO", "HORA FIN", "ÁREA SOLICITANTE", "CATEGORÍA",
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

    cols_centradas = (1, 2, 3, 4, 5, 9, 10)
    for r_idx, f in enumerate(filas, start=2):
        row_data = [
            codigo_control_actividad(f["id"]),
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
    for actividad in actividades_lista:
        actividad["codigo_control"] = codigo_control_actividad(actividad["id"])

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

# ================= API PARA FRONTEND MODERNO =================
@app.route("/api/prestamos/resumen")
@login_required
def api_prestamos_resumen():
    conn = get_db()
    cursor = conn.cursor()

    categorias = ["MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"]
    resumen = {}

    for cat in categorias:
        row = cursor.execute(
            "SELECT COUNT(*) AS total, SUM(CASE WHEN estado='EN PRESTAMO' THEN 1 ELSE 0 END) AS prestados FROM equipos WHERE categoria=?",
            (cat,)
        ).fetchone()
        resumen[cat] = {
            "total": row[0] or 0,
            "prestados": row[1] or 0,
            "almacen": (row[0] or 0) - (row[1] or 0)
        }

    conn.close()
    return jsonify(resumen)

@app.route("/api/prestamos/<categoria>")
@login_required
def api_prestamos_categoria(categoria):
    cat = categoria.upper()
    if cat not in ["MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"]:
        return jsonify({"error": "Categoría no válida"}), 400

    conn = get_db()
    cursor = conn.cursor()

    query = """
    SELECT
        e.id, e.codigo_margesi, e.descripcion, e.estado,
        p.id AS prestamo_id, t.nombres || ' ' || t.apellido_paterno AS custodio
    FROM equipos e
    LEFT JOIN prestamos p ON e.id = p.equipo_id AND p.estado = 'ACTIVO'
    LEFT JOIN trabajadores t ON p.trabajador_id = t.id
    WHERE e.categoria = ?
    ORDER BY CASE WHEN e.estado = 'EN PRESTAMO' THEN 0 ELSE 1 END, e.codigo_margesi ASC
    """
    equipos = cursor.execute(query, (cat,)).fetchall()
    conn.close()

    return jsonify([dict(row) for row in equipos])

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
        t.id AS trabajador_id,
        t.dni AS trabajador_dni,
        (t.nombres || ' ' || t.apellido_paterno || ' ' || t.apellido_materno) AS trabajador_nombre,
        t.cargo AS trabajador_cargo,
        t.telefono AS trabajador_telefono,
        t.celular_institucional AS trabajador_celular_institucional
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


@app.route("/api/equipos/verificar-codigo")
@login_required
def api_verificar_codigo_equipo():
    codigo = request.args.get("codigo", "").strip()
    if not codigo:
        return jsonify({"existe": False})

    conn = get_db()
    equipo = buscar_equipo_por_codigo(conn.cursor(), codigo)
    conn.close()
    if not equipo:
        return jsonify({"existe": False})

    return jsonify({
        "existe": True,
        "codigo": equipo["codigo_margesi"] or equipo["codigo_interno"],
        "categoria": equipo["categoria"],
        "descripcion": equipo["descripcion"] or "",
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

    cat = request.form.get("categoria", "").strip().upper()
    cod = request.form.get("codigo_margesi", "").strip()
    desc = request.form.get("descripcion", "").strip()
    serie = request.form.get("serie_medidas", "").strip()
    try:
        celular_institucional = validar_numero_contacto(
            request.form.get("celular_institucional"), "El celular institucional"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))
    if cat not in {"MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"}:
        flash("La categoría seleccionada no es válida.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))
    if not cod or not desc:
        flash("El código y la descripción del bien son obligatorios.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE")
    equipo_duplicado = buscar_equipo_por_codigo(cursor, cod)
    if equipo_duplicado:
        codigo_existente = equipo_duplicado["codigo_margesi"] or equipo_duplicado["codigo_interno"]
        conn.close()
        flash(
            f"El código Margesí {cod} ya está registrado "
            f"(equipo {codigo_existente}). No se puede registrar dos veces.",
            "danger",
        )
        return redirect(f"/prestamos/{cat.lower()}")

    equipo_con_imei = (
        buscar_movil_con_imei(cursor, serie, codigo=cod)
        if cat == "MOVILES" else None
    )
    if equipo_con_imei:
        codigo_existente = (
            equipo_con_imei["codigo_margesi"]
            or equipo_con_imei["codigo_interno"]
            or equipo_con_imei["id"]
        )
        conn.close()
        flash(
            f"El IMEI, serie o código {serie or cod} ya existe en la base de datos "
            f"(equipo {codigo_existente}). No se registró el móvil.",
            "danger",
        )
        return redirect(f"/prestamos/{cat.lower()}")

    if cat == "MOVILES" and celular_ya_registrado_en_otro_equipo(cursor, celular_institucional):
        conn.close()
        flash("Ese número celular ya está registrado en otro equipo móvil.", "danger")
        return redirect(f"/prestamos/{cat.lower()}")

    cursor.execute("""
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
        celular_institucional
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

    try:
        equipo_id = validar_id_formulario(request.form.get("equipo_id"), "bien")
    except ValueError:
        flash("El identificador del bien no es válido.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))
    cat = request.form.get("categoria", "INFORMATICOS").strip().upper()
    cod = request.form.get("codigo_margesi", "").strip()
    desc = request.form.get("descripcion", "").strip()
    serie = request.form.get("serie_medidas", "").strip()
    try:
        celular_institucional = validar_numero_contacto(
            request.form.get("celular_institucional"), "El celular institucional"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))
    if cat not in {"MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO"}:
        flash("La categoría seleccionada no es válida.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))
    if not cod or not desc:
        flash("El código y la descripción del bien son obligatorios.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE")
    equipo_actual = cursor.execute("""
        SELECT e.categoria, e.celular_institucional, p.trabajador_id
        FROM equipos e
        LEFT JOIN prestamos p ON p.equipo_id = e.id AND p.estado = 'ACTIVO'
        WHERE e.id = ?
    """, (equipo_id,)).fetchone()
    equipo_duplicado = buscar_equipo_por_codigo(cursor, cod, equipo_id)
    if equipo_duplicado:
        codigo_existente = equipo_duplicado["codigo_margesi"] or equipo_duplicado["codigo_interno"]
        conn.close()
        flash(
            f"El código Margesí {cod} ya está registrado "
            f"(equipo {codigo_existente}). No se guardaron los cambios.",
            "danger",
        )
        return redirect(f"/prestamos/{cat.lower()}")

    if equipo_actual and equipo_actual["categoria"] == "MOVILES":
        equipo_con_imei = buscar_movil_con_imei(cursor, serie, equipo_id, codigo=cod)
        if equipo_con_imei:
            codigo_existente = (
                equipo_con_imei["codigo_margesi"]
                or equipo_con_imei["codigo_interno"]
                or equipo_con_imei["id"]
            )
            conn.close()
            flash(
                f"El IMEI, serie o código {serie or cod} ya existe en la base de datos "
                f"(equipo {codigo_existente}). No se registraron los cambios.",
                "danger",
            )
            return redirect(f"/prestamos/{cat.lower()}")

    if cat == "MOVILES" and celular_ya_registrado_en_otro_equipo(
        cursor, celular_institucional, equipo_id
    ):
        conn.close()
        flash("Ese número celular ya está registrado en otro equipo móvil.", "danger")
        return redirect(f"/prestamos/{cat.lower()}")
    if equipo_actual and equipo_actual["categoria"] == "MOVILES" and equipo_actual["trabajador_id"]:
        trabajador_con_celular = buscar_trabajador_con_celular(
            cursor, celular_institucional, equipo_actual["trabajador_id"]
        )
        if trabajador_con_celular:
            nombre = " ".join((
                trabajador_con_celular["nombres"],
                trabajador_con_celular["apellido_paterno"],
                trabajador_con_celular["apellido_materno"],
            )).strip()
            conn.close()
            flash(
                f"El número celular ya está asignado al trabajador {nombre} "
                f"(DNI {trabajador_con_celular['dni']}).",
                "danger",
            )
            return redirect(f"/prestamos/{cat.lower()}")

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
          celular_institucional,
          equipo_id))
    if equipo_actual and equipo_actual["categoria"] == "MOVILES" and equipo_actual["trabajador_id"]:
        cursor.execute("""
            UPDATE trabajadores
            SET celular_institucional = ?, celular_institucional_normalizado = ?
            WHERE id = ?
        """, (
            celular_institucional,
            normalizar_numero_celular(celular_institucional),
            equipo_actual["trabajador_id"],
        ))
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
    if categoria_forzada and categoria_forzada.strip().upper() not in {
        "MOVILES", "TOKENS", "INFORMATICOS", "MOBILIARIO",
    }:
        flash("La categoría indicada para importar no es válida.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))
    if categoria_forzada:
        categoria_forzada = categoria_forzada.strip().upper()
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
    equipo_id_raw = request.form.get("equipo_id", "").strip()
    trabajador_id_raw = request.form.get("trabajador_id", "").strip()
    observaciones = request.form.get("observaciones_entrega", request.form.get("observaciones", "")).strip()

    if not re.fullmatch(r"[0-9]+", equipo_id_raw) or not re.fullmatch(r"[0-9]+", trabajador_id_raw):
        flash("Debe seleccionar el bien y el trabajador antes de confirmar el préstamo.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))
    equipo_id = int(equipo_id_raw)
    trabajador_id = int(trabajador_id_raw)
    if equipo_id <= 0 or trabajador_id <= 0:
        flash("El bien o el trabajador seleccionado no es válido.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE")

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

    celular_equipo = (eq["celular_institucional"] or "").strip()
    if eq["categoria"] == "MOVILES" and celular_equipo:
        try:
            celular_equipo = validar_numero_contacto(
                celular_equipo, "El celular institucional del bien"
            )
        except ValueError as error:
            conn.close()
            flash(
                f"{error} Corrija el número del bien antes de asignarlo.",
                "danger",
            )
            return redirect(request.referrer or url_for("prestamos_hub"))
        trabajador_con_celular = buscar_trabajador_con_celular(
            cursor, celular_equipo, trabajador_id
        )
        if trabajador_con_celular:
            nombre = " ".join((
                trabajador_con_celular["nombres"],
                trabajador_con_celular["apellido_paterno"],
                trabajador_con_celular["apellido_materno"],
            )).strip()
            conn.close()
            flash(
                f"El celular {celular_equipo} ya está registrado para {nombre} "
                f"(DNI {trabajador_con_celular['dni']}).",
                "danger",
            )
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
    if eq["categoria"] == "MOVILES" and celular_equipo:
        cursor.execute("""
            UPDATE trabajadores
            SET celular_institucional = ?, celular_institucional_normalizado = ?
            WHERE id = ?
        """, (celular_equipo, normalizar_numero_celular(celular_equipo), trabajador_id))
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
        if re.fullmatch(r"[0-9]+", equipo_id) and int(equipo_id) > 0:
            conn = get_db()
            prestamo = conn.execute(
                "SELECT id FROM prestamos WHERE equipo_id = ? AND estado = 'ACTIVO'",
                (int(equipo_id),)
            ).fetchone()
            conn.close()
            prestamo_id = str(prestamo["id"]) if prestamo else ""

    if not re.fullmatch(r"[0-9]+", prestamo_id) or int(prestamo_id) <= 0:
        flash("No se encontró un préstamo activo para devolver.", "warning")
        return redirect(request.referrer or url_for("prestamos_hub"))
    prestamo_id = int(prestamo_id)

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
    ensure_access_matrix_schema(conn)
    conn.commit()
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
    bienes_asignados = cursor.execute("""
         SELECT p.trabajador_id, e.codigo_margesi, e.categoria, e.descripcion,
             e.marca, e.modelo, e.numero_serie, e.celular_institucional
        FROM prestamos p
        JOIN equipos e ON e.id = p.equipo_id
        WHERE p.estado = 'ACTIVO'
        ORDER BY e.categoria, e.codigo_margesi
    """).fetchall()
    bienes_activos_por_trabajador = {}
    for bien in bienes_asignados:
        bienes_activos_por_trabajador.setdefault(bien["trabajador_id"], []).append(dict(bien))
    perfiles_cargo_nuevo = _obtener_perfiles_cargo_validados(cursor)
    conn.close()
    return render_template("trabajadores.html", trabajadores=trabajadores,
                           trabajadores_antiguos=trabajadores_antiguos,
                           contratos_por_trabajador=contratos_por_trabajador,
                           bienes_activos_por_trabajador=bienes_activos_por_trabajador,
                           perfiles_cargo_nuevo=perfiles_cargo_nuevo)

@app.route("/trabajadores/nuevo", methods=["POST"])
@permiso_requerido("personal", "modificar")
@login_required
def registrar_trabajador():
    try:
        dni = validar_dni(request.form.get("dni"))
        correo_personal = validar_correo(
            request.form.get("correo_personal"), "El correo personal"
        )
        correo_institucional = validar_correo(
            request.form.get("correo_institucional"), "El correo institucional"
        )
        telefono = validar_numero_contacto(request.form.get("telefono"), "El teléfono")
        celular_institucional = validar_numero_contacto(
            request.form.get("celular_institucional"), "El celular institucional"
        )
        fecha_nacimiento = validar_fecha_formulario(
            request.form.get("fecha_nacimiento"), "la fecha de nacimiento"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(request.referrer or url_for("ver_trabajadores"))

    nombres = request.form["nombres"].strip()
    apellido_paterno = request.form["apellido_paterno"].strip()
    apellido_materno = request.form["apellido_materno"].strip()
    perfil_cargo_raw = request.form.get("perfil_cargo_id", "").strip()
    tipo_contrato = request.form.get("tipo_contrato", "").strip()
    if not all((nombres, apellido_paterno, apellido_materno)):
        flash("Nombres y ambos apellidos son obligatorios.", "warning")
        return redirect(request.referrer or url_for("ver_trabajadores"))
    if any(len(valor) > maximo for valor, maximo in (
        (nombres, 120), (apellido_paterno, 100), (apellido_materno, 100),
        (request.form.get("cargo", "").strip(), 150),
        (request.form.get("correo_personal", "").strip(), 254),
        (request.form.get("correo_institucional", "").strip(), 254),
    )):
        flash("Uno o más campos de texto superan la longitud permitida.", "warning")
        return redirect(request.referrer or url_for("ver_trabajadores"))
    if len(tipo_contrato) > 100:
        flash("El contrato o régimen laboral no puede superar los 100 caracteres.", "warning")
        return redirect(request.referrer or url_for("ver_trabajadores"))
    try:
        perfil_cargo_id = validar_id_formulario(perfil_cargo_raw, "perfil de cargo")
    except ValueError:
        flash("Seleccione un cargo con una matriz de accesos previamente validada.", "warning")
        return redirect(request.referrer or url_for("ver_trabajadores"))

    conn = get_db()
    ensure_access_matrix_schema(conn)
    conn.commit()
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE")
    try:
        perfil = cursor.execute("""
            SELECT id, cargo, funciones, autorizador, reportes_requeridos, accion_acceso
            FROM perfiles_cargo_ut
            WHERE id = ? AND activo = 1 AND validado = 1
              AND TRIM(funciones) != '' AND TRIM(autorizador) != ''
        """, (perfil_cargo_id,)).fetchone()
        if not perfil:
            conn.rollback()
            flash(
                "El cargo ya no tiene una matriz activa y validada. "
                "Revise los accesos por cargo antes de registrar al trabajador.",
                "warning",
            )
            return redirect(request.referrer or url_for("ver_trabajadores"))

        cargo = perfil["cargo"]
        if not _obtener_niveles_perfil_cargo_ut(cursor, perfil["id"]):
            conn.rollback()
            flash(
                f"El perfil del cargo «{cargo}» no tiene sistemas asignados. "
                "Configure al menos los accesos generales en la matriz.",
                "warning",
            )
            return redirect(request.referrer or url_for("ver_trabajadores"))

        existente = cursor.execute("SELECT id FROM trabajadores WHERE dni = ?", (dni,)).fetchone()
        trabajador_con_celular = buscar_trabajador_con_celular(
            cursor,
            celular_institucional,
            existente["id"] if existente else None,
        )
        if trabajador_con_celular:
            nombre = " ".join((
                trabajador_con_celular["nombres"],
                trabajador_con_celular["apellido_paterno"],
                trabajador_con_celular["apellido_materno"],
            )).strip()
            flash(
                f"El celular {celular_institucional} ya está registrado para {nombre} "
                f"(DNI {trabajador_con_celular['dni']}).",
                "danger",
            )
            return redirect(request.referrer or url_for("ver_trabajadores"))

        celular_normalizado = normalizar_numero_celular(celular_institucional)
        datos = (nombres, apellido_paterno, apellido_materno, cargo, correo_personal,
                 correo_institucional, tipo_contrato, telefono, celular_institucional,
                 fecha_nacimiento, celular_normalizado)
        if existente:
            cursor.execute("""
                UPDATE trabajadores
                SET nombres = ?, apellido_paterno = ?, apellido_materno = ?, cargo = ?,
                    correo_personal = ?, correo_institucional = ?, tipo_contrato = ?, telefono = ?,
                    celular_institucional = ?, fecha_nacimiento = ?, activo = 1,
                    celular_institucional_normalizado = ?
                WHERE id = ?
            """, datos + (existente["id"],))
            trabajador_id = existente["id"]
        else:
            cursor.execute("""
                INSERT INTO trabajadores (dni, nombres, apellido_paterno, apellido_materno, cargo,
                                          correo_personal, correo_institucional, tipo_contrato, telefono,
                                          celular_institucional, fecha_nacimiento, activo,
                                          celular_institucional_normalizado)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
            """, (dni,) + datos)
            trabajador_id = cursor.lastrowid

        _registrar_snapshot_accesos_trabajador_ut(
            cursor, trabajador_id, perfil, session.get("usuario_id")
        )
        conn.commit()
        flash(
            f"Trabajador {nombres} {apellido_paterno} registrado o reactivado. "
            f"Se cargó la matriz del cargo «{cargo}» y quedó pendiente de validación.",
            "success",
        )
    except Exception:
        conn.rollback()
        app.logger.exception("No se pudo registrar al trabajador y guardar su solicitud de accesos.")
        flash(f"No se pudo guardar al trabajador. Verifique si el DNI {dni} ya está registrado.", "danger")
    finally:
        conn.close()

    return redirect(request.referrer or url_for("ver_trabajadores"))

@app.route("/trabajadores/<int:trabajador_id>/ficha-accesos")
@permiso_requerido("personal", "ver")
@login_required
def ficha_accesos_trabajador(trabajador_id):
    conn = get_db()
    ensure_access_matrix_schema(conn)
    conn.commit()
    cursor = conn.cursor()
    trabajador = cursor.execute("""
        SELECT id, dni, nombres, apellido_paterno, apellido_materno, cargo,
               tipo_contrato, activo
        FROM trabajadores
        WHERE id = ?
    """, (trabajador_id,)).fetchone()
    if not trabajador:
        conn.close()
        flash("No se encontró al trabajador solicitado.", "warning")
        return redirect(url_for("ver_trabajadores"))

    solicitud = cursor.execute("""
        SELECT s.*
        FROM solicitudes_acceso_trabajador_ut s
        WHERE s.trabajador_id = ?
        ORDER BY s.id DESC
        LIMIT 1
    """, (trabajador_id,)).fetchone()
    asignaciones = []
    estado = "PERFIL DEL CARGO PENDIENTE DE VALIDACIÓN"
    responsable = ""
    funciones = ""
    if solicitud:
        asignaciones = [
            dict(row) for row in cursor.execute("""
                SELECT m.id, m.nombre, m.tipo, m.padre_id, a.nivel_acceso
                FROM modulos_solicitud_acceso_trabajador_ut a
                JOIN sistemas_modulos_ut m ON m.id = a.sistema_modulo_id
                WHERE a.solicitud_id = ?
                ORDER BY m.nombre COLLATE NOCASE
            """, (solicitud["id"],)).fetchall()
        ]
        estado = solicitud["estado"]
        responsable = solicitud["autorizador"]
        funciones = solicitud["funciones"]
    else:
        perfil = cursor.execute("""
            SELECT id, funciones, autorizador, validado
            FROM perfiles_cargo_ut
            WHERE LOWER(TRIM(cargo)) = LOWER(TRIM(?)) AND activo = 1
            LIMIT 1
        """, (trabajador["cargo"] or "",)).fetchone()
        if perfil and perfil["validado"]:
            niveles = {
                row["id"]: "CUENTA"
                for row in cursor.execute("""
                    SELECT id FROM sistemas_modulos_ut
                    WHERE activo = 1 AND obligatorio_todos = 1
                """).fetchall()
            }
            niveles.update({
                row["sistema_modulo_id"]: row["nivel_acceso"]
                for row in cursor.execute("""
                    SELECT a.sistema_modulo_id, a.nivel_acceso
                    FROM accesos_perfil_cargo_ut a
                    JOIN sistemas_modulos_ut m
                      ON m.id = a.sistema_modulo_id AND m.activo = 1
                    WHERE a.perfil_id = ?
                """, (perfil["id"],)).fetchall()
            })
            asignaciones = [
                dict(row, nivel_acceso=niveles[row["id"]])
                for row in cursor.execute("""
                    SELECT id, nombre, tipo, padre_id
                    FROM sistemas_modulos_ut
                    WHERE activo = 1
                """).fetchall()
                if row["id"] in niveles
            ]
            estado = "PROPUESTA SEGÚN MATRIZ DEL CARGO"
            responsable = perfil["autorizador"]
            funciones = perfil["funciones"]
        else:
            asignaciones = [
                dict(row, nivel_acceso="CUENTA")
                for row in cursor.execute("""
                    SELECT id, nombre, tipo, padre_id
                    FROM sistemas_modulos_ut
                    WHERE activo = 1 AND obligatorio_todos = 1
                """).fetchall()
            ]
    modulos = {
        row["id"]: dict(row)
        for row in cursor.execute("""
            SELECT id, nombre, padre_id
            FROM sistemas_modulos_ut
        """).fetchall()
    }
    conn.close()

    for asignacion in asignaciones:
        partes = [asignacion["nombre"]]
        padre_id = asignacion["padre_id"]
        visitados = {asignacion["id"]}
        while padre_id in modulos and padre_id not in visitados:
            padre = modulos[padre_id]
            partes.append(padre["nombre"])
            visitados.add(padre_id)
            padre_id = padre["padre_id"]
        asignacion["ruta"] = " / ".join(reversed(partes))
        asignacion["nivel_nombre"] = _etiqueta_nivel_acceso_ut(
            asignacion["nivel_acceso"]
        )
    asignaciones.sort(key=lambda item: item["ruta"].casefold())

    buffer = BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=42,
        leftMargin=42,
        topMargin=42,
        bottomMargin=42,
        title=f"Ficha de accesos - {trabajador['nombres']} {trabajador['apellido_paterno']}",
    )
    estilos = getSampleStyleSheet()
    estilo_celda = ParagraphStyle(
        "CeldaFichaAccesos",
        parent=estilos["BodyText"],
        fontSize=8,
        leading=10,
    )
    nombre_completo = " ".join(
        valor for valor in (
            trabajador["nombres"],
            trabajador["apellido_paterno"],
            trabajador["apellido_materno"],
        ) if valor
    )
    elementos = [
        Paragraph("FICHA DE SISTEMAS Y PERMISOS ASIGNADOS", estilos["Title"]),
        Spacer(1, 12),
        Paragraph(f"<b>Trabajador:</b> {escape_xml(nombre_completo)}", estilos["BodyText"]),
        Paragraph(f"<b>DNI:</b> {escape_xml(trabajador['dni'] or '-')}", estilos["BodyText"]),
        Paragraph(f"<b>Cargo:</b> {escape_xml(trabajador['cargo'] or 'Sin cargo registrado')}", estilos["BodyText"]),
        Paragraph(f"<b>Contrato / régimen:</b> {escape_xml(trabajador['tipo_contrato'] or '-')}", estilos["BodyText"]),
        Paragraph(f"<b>Estado de la matriz:</b> {escape_xml(estado)}", estilos["BodyText"]),
        Paragraph(f"<b>Responsable que valida:</b> {escape_xml(responsable or '-')}", estilos["BodyText"]),
        Paragraph(f"<b>Funciones del perfil:</b> {escape_xml(funciones or '-')}", estilos["BodyText"]),
        Spacer(1, 12),
    ]
    datos_tabla = [[
        Paragraph("<b>Sistema / módulo / aplicativo</b>", estilo_celda),
        Paragraph("<b>Nivel de acceso</b>", estilo_celda),
    ]]
    for asignacion in asignaciones:
        datos_tabla.append([
            Paragraph(escape_xml(asignacion["ruta"]), estilo_celda),
            Paragraph(escape_xml(asignacion["nivel_nombre"]), estilo_celda),
        ])
    if not asignaciones:
        datos_tabla.append([
            Paragraph("No hay sistemas registrados para este trabajador.", estilo_celda),
            "",
        ])
    tabla = Table(datos_tabla, colWidths=[365, 130], repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B7C9D6")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EFF4F8")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elementos.extend([
        tabla,
        Spacer(1, 14),
        Paragraph(
            "Esta ficha documenta los accesos requeridos; no crea cuentas ni concede permisos "
            "en los sistemas externos. La habilitación debe realizarse luego de la validación "
            "correspondiente.",
            estilos["Italic"],
        ),
    ])
    documento.build(elementos)
    buffer.seek(0)
    nombre_archivo = f"Ficha_Accesos_{trabajador_id}.pdf"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=nombre_archivo,
        mimetype="application/pdf",
    )

@app.route("/trabajadores/editar", methods=["POST"])
@permiso_requerido("personal", "modificar")
@login_required
def editar_trabajador():
    try:
        trabajador_id = validar_id_formulario(
            request.form.get("trabajador_id"), "trabajador"
        )
    except ValueError:
        flash("El identificador del trabajador no es válido.", "warning")
        return redirect(url_for("ver_trabajadores"))
    try:
        dni = validar_dni(request.form.get("dni"))
        correo_personal = validar_correo(
            request.form.get("correo_personal"), "El correo personal"
        )
        correo_institucional = validar_correo(
            request.form.get("correo_institucional"), "El correo institucional"
        )
        telefono = validar_numero_contacto(request.form.get("telefono"), "El teléfono")
        celular_institucional = validar_numero_contacto(
            request.form.get("celular_institucional"), "El celular institucional"
        )
        fecha_nacimiento = validar_fecha_formulario(
            request.form.get("fecha_nacimiento"), "la fecha de nacimiento"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_trabajadores"))

    nombres = request.form["nombres"].strip()
    apellido_paterno = request.form["apellido_paterno"].strip()
    apellido_materno = request.form["apellido_materno"].strip()
    cargo = request.form.get("cargo", "").strip()
    tipo_contrato = request.form.get("tipo_contrato", "").strip()
    if not all((nombres, apellido_paterno, apellido_materno)):
        flash("Nombres y ambos apellidos son obligatorios.", "warning")
        return redirect(url_for("ver_trabajadores"))
    if any(len(valor) > maximo for valor, maximo in (
        (nombres, 120), (apellido_paterno, 100), (apellido_materno, 100),
        (cargo, 150), (correo_personal, 254), (correo_institucional, 254),
    )):
        flash("Uno o más campos de texto superan la longitud permitida.", "warning")
        return redirect(url_for("ver_trabajadores"))
    if len(tipo_contrato) > 100:
        flash("El contrato o régimen laboral no puede superar los 100 caracteres.", "warning")
        return redirect(url_for("ver_trabajadores"))

    conn = get_db()
    ensure_access_matrix_schema(conn)
    conn.commit()
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE")
    trabajador_actual = cursor.execute(
        "SELECT cargo FROM trabajadores WHERE id = ?",
        (trabajador_id,),
    ).fetchone()
    if not trabajador_actual:
        conn.rollback()
        conn.close()
        flash("No se encontró al trabajador que desea editar.", "warning")
        return redirect(url_for("ver_trabajadores"))

    perfil_nuevo_cargo = None
    cargo_actual = (trabajador_actual["cargo"] or "").strip()
    if cargo.casefold() != cargo_actual.casefold():
        perfil_nuevo_cargo = cursor.execute("""
            SELECT id, cargo, funciones, autorizador, reportes_requeridos, accion_acceso
            FROM perfiles_cargo_ut
            WHERE LOWER(TRIM(cargo)) = LOWER(TRIM(?))
              AND activo = 1 AND validado = 1
              AND TRIM(funciones) != '' AND TRIM(autorizador) != ''
            LIMIT 1
        """, (cargo,)).fetchone()
        if not perfil_nuevo_cargo:
            conn.rollback()
            conn.close()
            flash(
                "Para cambiar el cargo, primero configure y valide su matriz de accesos.",
                "warning",
            )
            return redirect(url_for("ver_trabajadores"))
        if not _obtener_niveles_perfil_cargo_ut(cursor, perfil_nuevo_cargo["id"]):
            conn.rollback()
            conn.close()
            flash("El nuevo cargo no tiene sistemas activos asignados en su matriz.", "warning")
            return redirect(url_for("ver_trabajadores"))

    trabajador_con_celular = buscar_trabajador_con_celular(
        cursor, celular_institucional, trabajador_id
    )
    if trabajador_con_celular:
        nombre = " ".join((
            trabajador_con_celular["nombres"],
            trabajador_con_celular["apellido_paterno"],
            trabajador_con_celular["apellido_materno"],
        )).strip()
        conn.close()
        flash(
            f"El celular {celular_institucional} ya está registrado para {nombre} "
            f"(DNI {trabajador_con_celular['dni']}).",
            "danger",
        )
        return redirect(url_for("ver_trabajadores"))

    try:
        cursor.execute("""
            UPDATE trabajadores 
            SET dni = ?, nombres = ?, apellido_paterno = ?, apellido_materno = ?, cargo = ?,
                                correo_personal = ?, correo_institucional = ?, tipo_contrato = ?, telefono = ?,
                                celular_institucional = ?, fecha_nacimiento = ?,
                                celular_institucional_normalizado = ?
            WHERE id = ?
                """, (dni, nombres, apellido_paterno, apellido_materno, cargo, correo_personal,
                            correo_institucional, tipo_contrato, telefono, celular_institucional,
                            fecha_nacimiento, normalizar_numero_celular(celular_institucional),
                            trabajador_id))
        if perfil_nuevo_cargo:
            _registrar_snapshot_accesos_trabajador_ut(
                cursor,
                trabajador_id,
                perfil_nuevo_cargo,
                session.get("usuario_id"),
            )
        conn.commit()
        flash(
            f"Datos de {nombres} {apellido_paterno} actualizados correctamente."
            + (
                f" Se cargó la matriz validada del cargo «{perfil_nuevo_cargo['cargo']}»."
                if perfil_nuevo_cargo else ""
            ),
            "success",
        )
    except Exception:
        conn.rollback()
        app.logger.exception("No se pudieron guardar los datos o actualizar los accesos del trabajador.")
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
         SELECT e.codigo_margesi, e.categoria, e.descripcion, e.marca, e.modelo,
             e.numero_serie, e.celular_institucional
        FROM prestamos p
        JOIN equipos e ON e.id = p.equipo_id
        WHERE p.trabajador_id = ? AND p.estado = 'ACTIVO'
        ORDER BY e.categoria, e.codigo_margesi
    """, (trabajador_id,)).fetchall()

    if prestamos_activos:
        bienes_pendientes = []
        nombres_categoria = {
            "MOVILES": "Celulares y equipos móviles",
            "TOKENS": "Tokens y firmas digitales",
            "INFORMATICOS": "Equipos informáticos",
            "MOBILIARIO": "Mobiliario institucional",
        }
        for bien in prestamos_activos:
            detalle = " - ".join(str(valor) for valor in (
                nombres_categoria.get(bien["categoria"], bien["categoria"]),
                bien["codigo_margesi"], bien["descripcion"], bien["marca"],
                bien["modelo"], bien["numero_serie"],
                f"Línea: {bien['celular_institucional']}" if bien["celular_institucional"] else None
            ) if valor)
            bienes_pendientes.append(detalle or "Bien institucional sin descripción")
        flash("No se puede desactivar al trabajador. Registre la devolución de estos bienes: " + "; ".join(bienes_pendientes), "warning")
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
            celulares_duplicados = resultado.get("celulares_duplicados", [])
            filas_invalidas = resultado.get("filas_invalidas", 0)
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
            if celulares_duplicados:
                detalle = "; ".join(
                    f"DNI {fila['dni']} omitido (celular ya registrado para DNI {fila['dni_existente']})"
                    for fila in celulares_duplicados[:3]
                )
                if len(celulares_duplicados) > 3:
                    detalle += f" ... y {len(celulares_duplicados) - 3} más"
                flash(f"No se importaron números celulares duplicados: {detalle}.", "warning")
            if filas_invalidas:
                flash(f"Se omitieron {filas_invalidas} filas por DNI, teléfono, celular o fecha inválidos.", "warning")
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
            celulares_duplicados = resultado.get("celulares_duplicados", [])
            filas_invalidas = resultado.get("filas_invalidas", 0)
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
            if celulares_duplicados:
                detalle = "; ".join(
                    f"DNI {fila['dni']} omitido (celular ya registrado para DNI {fila['dni_existente']})"
                    for fila in celulares_duplicados[:3]
                )
                if len(celulares_duplicados) > 3:
                    detalle += f" ... y {len(celulares_duplicados) - 3} más"
                flash(f"No se importaron números celulares duplicados: {detalle}.", "warning")
            if filas_invalidas:
                flash(f"Se omitieron {filas_invalidas} filas por DNI, teléfono, celular o fecha inválidos.", "warning")
        else:
            flash(f"Se importaron / actualizaron {resultado} trabajadores exitosamente.", "success")
    else:
        flash("Formato inválido. Seleccione un archivo Excel (.xlsx) para personal.", "danger")
    return redirect(url_for("ver_trabajadores"))

# ================= GESTIÓN DE CERTIFICADOS DIGITALES =================
@app.route("/tokens")
@permiso_requerido("tokens", "ver")
@login_required
def ver_tokens():
    filtro = request.args.get("estado", "ACTIVOS").upper()
    condiciones = {
        "ACTIVOS": "estado = 'VIGENTE'",
        "RETIRADOS": "estado = 'RETIRADO'",
        "TODOS": "1 = 1",
        "ALERTAS": "estado = 'VIGENTE' AND fecha_caducidad <= ?",
    }
    if filtro not in condiciones:
        filtro = "ACTIVOS"

    hoy = date.today()
    conn = get_db()
    cursor = conn.cursor()
    parametros = []
    if filtro == "ALERTAS":
        parametros.append((date.today() + timedelta(days=DIAS_ALERTA_CERTIFICADO)).isoformat())
    certificados = cursor.execute(
        "SELECT * FROM certificados_digitales WHERE "
        + condiciones[filtro]
        + " ORDER BY CASE WHEN estado = 'VIGENTE' THEN 0 ELSE 1 END, fecha_caducidad, apellido_paterno, nombres",
        parametros,
    ).fetchall()
    alertas = cursor.execute("""
        SELECT * FROM certificados_digitales
        WHERE estado = 'VIGENTE' AND fecha_caducidad <= ?
        ORDER BY fecha_caducidad, apellido_paterno, nombres
    """, ((date.today() + timedelta(days=DIAS_ALERTA_CERTIFICADO)).isoformat(),)).fetchall()
    total_activos = cursor.execute(
        """
        SELECT COUNT(*) FROM certificados_digitales
        WHERE estado = 'VIGENTE' AND fecha_inicio <= ? AND fecha_caducidad >= ?
        """,
        (hoy.isoformat(), hoy.isoformat()),
    ).fetchone()[0]
    total_retirados = cursor.execute(
        "SELECT COUNT(*) FROM certificados_digitales WHERE estado = 'RETIRADO'"
    ).fetchone()[0]
    trabajadores = cursor.execute("""
        SELECT id, dni, nombres, apellido_paterno, apellido_materno, cargo,
               correo_personal, correo_institucional
        FROM trabajadores
        WHERE activo = 1
        ORDER BY apellido_paterno COLLATE NOCASE, apellido_materno COLLATE NOCASE, nombres COLLATE NOCASE
    """).fetchall()
    conn.close()

    filas_certificados = []
    for certificado in certificados:
        fila = dict(certificado)
        dias = (date.fromisoformat(fila["fecha_caducidad"]) - hoy).days
        fila["fecha_inicio_display"] = date.fromisoformat(fila["fecha_inicio"]).strftime("%d/%m/%Y")
        fila["fecha_caducidad_display"] = date.fromisoformat(fila["fecha_caducidad"]).strftime("%d/%m/%Y")
        fila["dias_restantes"] = dias
        fila["alerta"] = (
            "RETIRADO" if fila["estado"] == "RETIRADO"
            else "VENCIDO" if dias < 0
            else "POR INICIAR" if date.fromisoformat(fila["fecha_inicio"]) > hoy
            else "POR VENCER" if dias <= DIAS_ALERTA_CERTIFICADO
            else "VIGENTE"
        )
        filas_certificados.append(fila)

    filas_alerta = []
    for certificado in alertas:
        fila = dict(certificado)
        fila["dias_restantes"] = (date.fromisoformat(fila["fecha_caducidad"]) - hoy).days
        fila["fecha_caducidad_display"] = date.fromisoformat(fila["fecha_caducidad"]).strftime("%d/%m/%Y")
        filas_alerta.append(fila)

    return render_template(
        "tokens.html",
        certificados=filas_certificados,
        alertas=filas_alerta,
        filtro=filtro,
        total_activos=total_activos,
        total_retirados=total_retirados,
        total_alertas=len(filas_alerta),
        dias_alerta=DIAS_ALERTA_CERTIFICADO,
        trabajadores=trabajadores,
    )


@app.route("/tokens/nuevo", methods=["POST"])
@permiso_requerido("tokens", "modificar")
@login_required
def nuevo_token():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE")
    try:
        datos = datos_certificado_desde_formulario(
            request.form, cursor, exigir_titular=True
        )
    except ValueError as error:
        conn.close()
        flash(str(error), "danger")
        return redirect(url_for("ver_tokens"))

    if buscar_certificado_duplicado(cursor, datos):
        conn.close()
        flash("Ese certificado ya está registrado con las mismas fechas para esta persona.", "danger")
        return redirect(url_for("ver_tokens"))
    cursor.execute("""
        INSERT INTO certificados_digitales (
            dni, apellido_paterno, apellido_materno, nombres, cargo, oficina,
            correo, tipo_certificado, trabajador_id, numero_resolucion,
            fecha_inicio, fecha_caducidad
        ) VALUES (
            :dni, :apellido_paterno, :apellido_materno, :nombres, :cargo, :oficina,
            :correo, :tipo_certificado, :trabajador_id, :numero_resolucion,
            :fecha_inicio, :fecha_caducidad
        )
    """, datos)
    conn.commit()
    conn.close()
    flash("Certificado digital registrado correctamente.", "success")
    return redirect(url_for("ver_tokens"))


@app.route("/tokens/editar/<int:certificado_id>", methods=["POST"])
@permiso_requerido("tokens", "modificar")
@login_required
def editar_token(certificado_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE")
    existe = cursor.execute(
        "SELECT id FROM certificados_digitales WHERE id = ?", (certificado_id,)
    ).fetchone()
    if not existe:
        conn.close()
        flash("No se encontró el certificado que desea editar.", "warning")
        return redirect(url_for("ver_tokens"))
    try:
        datos = datos_certificado_desde_formulario(
            request.form, cursor, exigir_titular=True
        )
    except ValueError as error:
        conn.close()
        flash(str(error), "danger")
        return redirect(url_for("ver_tokens"))
    if buscar_certificado_duplicado(cursor, datos, certificado_id):
        conn.close()
        flash("Ya existe otro registro para esa persona, tipo y periodo de certificado.", "danger")
        return redirect(url_for("ver_tokens"))

    cursor.execute("""
        UPDATE certificados_digitales
        SET dni = :dni, apellido_paterno = :apellido_paterno,
            apellido_materno = :apellido_materno, nombres = :nombres, cargo = :cargo,
            oficina = :oficina, correo = :correo, tipo_certificado = :tipo_certificado,
            trabajador_id = :trabajador_id, numero_resolucion = :numero_resolucion,
            fecha_inicio = :fecha_inicio, fecha_caducidad = :fecha_caducidad,
            actualizado_el = CURRENT_TIMESTAMP
        WHERE id = :id
    """, dict(datos, id=certificado_id))
    conn.commit()
    conn.close()
    flash("Certificado digital actualizado correctamente.", "success")
    return redirect(url_for("ver_tokens"))


@app.route("/tokens/retirar/<int:certificado_id>", methods=["POST"])
@permiso_requerido("tokens", "eliminar")
@login_required
def retirar_token(certificado_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE certificados_digitales
        SET estado = 'RETIRADO', actualizado_el = CURRENT_TIMESTAMP
        WHERE id = ? AND estado = 'VIGENTE'
    """, (certificado_id,))
    if cursor.rowcount:
        conn.commit()
        flash("Certificado retirado. Se conserva su registro para mantener el historial.", "success")
    else:
        conn.rollback()
        flash("No se encontró un certificado vigente con ese identificador.", "warning")
    conn.close()
    return redirect(url_for("ver_tokens"))


@app.route("/tokens/importar-excel", methods=["POST"])
@permiso_requerido("tokens", "modificar")
@login_required
def importar_tokens_excel():
    archivo = request.files.get("archivo_excel_tokens")
    if not archivo or not archivo.filename or not archivo.filename.lower().endswith(".xlsx"):
        flash("Seleccione un archivo Excel válido (.xlsx).", "danger")
        return redirect(url_for("ver_tokens"))

    try:
        libro = openpyxl.load_workbook(archivo, data_only=True, read_only=True)
    except Exception:
        flash("No se pudo leer el archivo Excel. Verifique que sea un .xlsx válido.", "danger")
        return redirect(url_for("ver_tokens"))

    aliases = {
        "dni": {"dni", "documento", "documento_de_identidad", "identificacion"},
        "apellido_paterno": {"apellido_paterno", "apell_paterno", "paterno"},
        "apellido_materno": {"apellido_materno", "apell_materno", "materno"},
        "nombres": {"nombres", "nombre"},
        "cargo": {"cargo", "puesto"},
        "oficina": {"oficina", "area", "dependencia"},
        "correo": {"correo", "email", "correo_electronico"},
        "tipo_certificado": {"tipo_certificado", "tipo_de_certificado", "certificado", "tipo"},
        "tipo_titular": {"tipo_titular", "tipo_de_titular", "titular", "origen"},
        "numero_resolucion": {
            "numero_resolucion", "numero_de_resolucion", "nro_resolucion",
            "n_de_resolucion", "resolucion",
        },
        "fecha_inicio": {"fecha_inicio", "inicio", "fecha_de_inicio", "fecha_emision"},
        "fecha_caducidad": {"fecha_caducidad", "fecha_expiracion", "fecha_de_expiracion", "caducidad", "vencimiento"},
    }
    requeridos = {
        "dni", "apellido_paterno", "apellido_materno", "nombres",
        "tipo_certificado", "fecha_inicio", "fecha_caducidad",
    }

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("BEGIN IMMEDIATE")
    importados = 0
    duplicados = 0
    errores_filas = []

    for hoja in libro.worksheets:
        encabezado_encontrado = None
        mapa_columnas = {}
        for numero_fila, fila in enumerate(hoja.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
            candidatos = {
                campo: indice
                for indice, valor in enumerate(fila)
                for campo, nombres in aliases.items()
                if normalizar_encabezado_certificado(valor) in nombres
            }
            if requeridos <= candidatos.keys():
                encabezado_encontrado = numero_fila
                mapa_columnas = candidatos
                break
        if encabezado_encontrado is None:
            errores_filas.append(f"{hoja.title}: no se detectaron los encabezados requeridos")
            continue

        for numero_fila, fila in enumerate(
            hoja.iter_rows(min_row=encabezado_encontrado + 1, values_only=True),
            start=encabezado_encontrado + 1,
        ):
            if not any(valor is not None and str(valor).strip() for valor in fila):
                continue
            datos_fila = {}
            for campo, indice in mapa_columnas.items():
                if campo in ("fecha_inicio", "fecha_caducidad"):
                    continue
                valor = fila[indice] if indice < len(fila) else None
                if isinstance(valor, float) and valor.is_integer():
                    valor = int(valor)
                datos_fila[campo] = str(valor).strip() if valor is not None else ""
            try:
                for campo in ("fecha_inicio", "fecha_caducidad"):
                    indice = mapa_columnas.get(campo)
                    datos_fila[campo] = fecha_certificado_iso(
                        fila[indice] if indice is not None and indice < len(fila) else None
                    )
                datos = datos_certificado_desde_formulario(datos_fila, cursor)
            except ValueError as error:
                errores_filas.append(f"{hoja.title}, fila {numero_fila}: {error}")
                continue

            if buscar_certificado_duplicado(cursor, datos):
                duplicados += 1
                continue
            cursor.execute("""
                INSERT INTO certificados_digitales (
                    dni, apellido_paterno, apellido_materno, nombres, cargo, oficina,
                        correo, tipo_certificado, trabajador_id, numero_resolucion,
                        fecha_inicio, fecha_caducidad
                    ) VALUES (
                        :dni, :apellido_paterno, :apellido_materno, :nombres, :cargo, :oficina,
                        :correo, :tipo_certificado, :trabajador_id, :numero_resolucion,
                        :fecha_inicio, :fecha_caducidad
                    )
            """, datos)
            importados += 1

    libro.close()
    conn.commit()
    conn.close()
    flash(f"Importación terminada: {importados} certificados agregados, {duplicados} duplicados omitidos.", "success")
    if errores_filas:
        detalle = "; ".join(errores_filas[:4])
        if len(errores_filas) > 4:
            detalle += f"; y {len(errores_filas) - 4} errores más"
        flash(f"Algunas filas no se importaron: {detalle}.", "warning")
    return redirect(url_for("ver_tokens"))


@app.route("/tokens/exportar-excel")
@permiso_requerido("tokens", "reportar")
@login_required
def exportar_tokens_excel():
    filtro = request.args.get("estado", "TODOS").upper()
    condiciones = {
        "TODOS": "1 = 1",
        "ACTIVOS": "estado = 'VIGENTE'",
        "RETIRADOS": "estado = 'RETIRADO'",
        "ALERTAS": "estado = 'VIGENTE' AND fecha_caducidad <= ?",
    }
    if filtro not in condiciones:
        filtro = "TODOS"
    parametros = (
        ((date.today() + timedelta(days=DIAS_ALERTA_CERTIFICADO)).isoformat(),)
        if filtro == "ALERTAS" else ()
    )
    conn = get_db()
    certificados = conn.execute(
        "SELECT * FROM certificados_digitales WHERE "
        + condiciones[filtro]
        + " ORDER BY fecha_caducidad, apellido_paterno, nombres",
        parametros,
    ).fetchall()
    conn.close()

    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Certificados digitales"
    encabezados = [
        "N°", "DNI", "APELL. PATERNO", "APELL. MATERNO", "NOMBRES", "CARGO",
        "OFICINA", "CORREO", "TITULAR", "N° DE RESOLUCIÓN", "TIPO DE CERTIFICADO", "FECHA DE INICIO",
        "FECHA EXPIRACIÓN", "ESTADO",
    ]
    hoja.append(encabezados)
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    border = Border(
        left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"),
    )
    for celda in hoja[1]:
        celda.fill = header_fill
        celda.font = header_font
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = border

    hoy = date.today()
    for numero, certificado in enumerate(certificados, start=1):
        caducidad = date.fromisoformat(certificado["fecha_caducidad"])
        dias = (caducidad - hoy).days
        alerta = (
            "RETIRADO" if certificado["estado"] == "RETIRADO"
            else "VENCIDO" if dias < 0
            else "POR INICIAR" if date.fromisoformat(certificado["fecha_inicio"]) > hoy
            else "POR VENCER" if dias <= DIAS_ALERTA_CERTIFICADO
            else "VIGENTE"
        )
        hoja.append([
            numero, certificado["dni"], certificado["apellido_paterno"],
            certificado["apellido_materno"], certificado["nombres"], certificado["cargo"],
            certificado["oficina"], certificado["correo"],
            (
                "PERSONAL" if certificado["trabajador_id"]
                else "MIEMBRO DEL COMITÉ" if certificado["numero_resolucion"]
                else "REGISTRO PREVIO"
            ),
            certificado["numero_resolucion"], certificado["tipo_certificado"],
            certificado["fecha_inicio"], certificado["fecha_caducidad"], alerta,
        ])
        for celda in hoja[hoja.max_row]:
            celda.border = border
            celda.alignment = Alignment(vertical="center", wrap_text=True)
    for columna in hoja.columns:
        largo = max(len(str(celda.value or "")) for celda in columna)
        letra = openpyxl.utils.get_column_letter(columna[0].column)
        hoja.column_dimensions[letra].width = min(max(largo + 3, 12), 32)
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    buffer = BytesIO()
    libro.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"Certificados_Digitales_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

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
    try:
        f_desde, f_hasta = validar_rango_fechas(
            request.args.get("desde", ""), request.args.get("hasta", ""), "reporte"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_reportes"))
    
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
    try:
        f_desde, f_hasta = validar_rango_fechas(
            request.args.get("desde", ""), request.args.get("hasta", ""), "reporte"
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("ver_reportes"))

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
