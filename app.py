import os
import json
import hashlib
import urllib.parse
from io import BytesIO
from datetime import datetime, date, timedelta
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, session
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from database import get_db, init_db
from excel_loader import importar_trabajadores_excel, formatear_solo_fecha
from equipos_loader import importar_equipos_excel
from docx_generator import rellenar_plantilla_docx
from sgd_processor import procesar_sgd_excel, generar_excel_sgd

app = Flask(__name__)
app.secret_key = "clave_secreta_uti_apurimac_segura"

UPLOAD_FOLDER = "uploads"
ACTAS_FOLDER = "actas_generadas"
PLANTILLAS_FOLDER = "plantillas"
SGD_FOLDER = "uploads_sgd"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(ACTAS_FOLDER, exist_ok=True)
os.makedirs(PLANTILLAS_FOLDER, exist_ok=True)
os.makedirs(SGD_FOLDER, exist_ok=True)

def hash_pass(p):
    return hashlib.sha256(p.encode()).hexdigest()

def check_pass(p, h):
    return hash_pass(p) == h

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
    return dict(calcular_duracion=calcular_duracion)

@app.route("/favicon.ico")
def favicon():
    return redirect(url_for("static", filename="logo-pae.svg", v="2"))

# ================= LOGIN Y AUTENTICACIÓN =================
@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("usuario_id"):
        return redirect(url_for("inicio"))

    if request.method == "POST":
        user_input = request.form["usuario"].strip()
        pass_input = request.form["password"].strip()

        conn = get_db()
        cursor = conn.cursor()
        usuario_db = cursor.execute("SELECT * FROM usuarios_sistema WHERE usuario = ?", (user_input,)).fetchone()
        conn.close()

        if usuario_db and check_pass(pass_input, usuario_db["password"]):
            session["usuario_id"] = usuario_db["id"]
            session["usuario"] = usuario_db["usuario"]
            session["nombre"] = usuario_db["nombre_completo"]
            session["rol"] = usuario_db["rol"]
            flash(f"¡Bienvenido(a), {usuario_db['nombre_completo']}!", "success")
            return redirect(url_for("inicio"))
        else:
            flash("Usuario o contraseña incorrectos.", "danger")

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    flash("Ha cerrado sesión correctamente.", "info")
    return redirect(url_for("login"))

# ================= GESTIÓN DE USUARIOS (SOLO ADMINISTRADOR) =================
@app.route("/usuarios")
@admin_required
def ver_usuarios():
    conn = get_db()
    usuarios = conn.cursor().execute("SELECT id, usuario, nombre_completo, rol, creado_el FROM usuarios_sistema ORDER BY id ASC").fetchall()
    conn.close()
    return render_template("usuarios.html", usuarios=usuarios)

@app.route("/usuarios/nuevo", methods=["POST"])
@admin_required
def nuevo_usuario():
    user_input = request.form["usuario"].strip()
    nombre_input = request.form["nombre_completo"].strip()
    pass_input = request.form["password"].strip()
    rol_input = request.form.get("rol", "USUARIO").strip().upper()

    if not user_input or not pass_input or not nombre_input:
        flash("Todos los campos obligatorios deben ser completados.", "warning")
        return redirect(url_for("ver_usuarios"))

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO usuarios_sistema (usuario, password, nombre_completo, rol)
            VALUES (?, ?, ?, ?)
        """, (user_input, hash_pass(pass_input), nombre_input, rol_input))
        conn.commit()
        flash(f"Usuario '{user_input}' creado exitosamente.", "success")
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

    conn = get_db()
    cursor = conn.cursor()

    if pass_input:
        cursor.execute("""
            UPDATE usuarios_sistema 
            SET nombre_completo = ?, password = ?, rol = ?
            WHERE id = ?
        """, (nombre_input, hash_pass(pass_input), rol_input, user_id))
    else:
        cursor.execute("""
            UPDATE usuarios_sistema 
            SET nombre_completo = ?, rol = ?
            WHERE id = ?
        """, (nombre_input, rol_input, user_id))

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
    total_equipos = total_moviles + total_tokens + total_informaticos
    equipos_prestamo = cursor.execute("SELECT COUNT(*) FROM equipos WHERE estado = 'EN PRESTAMO'").fetchone()[0]
    total_personal = cursor.execute("SELECT COUNT(*) FROM trabajadores").fetchone()[0]
    
    tareas_pendientes = cursor.execute("SELECT COUNT(*) FROM actividades_soporte WHERE estado != 'COMPLETADA'").fetchone()[0]
    tareas_hoy = cursor.execute("SELECT COUNT(*) FROM actividades_soporte WHERE fecha_programada = CURRENT_DATE").fetchone()[0]
    conn.close()

    return render_template("inicio.html", 
                           total_equipos=total_equipos,
                           total_moviles=total_moviles,
                           total_tokens=total_tokens,
                           total_informaticos=total_informaticos,
                           equipos_prestamo=equipos_prestamo,
                           total_personal=total_personal,
                           tareas_pendientes=tareas_pendientes,
                           tareas_hoy=tareas_hoy)

# ================= MÓDULO: MONITOR DE SGD (HISTORIAL POR FECHAS DE CORTE EN BD) =================
@app.route("/sgd")
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
                resultado = json.loads(corte_db["datos_json"])
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

    nombre_guardado = f"sgd_{tipo_key.lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    ruta_guardado = os.path.join(SGD_FOLDER, nombre_guardado)
    archivo.save(ruta_guardado)

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

@app.route("/sgd/exportar-excel/<tipo>")
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
    fecha_corte_str = corte_db["fecha_corte"]
    buffer = generar_excel_sgd(resultado, nombre_estado, fecha_corte_str)
    nombre_descarga = f"Reporte_SGD_{tipo_key}_{corte_db['id']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/sgd/eliminar-corte/<int:corte_id>", methods=["POST"])
@admin_required
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

# ================= MÓDULO: REGISTRO DE ACTIVIDADES (SOPORTE TI) =================
@app.route("/actividades")
@login_required
def ver_actividades():
    filtro_estado = request.args.get("estado", "TODOS")
    filtro_cat = request.args.get("categoria", "TODAS")
    conn = get_db()
    cursor = conn.cursor()

    base_query = "SELECT * FROM actividades_soporte WHERE 1=1"
    params = []

    if filtro_estado != "TODOS":
        base_query += " AND estado = ?"
        params.append(filtro_estado)
    if filtro_cat != "TODAS":
        base_query += " AND categoria = ?"
        params.append(filtro_cat)

    base_query += " ORDER BY fecha_programada DESC, hora_inicio ASC, id DESC"
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
        ORDER BY nombres ASC
    """).fetchall()

    total_pendientes = cursor.execute("SELECT COUNT(*) FROM actividades_soporte WHERE estado = 'PENDIENTE'").fetchone()[0]
    total_proceso = cursor.execute("SELECT COUNT(*) FROM actividades_soporte WHERE estado = 'EN PROGRESO'").fetchone()[0]
    total_completadas = cursor.execute("SELECT COUNT(*) FROM actividades_soporte WHERE estado = 'COMPLETADA'").fetchone()[0]
    total_general = cursor.execute("SELECT COUNT(*) FROM actividades_soporte").fetchone()[0]
    
    ahora_dt = datetime.now()
    hora_actual = ahora_dt.strftime("%H:%M")
    hora_fin_estimada = (ahora_dt + timedelta(hours=1)).strftime("%H:%M")
    conn.close()

    return render_template("actividades.html",
                           actividades=actividades,
                           trabajadores=trabajadores,
                           filtro_estado=filtro_estado,
                           filtro_cat=filtro_cat,
                           total_pendientes=total_pendientes,
                           total_proceso=total_proceso,
                           total_completadas=total_completadas,
                           total_general=total_general,
                           hoy=datetime.now().strftime("%Y-%m-%d"),
                           hora_actual=hora_actual,
                           hora_fin_estimada=hora_fin_estimada)

@app.route("/actividades/nueva", methods=["POST"])
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
        INSERT INTO actividades_soporte (titulo, descripcion, categoria, area_solicitante, prioridad, estado, fecha_programada, hora_inicio, hora_fin)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (titulo, descripcion, categoria, area, prioridad, estado, fecha_prog, h_inicio, h_fin))
    conn.commit()
    conn.close()

    flash("Actividad registrada exitosamente.", "success")
    return redirect(url_for("ver_actividades"))

@app.route("/actividades/editar", methods=["POST"])
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
    conn.cursor().execute("""
        UPDATE actividades_soporte 
        SET titulo = ?, descripcion = ?, categoria = ?, area_solicitante = ?, prioridad = ?,
            estado = ?, fecha_programada = ?, hora_inicio = ?, hora_fin = ?, solucion_aplicada = ?
        WHERE id = ?
    """, (titulo, descripcion, categoria, area, prioridad, estado, fecha_prog, h_inicio, h_fin, solucion, actividad_id))
    conn.commit()
    conn.close()

    flash("Actividad actualizada correctamente.", "success")
    return redirect(url_for("ver_actividades"))

@app.route("/actividades/estado/<int:actividad_id>/<nuevo_estado>")
@login_required
def cambiar_estado_actividad(actividad_id, nuevo_estado):
    if nuevo_estado in ["PENDIENTE", "EN PROGRESO", "COMPLETADA"]:
        conn = get_db()
        conn.cursor().execute("UPDATE actividades_soporte SET estado = ? WHERE id = ?", (nuevo_estado, actividad_id))
        conn.commit()
        conn.close()
        flash(f"Estado de la actividad cambiado a {nuevo_estado}.", "info")
    return redirect(request.referrer or url_for("ver_actividades"))

@app.route("/actividades/acciones-realizadas", methods=["POST"])
@login_required
def guardar_acciones_realizadas():
    actividad_id = request.form["actividad_id"]
    acciones = request.form.get("solucion_aplicada", "").strip()

    if not acciones:
        flash("Debe registrar las acciones realizadas antes de completar la actividad.", "warning")
        return redirect(url_for("ver_actividades"))

    conn = get_db()
    conn.cursor().execute("""
        UPDATE actividades_soporte
        SET estado = 'COMPLETADA', solucion_aplicada = ?
        WHERE id = ?
    """, (acciones, actividad_id))
    conn.commit()
    conn.close()

    flash("Acciones realizadas guardadas correctamente.", "success")
    return redirect(url_for("ver_actividades"))

@app.route("/actividades/eliminar/<int:actividad_id>", methods=["POST"])
@admin_required
def eliminar_actividad(actividad_id):
    conn = get_db()
    conn.cursor().execute("DELETE FROM actividades_soporte WHERE id = ?", (actividad_id,))
    conn.commit()
    conn.close()
    flash("Actividad eliminada de la bitácora.", "success")
    return redirect(url_for("ver_actividades"))

@app.route("/api/actividades-calendario")
@login_required
def api_actividades_calendario():
    conn = get_db()
    filas = conn.cursor().execute("SELECT * FROM actividades_soporte").fetchall()
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
@login_required
def exportar_actividades_excel():
    conn = get_db()
    cursor = conn.cursor()
    filas = cursor.execute("""
        SELECT fecha_programada, hora_inicio, hora_fin, area_solicitante, categoria, 
               titulo, descripcion, prioridad, estado, solucion_aplicada
        FROM actividades_soporte
        ORDER BY fecha_programada DESC, hora_inicio ASC
    """).fetchall()
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
            formatear_solo_fecha(f["fecha_programada"]),
            f["hora_inicio"] or "",
            f["hora_fin"] or "",
            f["area_solicitante"] or "",
            f["categoria"] or "",
            f["titulo"] or "",
            f["descripcion"] or "",
            f["prioridad"] or "",
            f["estado"] or "",
            f["solucion_aplicada"] or ""
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

# ================= HUB DE PRÉSTAMOS =================
@app.route("/prestamos")
@login_required
def prestamos_hub():
    conn = get_db()
    cursor = conn.cursor()
    
    moviles = cursor.execute("SELECT COUNT(*) AS total, SUM(CASE WHEN estado='EN PRESTAMO' THEN 1 ELSE 0 END) AS prestados FROM equipos WHERE categoria='MOVILES'").fetchone()
    tokens = cursor.execute("SELECT COUNT(*) AS total, SUM(CASE WHEN estado='EN PRESTAMO' THEN 1 ELSE 0 END) AS prestados FROM equipos WHERE categoria='TOKENS'").fetchone()
    informaticos = cursor.execute("SELECT COUNT(*) AS total, SUM(CASE WHEN estado='EN PRESTAMO' THEN 1 ELSE 0 END) AS prestados FROM equipos WHERE categoria='INFORMATICOS'").fetchone()
    conn.close()

    return render_template("prestamos_hub.html", moviles=moviles, tokens=tokens, informaticos=informaticos)

# ================= VISTA DE CADA SUBMÓDULO DE PRÉSTAMO =================
@app.route("/prestamos/<categoria>")
@login_required
def ver_submodulo(categoria):
    cat = categoria.upper()
    if cat not in ["MOVILES", "TOKENS", "INFORMATICOS"]:
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
        "INFORMATICOS": "Submódulo de Equipos Informáticos y Audiovisuales"
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
@admin_required
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
@login_required
def prestar_equipo():
    equipo_id = request.form["equipo_id"]
    trabajador_id = request.form["trabajador_id"]
    observaciones = request.form.get("observaciones", "").strip()

    conn = get_db()
    cursor = conn.cursor()

    eq = cursor.execute("SELECT * FROM equipos WHERE id = ?", (equipo_id,)).fetchone()
    tr = cursor.execute("SELECT * FROM trabajadores WHERE id = ?", (trabajador_id,)).fetchone()

    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    cod_mostrar = eq["codigo_margesi"] or eq["codigo_interno"] or "EQ"
    nombre_archivo = f"Acta_{cod_mostrar}_{tr['dni']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
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

    if os.path.exists(ruta_plantilla):
        rellenar_plantilla_docx(ruta_plantilla, datos_word, ruta_salida)

    cursor.execute("""
        INSERT INTO prestamos (equipo_id, trabajador_id, fecha_prestamo, estado, observaciones_entrega, archivo_acta_entrega)
        VALUES (?, ?, CURRENT_TIMESTAMP, 'ACTIVO', ?, ?)
    """, (equipo_id, trabajador_id, observaciones, nombre_archivo))

    cursor.execute("UPDATE equipos SET estado = 'EN PRESTAMO' WHERE id = ?", (equipo_id,))
    conn.commit()
    conn.close()

    flash(f"Préstamo registrado para {nombre_completo}.", "success")
    return redirect(request.referrer or url_for("prestamos_hub"))

@app.route("/devolver", methods=["GET", "POST"])
@app.route("/prestamos/devolver", methods=["GET", "POST"])
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
@login_required
def ver_trabajadores():
    conn = get_db()
    cursor = conn.cursor()

    trabajadores = cursor.execute("""
        SELECT * FROM trabajadores 
        ORDER BY apellido_paterno ASC, apellido_materno ASC, nombres ASC
    """).fetchall()
    conn.close()
    return render_template("trabajadores.html", trabajadores=trabajadores)

@app.route("/trabajadores/nuevo", methods=["POST"])
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
        cursor.execute("""
            INSERT INTO trabajadores (dni, nombres, apellido_paterno, apellido_materno, cargo,
                                      correo_personal, correo_institucional, tipo_contrato, telefono,
                                      celular_institucional, fecha_nacimiento)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (dni, nombres, apellido_paterno, apellido_materno, cargo, correo_personal,
              correo_institucional, tipo_contrato, telefono, celular_institucional, fecha_nacimiento))
        conn.commit()
        flash(f"Trabajador {nombres} {apellido_paterno} registrado correctamente.", "success")
    except Exception:
        flash(f"Error: El DNI {dni} ya se encuentra registrado.", "danger")
    finally:
        conn.close()

    return redirect(request.referrer or url_for("ver_trabajadores"))

@app.route("/trabajadores/editar", methods=["POST"])
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
@admin_required
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
        cursor.execute("DELETE FROM trabajadores WHERE id = ?", (trabajador_id,))
        conn.commit()
        flash("Trabajador eliminado del padrón correctamente.", "success")

    conn.close()
    return redirect(url_for("ver_trabajadores"))

@app.route("/trabajadores/exportar-excel")
@login_required
def exportar_trabajadores_excel():
    conn = get_db()
    cursor = conn.cursor()
    filas = cursor.execute("""
        SELECT nombres, apellido_paterno, apellido_materno, dni, cargo,
             correo_personal, correo_institucional, tipo_contrato, telefono,
             celular_institucional, fecha_nacimiento
        FROM trabajadores
        ORDER BY apellido_paterno ASC, apellido_materno ASC, nombres ASC
    """).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Personal"

    headers = [
        "NOMBRES", "APELLIDO PATERNO", "APELLIDO MATERNO", "DNI", "CARGO",
        "CORREO ELECTRÓNICO PERSONAL", "CORREO ELECTRÓNICO INSTITUCIONAL",
        "TIPO DE CONTRATO", "NÚMERO DE TELÉFONO", "CELULAR INSTITUCIONAL",
        "FECHA DE NACIMIENTO"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
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

    cols_trabajadores = (4, 9, 10)
    for r_idx, f in enumerate(filas, start=2):
        row_data = [
            f["nombres"] or "",
            f["apellido_paterno"] or "",
            f["apellido_materno"] or "",
            str(f["dni"] or ""),
            f["cargo"] or "",
            f["correo_personal"] or "",
            f["correo_institucional"] or "",
            f["tipo_contrato"] or "",
            str(f["telefono"] or ""),
            str(f["celular_institucional"] or ""),
            formatear_solo_fecha(f["fecha_nacimiento"])
        ]
        ws.append(row_data)
        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=r_idx, column=col_num)
            c.border = thin_border
            if col_num in cols_trabajadores:
                c.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    nombre_descarga = f"Lista_Trabajadores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=nombre_descarga, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/importar-excel", methods=["POST"])
@login_required
def subir_excel():
    archivo = request.files.get("archivo_excel")
    if archivo and archivo.filename.endswith((".xlsx", ".xls")):
        ruta = os.path.join(UPLOAD_FOLDER, "personal_actualizado.xlsx")
        archivo.save(ruta)
        total = importar_trabajadores_excel(ruta)
        flash(f"Se importaron / actualizaron {total} trabajadores exitosamente.", "success")
    else:
        flash("Formato inválido. Seleccione un archivo Excel (.xlsx).", "danger")
    return redirect(request.referrer or url_for("ver_trabajadores"))

# ================= MÓDULO DE REPORTES Y TRAZABILIDAD =================
@app.route("/reportes")
@login_required
def ver_reportes():
    filtro_estado = request.args.get("estado", "TODOS")
    categoria = request.args.get("categoria", "TODAS").upper()
    if categoria not in ["TODAS", "MOVILES", "TOKENS", "INFORMATICOS"]:
        categoria = "TODAS"
    periodo = request.args.get("periodo", "TODOS").upper()
    f_desde = request.args.get("desde", "")
    f_hasta = request.args.get("hasta", "")
    
    conn = get_db()
    cursor = conn.cursor()

    categoria_where = "" if categoria == "TODAS" else " AND categoria = ?"
    categoria_param = [] if categoria == "TODAS" else [categoria]

    total_equipos = cursor.execute(
        "SELECT COUNT(*) FROM equipos WHERE categoria IN ('MOVILES', 'TOKENS', 'INFORMATICOS')" + categoria_where,
        categoria_param
    ).fetchone()[0]
    total_almacen = cursor.execute(
        "SELECT COUNT(*) FROM equipos WHERE estado = 'EN ALMACEN' AND categoria IN ('MOVILES', 'TOKENS', 'INFORMATICOS')" + categoria_where,
        categoria_param
    ).fetchone()[0]
    total_prestamo = cursor.execute(
        "SELECT COUNT(*) FROM equipos WHERE estado = 'EN PRESTAMO' AND categoria IN ('MOVILES', 'TOKENS', 'INFORMATICOS')" + categoria_where,
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
        WHERE estado = 'EN ALMACEN' AND categoria IN ('MOVILES', 'TOKENS', 'INFORMATICOS')
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
@login_required
def exportar_reporte_excel():
    filtro_estado = request.args.get("estado", "TODOS")
    categoria = request.args.get("categoria", "TODAS").upper()
    if categoria not in ["TODAS", "MOVILES", "TOKENS", "INFORMATICOS"]:
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
    app.run(host="0.0.0.0", port=5000, debug=False)
