import openpyxl
import re
from datetime import datetime, date
from database import get_db

def limpiar_texto(txt):
    if txt is None:
        return ""
    txt = str(txt).strip().upper()
    txt = txt.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    txt = re.sub(r'[\r\n\t]+', ' ', txt)
    txt = re.sub(r'\s+', ' ', txt)
    return txt.strip()

def formatear_solo_fecha(val):
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%d/%m/%Y")
    
    val_str = str(val).strip()
    if val_str.upper() == "NONE" or not val_str:
        return ""
    
    # 1. Si viene como YYYY-MM-DD o YYYY-MM-DD HH:MM:SS
    match_iso = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', val_str)
    if match_iso:
        anio, mes, dia = match_iso.group(1), match_iso.group(2).zfill(2), match_iso.group(3).zfill(2)
        return f"{dia}/{mes}/{anio}"
        
    # 2. Si ya viene como DD/MM/YYYY
    match_latam = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})', val_str)
    if match_latam:
        dia, mes, anio = match_latam.group(1).zfill(2), match_latam.group(2).zfill(2), match_latam.group(3)
        return f"{dia}/{mes}/{anio}"

    # 3. Limpieza de cualquier formato residual
    nums = re.findall(r'\d+', val_str)
    if len(nums) >= 3:
        if len(nums[0]) == 4: # Año al inicio
            return f"{nums.zfill(2)}/{nums.zfill(2)}/{nums[0]}"
        elif len(nums) == 4: # Año al final
            return f"{nums[0].zfill(2)}/{nums.zfill(2)}/{nums}"

    return val_str

def importar_trabajadores_excel(ruta_excel):
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    
    sheet_seleccionada = None
    fila_cabecera = None
    col_map = {}

    for sheet in wb.worksheets:
        for row_idx, row in enumerate(sheet.iter_rows(max_row=20, values_only=True), start=1):
            row_clean = [limpiar_texto(c) for c in row if c is not None]
            texto_unido = " ".join(row_clean)
            
            if "DNI" in texto_unido or ("PATERNO" in texto_unido and "MATERNO" in texto_unido):
                sheet_seleccionada = sheet
                fila_cabecera = row_idx
                
                for idx, cell_val in enumerate(row):
                    h = limpiar_texto(cell_val)
                    if not h: continue
                    
                    if "DNI" in h or "DOCUMENTO" in h: col_map["dni"] = idx
                    elif "PATERNO" in h: col_map["paterno"] = idx
                    elif "MATERNO" in h: col_map["materno"] = idx
                    elif "NOMBRE" in h and "PATERNO" not in h and "MATERNO" not in h: col_map["nombres"] = idx
                    elif "CARGO" in h or "PUESTO" in h: col_map["cargo"] = idx
                    elif "PERSONAL" in h and ("CORREO" in h or "EMAIL" in h): col_map["correo_personal"] = idx
                    elif "INSTITUCIONAL" in h and ("CORREO" in h or "EMAIL" in h): col_map["correo_institucional"] = idx
                    elif "CONTRATO" in h or "REGIMEN" in h: col_map["tipo_contrato"] = idx
                    elif "TELEFONO" in h or "CELULAR" in h or "MOVIL" in h: col_map["telefono"] = idx
                    elif "NACIMIENTO" in h: col_map["fecha_nacimiento"] = idx
                break
        if sheet_seleccionada:
            break

    if not sheet_seleccionada or "dni" not in col_map:
        sheet_seleccionada = wb.active
        fila_cabecera = 1

    conn = get_db()
    cursor = conn.cursor()
    total_cargados = 0

    for row in sheet_seleccionada.iter_rows(min_row=fila_cabecera + 1, values_only=True):
        if not row:
            continue

        dni_idx = col_map.get("dni")
        if dni_idx is None or dni_idx >= len(row):
            continue

        dni_val = row[dni_idx]
        if dni_val is None or str(dni_val).strip() == "":
            continue

        if isinstance(dni_val, float):
            dni = str(int(dni_val)).strip()
        else:
            dni = str(dni_val).strip()
            if "." in dni:
                dni = dni.split(".")[0].strip()

        if not dni or dni == "0" or dni.upper() == "NONE":
            continue

        def get_val(campo):
            idx = col_map.get(campo)
            if idx is not None and idx < len(row) and row[idx] is not None:
                val_str = str(row[idx]).strip()
                return "" if val_str.upper() == "NONE" else val_str
            return ""

        nombres = get_val("nombres")
        paterno = get_val("paterno")
        materno = get_val("materno")
        cargo = get_val("cargo")
        c_pers = get_val("correo_personal")
        c_inst = get_val("correo_institucional")
        contrato = get_val("tipo_contrato")
        telefono = get_val("telefono")
        
        # Extracción y formateo estricto DÍA/MES/AÑO
        raw_f_nac = row[col_map.get("fecha_nacimiento")] if col_map.get("fecha_nacimiento") is not None else None
        f_nac = formatear_solo_fecha(raw_f_nac)

        cursor.execute("""
            INSERT INTO trabajadores (dni, nombres, apellido_paterno, apellido_materno, cargo, 
                                      correo_personal, correo_institucional, tipo_contrato, telefono, fecha_nacimiento)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(dni) DO UPDATE SET
                nombres = excluded.nombres,
                apellido_paterno = excluded.apellido_paterno,
                apellido_materno = excluded.apellido_materno,
                cargo = excluded.cargo,
                correo_personal = excluded.correo_personal,
                correo_institucional = excluded.correo_institucional,
                tipo_contrato = excluded.tipo_contrato,
                telefono = excluded.telefono,
                fecha_nacimiento = excluded.fecha_nacimiento
        """, (dni, nombres, paterno, materno, cargo, c_pers, c_inst, contrato, telefono, f_nac))
        total_cargados += 1

    conn.commit()
    conn.close()
    return total_cargados