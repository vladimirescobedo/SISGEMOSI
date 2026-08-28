import os
from docx import Document

def procesar_parrafo(parrafo, datos_reemplazo):
    if not parrafo.text:
        return

    # Verificar si alguna etiqueta está en el texto del párrafo
    hay_etiquetas = any(clave in parrafo.text for clave in datos_reemplazo)
    if not hay_etiquetas:
        return

    # 1. Intentar reemplazo simple por run
    for clave, valor in datos_reemplazo.items():
        if clave in parrafo.text:
            for run in parrafo.runs:
                if clave in run.text:
                    run.text = run.text.replace(clave, str(valor or ""))

    # 2. Si Word dividió la etiqueta en varios runs, reemplazar sobre el texto completo
    if any(clave in parrafo.text for clave in datos_reemplazo):
        texto_completo = parrafo.text
        for clave, valor in datos_reemplazo.items():
            texto_completo = texto_completo.replace(clave, str(valor or ""))
        
        # Asignar texto reemplazado al primer run y vaciar los demás para no duplicar
        if parrafo.runs:
            parrafo.runs[0].text = texto_completo
            for r in parrafo.runs[1:]:
                r.text = ""
        else:
            parrafo.text = texto_completo

def rellenar_plantilla_docx(ruta_plantilla, datos_reemplazo, ruta_salida):
    doc = Document(ruta_plantilla)

    # 1. Reemplazar en párrafos del cuerpo
    for p in doc.paragraphs:
        procesar_parrafo(p, datos_reemplazo)

    # 2. Reemplazar en todas las tablas y celdas
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    procesar_parrafo(p, datos_reemplazo)

    # 3. Reemplazar en encabezados y pies de página si existen
    for seccion in doc.sections:
        for p in seccion.header.paragraphs:
            procesar_parrafo(p, datos_reemplazo)
        for p in seccion.footer.paragraphs:
            procesar_parrafo(p, datos_reemplazo)

    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    doc.save(ruta_salida)
    return ruta_salida