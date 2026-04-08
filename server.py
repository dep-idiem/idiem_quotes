"""
IDIEM — Servidor local para generación de propuestas Word
=========================================================
Cómo usar:
1. Instala dependencias:  pip install flask python-docx
2. Corre el servidor:     python server.py
3. Abre el HTML en Chrome y genera propuestas normalmente

El servidor corre en http://localhost:5050
No cierres esta ventana mientras usas el generador.
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os
import re
from datetime import datetime

app = Flask(__name__)
CORS(app)  # Permite que el HTML llame al servidor sin problemas de CORS

# ============================================================
# CONFIGURACIÓN 
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = {
    'diagnostico': os.path.join(BASE_DIR, 'plantillas', 'template_diagnostico_IDIEM.docx'),
    # 'evaluacion':  os.path.join(BASE_DIR, 'plantillas', 'template_evaluacion_IDIEM.docx'),
    # 'calculo':     os.path.join(BASE_DIR, 'plantillas', 'template_calculo_IDIEM.docx'),
}

# Marcadores en el template Word — deben coincidir exactamente con el texto en rojo
MARKERS = {
    'alcance':             '▶ PEGAR AQUÍ: SECCIÓN 1 — ALCANCE (texto generado por el generador HTML)',
    'propuesta_tecnica':   '▶ PEGAR AQUÍ: SECCIÓN 2 — PROPUESTA TÉCNICA (texto generado por el generador HTML)',
    'plazos':              '▶ PEGAR AQUÍ: SECCIÓN 3 — PLAZOS (texto generado por el generador HTML)',
    'propuesta_economica': '▶ PEGAR AQUÍ: SECCIÓN 4 — PROPUESTA ECONÓMICA (texto generado por el generador HTML)',
    'recursos':            '▶ PEGAR AQUÍ: SECCIÓN 5 — RECURSOS DEL MANDANTE (texto generado por el generador HTML)',
    'exclusiones':         '▶ PEGAR AQUÍ: SECCIÓN 6 — EXCLUSIONES (texto generado por el generador HTML)',
}

# ============================================================
# FUNCIONES DE UTILIDAD
# ============================================================

def parse_sections(texto_completo):
    """
    Recibe el texto completo generado por Ollama y lo divide en secciones.
    Busca los títulos de sección para separar el contenido.
    """
    secciones = {
        'alcance': '',
        'propuesta_tecnica': '',
        'plazos': '',
        'propuesta_economica': '',
        'recursos': '',
        'exclusiones': '',
    }

    # Patrones para detectar inicio de cada sección
    patrones = {
        'alcance':             r'(?:^|\n)\s*(?:\d+[\.\-]\s*)?(?:1[\.\-]\s*)?ALCANCE',
        'propuesta_tecnica':   r'(?:^|\n)\s*(?:\d+[\.\-]\s*)?(?:2[\.\-]\s*)?PROPUESTA\s+T[ÉE]CNICA',
        'plazos':              r'(?:^|\n)\s*(?:\d+[\.\-]\s*)?(?:3[\.\-]\s*)?PLAZOS?',
        'propuesta_economica': r'(?:^|\n)\s*(?:\d+[\.\-]\s*)?(?:4[\.\-]\s*)?PROPUESTA\s+ECON[ÓO]MICA',
        'recursos':            r'(?:^|\n)\s*(?:\d+[\.\-]\s*)?(?:5[\.\-]\s*)?RECURSOS',
        'exclusiones':         r'(?:^|\n)\s*(?:\d+[\.\-]\s*)?(?:6[\.\-]\s*)?EXCLUSIONES',
    }

    orden = ['alcance', 'propuesta_tecnica', 'plazos', 'propuesta_economica', 'recursos', 'exclusiones']
    posiciones = {}

    for key, patron in patrones.items():
        match = re.search(patron, texto_completo, re.IGNORECASE | re.MULTILINE)
        if match:
            posiciones[key] = match.start()

    # Ordenar por posición en el texto
    orden_encontrado = sorted(posiciones.keys(), key=lambda k: posiciones[k])

    # Extraer contenido entre secciones
    for i, key in enumerate(orden_encontrado):
        inicio = posiciones[key]
        if i + 1 < len(orden_encontrado):
            siguiente = orden_encontrado[i + 1]
            fin = posiciones[siguiente]
            secciones[key] = texto_completo[inicio:fin].strip()
        else:
            secciones[key] = texto_completo[inicio:].strip()

    # Si no se encontraron secciones bien separadas, poner todo en alcance
    if not any(secciones.values()):
        secciones['alcance'] = texto_completo

    return secciones


def reemplazar_marcador(doc, marcador, contenido):
    """
    Busca el marcador en el documento Word y lo reemplaza con el contenido.
    Mantiene el estilo del párrafo original.
    """
    for i, para in enumerate(doc.paragraphs):
        if marcador in para.text:
            # Guardar el estilo del párrafo
            style_name = para.style.name

            # Limpiar el párrafo marcador
            for run in para.runs:
                run.text = ''

            # Dividir el contenido en líneas
            lineas = contenido.split('\n')

            # Primera línea va en el párrafo existente
            if lineas:
                primera_linea = lineas[0].strip()
                if primera_linea:
                    run = para.add_run(primera_linea)
                    run.font.color.rgb = RGBColor(0, 0, 0)  # Negro
                    run.font.size = Pt(9)

            # Líneas siguientes: insertar nuevos párrafos después
            # Necesitamos insertar en orden inverso para mantener posición
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement
            import copy

            para_element = para._element
            parent = para_element.getparent()
            idx = list(parent).index(para_element)

            for j, linea in enumerate(lineas[1:], 1):
                linea = linea.strip()

                # Crear nuevo párrafo
                new_para = OxmlElement('w:p')

                # Copiar propiedades de párrafo del original
                if para_element.find(qn('w:pPr')) is not None:
                    pPr = copy.deepcopy(para_element.find(qn('w:pPr')))
                    new_para.append(pPr)

                # Agregar run con texto
                if linea:
                    new_r = OxmlElement('w:r')
                    new_rPr = OxmlElement('w:rPr')

                    # Color negro
                    new_color = OxmlElement('w:color')
                    new_color.set(qn('w:val'), '000000')
                    new_rPr.append(new_color)

                    # Tamaño fuente
                    new_sz = OxmlElement('w:sz')
                    new_sz.set(qn('w:val'), '18')  # 9pt = 18 half-points
                    new_rPr.append(new_sz)

                    new_r.append(new_rPr)
                    new_t = OxmlElement('w:t')
                    new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    new_t.text = linea
                    new_r.append(new_t)
                    new_para.append(new_r)

                parent.insert(idx + j, new_para)

            return True

    return False


def actualizar_portada(doc, datos):
    """
    Actualiza los campos variables de la portada:
    código, revisión, fecha, elaborado por, revisado por, destinatario.
    """
    campos = {
        'PR.DEP.2025.0215':        datos.get('codigo', 'PR.DEP.2025.XXXX'),
        'Revisión Nº2':            datos.get('revision', 'Revisión Nº1'),
        '10-10-2025':              datos.get('fecha', datetime.now().strftime('%d-%m-%Y')),
        'Pilar Castellón F.':      datos.get('elaborado', ''),
        'Pablo Herrera T.':        '',
        'Guillermo Sierra R.':     datos.get('revisado', 'Guillermo Sierra R.'),
        'Servicio de Salud Metropolitano': datos.get('cliente', ''),
        'Oriente':                 '',
    }

    for para in doc.paragraphs:
        for key, val in campos.items():
            if key in para.text and val:
                for run in para.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, val)

    # También buscar en tablas (la portada tiene tablas)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, val in campos.items():
                        if key in para.text and val:
                            for run in para.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, val)


# ============================================================
# RUTAS DEL SERVIDOR
# ============================================================

@app.route('/health', methods=['GET'])
def health():
    """Verificar que el servidor está corriendo"""
    return jsonify({'status': 'ok', 'message': 'Servidor IDIEM corriendo'})


@app.route('/generar-word', methods=['POST'])
def generar_word():
    """
    Recibe el texto generado por Ollama y los datos del formulario,
    rellena el template Word y lo devuelve para descarga.
    """
    try:
        data = request.get_json()

        tipo_servicio = data.get('tipo', 'diagnostico')
        texto_completo = data.get('texto', '')
        datos_formulario = data.get('formulario', {})

        # Verificar que existe el template
        template_path = TEMPLATES.get(tipo_servicio)
        if not template_path or not os.path.exists(template_path):
            return jsonify({
                'error': f'Template no encontrado: {template_path}\n'
                         f'Verifica que el archivo esté en la carpeta plantillas/'
            }), 404

        # Cargar el template
        doc = Document(template_path)

        # Actualizar portada con datos del formulario
        actualizar_portada(doc, datos_formulario)

        # Parsear el texto en secciones
        secciones = parse_sections(texto_completo)

        # Reemplazar cada marcador con su contenido
        for key, marcador in MARKERS.items():
            contenido = secciones.get(key, '')
            if contenido:
                encontrado = reemplazar_marcador(doc, marcador, contenido)
                if not encontrado:
                    print(f'⚠ Marcador no encontrado: {key}')
            else:
                print(f'⚠ Sin contenido para sección: {key}')

        # Generar nombre del archivo
        edificio = datos_formulario.get('edificio', 'propuesta').replace(' ', '_')
        codigo = datos_formulario.get('codigo', 'PR_DEP_2025_XXXX').replace('.', '_')
        nombre_archivo = f'{codigo}_{edificio}.docx'

        # Guardar en memoria y enviar
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        import traceback
        return jsonify({
            'error': str(e),
            'detalle': traceback.format_exc()
        }), 500


@app.route('/templates', methods=['GET'])
def listar_templates():
    """Lista los templates disponibles"""
    disponibles = {}
    for key, path in TEMPLATES.items():
        disponibles[key] = {
            'existe': os.path.exists(path),
            'ruta': path
        }
    return jsonify(disponibles)


# ============================================================
# INICIO DEL SERVIDOR
# ============================================================
if __name__ == '__main__':
    print("\n" + "="*50)
    print("  IDIEM — Servidor de propuestas")
    print("="*50)
    print(f"  URL: http://localhost:5050")
    print(f"  Templates:")
    for key, path in TEMPLATES.items():
        existe = "✓" if os.path.exists(path) else "✗ NO ENCONTRADO"
        print(f"    {existe}  {key}: {path}")
    print("="*50)
    print("  Deja esta ventana abierta mientras usas el generador.")
    print("  Para detener: Ctrl + C")
    print("="*50 + "\n")

    app.run(host='localhost', port=5050, debug=False)