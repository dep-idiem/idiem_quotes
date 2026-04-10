"""
IDIEM — Servidor local para generación de propuestas Word
=========================================================
Cómo usar:
1. Instala dependencias:  pip install flask flask-cors python-docx
2. Corre el servidor:     python server.py
3. Abre el HTML en Chrome y genera propuestas normalmente

El servidor corre en http://localhost:5050
No cierres esta ventana mientras usas el generador.
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from docx import Document
from docx.shared import Pt, RGBColor
import io
import os
import json
import re
from datetime import datetime

app = Flask(__name__)
CORS(app)

# ============================================================
# CONFIGURACIÓN
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES = {
    'diagnostico': os.path.join(BASE_DIR, 'plantillas', 'template_diagnostico_IDIEM.docx'),
    # 'evaluacion': os.path.join(BASE_DIR, 'plantillas', 'template_evaluacion_IDIEM.docx'),
}

PROMPTS = {
    'diagnostico': os.path.join(BASE_DIR, 'prompts', 'diagnostico.txt'),
    # 'evaluacion': os.path.join(BASE_DIR, 'prompts', 'evaluacion.txt'),
}

# Marcadores en rojo del template Word
MARKERS = {
    'alcance':             '▶ PEGAR AQUÍ: SECCIÓN 1 — ALCANCE (texto generado por el generador HTML)',
    'propuesta_tecnica':   '▶ PEGAR AQUÍ: SECCIÓN 2 — PROPUESTA TÉCNICA (texto generado por el generador HTML)',
    'plazos':              '▶ PEGAR AQUÍ: SECCIÓN 3 — PLAZOS (texto generado por el generador HTML)',
    'propuesta_economica': '▶ PEGAR AQUÍ: SECCIÓN 4 — PROPUESTA ECONÓMICA (texto generado por el generador HTML)',
    'recursos':            '▶ PEGAR AQUÍ: SECCIÓN 5 — RECURSOS DEL MANDANTE (texto generado por el generador HTML)',
    'exclusiones':         '▶ PEGAR AQUÍ: SECCIÓN 6 — EXCLUSIONES (texto generado por el generador HTML)',
}

# ============================================================
# CARGAR Y RELLENAR PROMPT
# ============================================================

def cargar_prompt(tipo, datos_formulario):
    prompt_path = PROMPTS.get(tipo)
    if not prompt_path or not os.path.exists(prompt_path):
        raise FileNotFoundError(f'Prompt no encontrado: {prompt_path}')

    with open(prompt_path, 'r', encoding='utf-8') as f:
        prompt = f.read()

    forma_map = {
        '50_inicio':         '50% al momento de aceptar el servicio, 50% al entregar el informe',
        '50_terreno':        '50% al término de los trabajos en terreno, 50% al entregar el informe de diagnóstico',
        'estados_mensuales': 'Estados de pago mensuales',
        'otro':              'Según detalle'
    }
    forma_pago = forma_map.get(datos_formulario.get('forma_pago', '50_inicio'), '')
    if datos_formulario.get('detalle_pago'):
        forma_pago += '. ' + datos_formulario['detalle_pago']

    actividades = datos_formulario.get('actividades', [])
    actividades_str = '\n'.join(f'- {a}' for a in actividades) if actividades else '- No especificadas'

    reemplazos = {
        '{{cliente}}':       datos_formulario.get('cliente', 'No indicado'),
        '{{contacto}}':      datos_formulario.get('contacto', 'No indicado'),
        '{{edificio}}':      datos_formulario.get('edificio', 'No indicado'),
        '{{direccion}}':     datos_formulario.get('direccion', 'No indicada'),
        '{{superficie}}':    datos_formulario.get('superficie', 'No indicada'),
        '{{anio}}':          datos_formulario.get('anio', 'No indicado'),
        '{{pisos}}':         datos_formulario.get('pisos', 'No indicado'),
        '{{materialidad}}':  datos_formulario.get('materialidad', 'No especificada'),
        '{{condicion}}':     datos_formulario.get('condicion', 'No especificada'),
        '{{motivacion}}':    datos_formulario.get('motivacion', 'No indicada'),
        '{{obs_tecnicas}}':  datos_formulario.get('obs_tecnicas', 'Ninguna'),
        '{{actividades}}':   actividades_str,
        '{{plazo_terreno}}': datos_formulario.get('plazo_terreno', 'A definir'),
        '{{plazo_gabinete}}':datos_formulario.get('plazo_gabinete', 'A definir'),
        '{{plazo_total}}':   datos_formulario.get('plazo_total', 'A definir'),
        '{{precio_uf}}':     datos_formulario.get('precio_uf', 'A definir'),
        '{{forma_pago}}':    forma_pago,
        '{{exclusiones}}':   datos_formulario.get('exclusiones', 'Ninguna adicional'),
        '{{codigo}}':        datos_formulario.get('codigo', 'PR.DEP.2025.XXXX'),
        '{{revision}}':      datos_formulario.get('revision', 'Rev. N°1'),
        '{{fecha}}':         datos_formulario.get('fecha', datetime.now().strftime('%d-%m-%Y')),
        '{{elaborado}}':     datos_formulario.get('elaborado', 'Pendiente'),
        '{{revisado}}':      datos_formulario.get('revisado', 'Guillermo Sierra R.'),
    }

    for placeholder, valor in reemplazos.items():
        prompt = prompt.replace(placeholder, str(valor))

    return prompt


# ============================================================
# PARSEAR JSON DE OLLAMA (robusto)
# ============================================================

def parsear_json_ollama(texto_crudo):
    texto = re.sub(r'```json\s*', '', texto_crudo)
    texto = re.sub(r'```\s*', '', texto)
    texto = texto.strip()

    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        pass

    inicio = texto.find('{')
    fin = texto.rfind('}')
    if inicio != -1 and fin != -1 and fin > inicio:
        try:
            return json.loads(texto[inicio:fin+1])
        except json.JSONDecodeError:
            pass

    print(f'⚠ No se pudo parsear JSON. Primeros 500 chars:\n{texto[:500]}')
    return {
        'alcance': texto,
        'propuesta_tecnica': '',
        'plazos': '',
        'propuesta_economica': '',
        'recursos': '',
        'exclusiones': ''
    }


# ============================================================
# RELLENAR WORD
# ============================================================

def reemplazar_marcador(doc, marcador, contenido):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import copy

    for para in doc.paragraphs:
        if marcador in para.text:
            para_element = para._element
            parent = para_element.getparent()
            idx = list(parent).index(para_element)

            for run in para.runs:
                run.text = ''

            lineas = contenido.split('\n')

            if lineas:
                primera = lineas[0].strip()
                if primera:
                    run = para.add_run(primera)
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    run.font.size = Pt(9)

            for j, linea in enumerate(lineas[1:], 1):
                new_para = OxmlElement('w:p')

                pPr_orig = para_element.find(qn('w:pPr'))
                if pPr_orig is not None:
                    pPr_new = copy.deepcopy(pPr_orig)
                    rPr = pPr_new.find(qn('w:rPr'))
                    if rPr is not None:
                        color = rPr.find(qn('w:color'))
                        if color is not None:
                            rPr.remove(color)
                    new_para.append(pPr_new)

                linea_limpia = linea.strip()
                if linea_limpia:
                    new_r = OxmlElement('w:r')
                    new_rPr = OxmlElement('w:rPr')

                    new_color = OxmlElement('w:color')
                    new_color.set(qn('w:val'), '000000')
                    new_rPr.append(new_color)

                    new_sz = OxmlElement('w:sz')
                    new_sz.set(qn('w:val'), '18')
                    new_rPr.append(new_sz)

                    new_r.append(new_rPr)
                    new_t = OxmlElement('w:t')
                    new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    new_t.text = linea_limpia
                    new_r.append(new_t)
                    new_para.append(new_r)

                parent.insert(idx + j, new_para)

            return True

    return False


def actualizar_portada(doc, datos):
    reemplazos = {
        'PR.DEP.2025.0215':   datos.get('codigo', 'PR.DEP.2025.XXXX'),
        'Revisión Nº2':       datos.get('revision', 'Rev. N°1'),
        '10-10-2025':         datos.get('fecha', datetime.now().strftime('%d-%m-%Y')),
        'Pilar Castellón F.': datos.get('elaborado', ''),
        'Pablo Herrera T.':   '',
        'Guillermo Sierra R.':datos.get('revisado', 'Guillermo Sierra R.'),
    }

    def reemplazar_en_parrafos(parrafos):
        for para in parrafos:
            for viejo, nuevo in reemplazos.items():
                if viejo in para.text and nuevo is not None:
                    for run in para.runs:
                        if viejo in run.text:
                            run.text = run.text.replace(viejo, nuevo)

    reemplazar_en_parrafos(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                reemplazar_en_parrafos(cell.paragraphs)


# ============================================================
# RUTAS
# ============================================================

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


@app.route('/construir-prompt', methods=['POST'])
def construir_prompt():
    """
    El HTML llama a este endpoint para obtener el prompt completo
    (con ejemplos few-shot incluidos) listo para enviar a Ollama.
    """
    try:
        data       = request.get_json()
        tipo       = data.get('tipo', 'diagnostico')
        datos_form = data.get('formulario', {})
        prompt     = cargar_prompt(tipo, datos_form)
        return jsonify({'prompt': prompt})
    except FileNotFoundError as e:
        return jsonify({'error': str(e)}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/generar-word', methods=['POST'])
def generar_word():
    """
    Recibe el texto JSON generado por Ollama,
    rellena el template Word y lo devuelve para descarga.
    """
    try:
        data         = request.get_json()
        tipo         = data.get('tipo', 'diagnostico')
        texto_ollama = data.get('texto', '')
        datos_form   = data.get('formulario', {})

        template_path = TEMPLATES.get(tipo)
        if not template_path or not os.path.exists(template_path):
            return jsonify({'error': f'Template no encontrado: {template_path}'}), 404

        secciones = parsear_json_ollama(texto_ollama)
        print(f'\nSecciones recibidas:')
        for k, v in secciones.items():
            print(f'  {k}: {len(v)} chars')

        doc = Document(template_path)
        actualizar_portada(doc, datos_form)

        for key, marcador in MARKERS.items():
            contenido = secciones.get(key, '')
            if contenido:
                ok = reemplazar_marcador(doc, marcador, contenido)
                if not ok:
                    print(f'⚠ Marcador no encontrado: {key}')
            else:
                print(f'⚠ Sección vacía: {key}')

        edificio = datos_form.get('edificio', 'propuesta').replace(' ', '_')
        codigo   = datos_form.get('codigo', 'PR_DEP_2025_XXXX').replace('.', '_')
        nombre   = f'{codigo}_{edificio}.docx'

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=nombre,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({'error': str(e), 'detalle': traceback.format_exc()}), 500


# ============================================================
# INICIO
# ============================================================
if __name__ == '__main__':
    print('\n' + '='*55)
    print('  IDIEM — Servidor de propuestas')
    print('='*55)
    print('  URL: http://localhost:5050')
    print('\n  Templates:')
    for key, path in TEMPLATES.items():
        print(f'    {"✓" if os.path.exists(path) else "✗ NO ENCONTRADO"}  {key}')
    print('\n  Prompts:')
    for key, path in PROMPTS.items():
        print(f'    {"✓" if os.path.exists(path) else "✗ NO ENCONTRADO"}  {key}')
    print('='*55)
    print('  Deja esta ventana abierta mientras usas el generador.')
    print('  Ctrl + C para detener.')
    print('='*55 + '\n')
    app.run(host='localhost', port=5050, debug=False)