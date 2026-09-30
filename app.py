import os
import uuid
import subprocess
from io import BytesIO
from xml.sax.saxutils import escape
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

app = Flask(__name__)
app.secret_key = os.urandom(24)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'docx', 'pdf'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB máximo


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


NS_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def is_production():
    """Detecta si estamos en Render (producción) o localmente"""
    return os.environ.get('RENDER') is not None


def convert_docx_to_pdf(input_path, output_path):
    """Convierte DOCX a PDF usando reportlab (soporte Unicode, tablas e imágenes)"""
    doc = Document(input_path)
    pdf_doc = SimpleDocTemplate(output_path, pagesize=letter)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'],
                                 fontSize=16, spaceAfter=12)
    normal_style = ParagraphStyle('CustomNormal', parent=styles['Normal'],
                                  fontSize=11, spaceAfter=6)

    max_width = 6 * inch
    para_map = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}
    story = []

    for element in doc.element.body:
        if element.tag.endswith('}p'):
            para = para_map.get(element)
            if not para:
                continue

            text = escape(para.text.strip())
            if text:
                style = title_style if para.style.name.startswith('Heading') else normal_style
                story.append(Paragraph(text, style))
            else:
                story.append(Spacer(1, 6))

            # Imágenes del párrafo, buscadas por su rId
            for blip in element.iter(NS_A + 'blip'):
                rid = blip.get(NS_R + 'embed')
                if not rid or rid not in doc.part.related_parts:
                    continue
                try:
                    blob = doc.part.related_parts[rid].blob
                    w, h = ImageReader(BytesIO(blob)).getSize()
                    scale = min(1, max_width / w)
                    story.append(Spacer(1, 6))
                    story.append(Image(BytesIO(blob), width=w * scale, height=h * scale))
                    story.append(Spacer(1, 6))
                except Exception:
                    pass  # formato no soportado (emf/wmf, etc.)

        elif element.tag.endswith('}tbl'):
            table = table_map.get(element)
            if not table:
                continue
            table_data = [[escape(c.text.strip()) for c in row.cells] for row in table.rows]
            if table_data:
                pdf_table = Table(table_data)
                pdf_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#667eea')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 11),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('FONTSIZE', (0, 1), (-1, -1), 10),
                ]))
                story.append(Spacer(1, 12))
                story.append(pdf_table)
                story.append(Spacer(1, 12))

    pdf_doc.build(story)


def convert_pdf_to_word_libreoffice(input_path, output_path):
    """Convierte PDF a Word usando LibreOffice (mejor calidad)"""
    result = subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'docx',
         '--outdir', os.path.dirname(output_path), input_path],
        capture_output=True, text=True, timeout=60
    )
    if result.returncode != 0:
        raise Exception(f"LibreOffice error: {result.stderr}")
    # LibreOffice genera el archivo con el mismo nombre pero .docx
    generated_docx = input_path.rsplit('.', 1)[0] + '.docx'
    if os.path.exists(generated_docx) and generated_docx != output_path:
        os.rename(generated_docx, output_path)


def convert_pdf_to_word_pdf2docx(input_path, output_path):
    """Convierte PDF a Word usando pdf2docx (local, sin dependencias)"""
    from pdf2docx import Converter
    cv = Converter(input_path)
    cv.convert(output_path)
    cv.close()


def convert_pdf_to_word(input_path, output_path):
    """Convierte PDF a Word usando el mejor método disponible"""
    if is_production():
        # En Render: usar LibreOffice (mejor calidad)
        convert_pdf_to_word_libreoffice(input_path, output_path)
    else:
        # Local: usar pdf2docx (no requiere instalación)
        convert_pdf_to_word_pdf2docx(input_path, output_path)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        conversion_type = request.form.get('conversion_type', 'word_to_pdf')

        if 'file' not in request.files:
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            unique_id = str(uuid.uuid4())[:8]
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{filename}")

            file.save(input_path)

            try:
                if conversion_type == 'word_to_pdf':
                    output_filename = f"{unique_id}_{filename.rsplit('.', 1)[0]}.pdf"
                    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                    convert_docx_to_pdf(input_path, output_path)
                else:  # pdf_to_word
                    output_filename = f"{unique_id}_{filename.rsplit('.', 1)[0]}.docx"
                    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                    convert_pdf_to_word(input_path, output_path)

                flash('¡Conversión exitosa!', 'success')
                return send_file(output_path, as_attachment=True, download_name=output_filename)
            except Exception as e:
                flash(f'Error al convertir: {str(e)}', 'error')
                return redirect(request.url)
            finally:
                if os.path.exists(input_path):
                    os.remove(input_path)

        else:
            flash('Formato no permitido. Solo archivos .docx o .pdf', 'error')
            return redirect(request.url)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True, port=5000)
