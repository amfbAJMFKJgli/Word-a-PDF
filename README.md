# Conversor Word ↔ PDF

Aplicación web para convertir archivos Word a PDF y viceversa, preservando el formato, tablas e imágenes.

## Características

- **Word → PDF:** Convierte documentos .docx a PDF con soporte para texto, tablas e imágenes
- **PDF → Word:** Convierte archivos PDF a documentos .docx editables
- **Sin dependencias externas:** No requiere Microsoft Word ni LibreOffice
- **Interfaz simple:** Sube tu archivo y descarga el resultado

## Tecnologías

- **Python 3** — Lenguaje principal
- **Flask** — Framework web
- **python-docx** — Lectura de archivos Word
- **pdf2docx** — Conversión PDF a Word
- **ReportLab** — Generación de PDFs

## Instalación local

1. Clona el repositorio:
   ```bash
   git clone https://github.com/tu-usuario/word-to-pdf.git
   cd word-to-pdf
   ```

2. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

3. Ejecuta la aplicación:
   ```bash
   python app.py
   ```

4. Abre tu navegador en **http://localhost:5000**

## Despliegue en Render

1. Sube este repositorio a GitHub
2. Crea una cuenta en [Render](https://render.com)
3. Click en **New** → **Web Service**
4. Conecta tu repositorio
5. Configuración:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
6. Click en **Create Web Service**

## Estructura del proyecto

```
word-to-pdf/
├── app.py              # Servidor Flask
├── templates/
│   └── index.html      # Interfaz web
├── uploads/            # Archivos subidos (temporal)
├── outputs/            # Archivos convertidos (temporal)
├── requirements.txt    # Dependencias
├── Procfile            # Configuración para Render
└── README.md           # Este archivo
```

## Licencia

MIT
