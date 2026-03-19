from flask import Flask, render_template, request, send_file
import os
from pdf2docx import Converter
from docx2pdf import convert
import comtypes.client
import pythoncom

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convertir', methods=['POST'])
def convertir():
    pythoncom.CoInitialize()
    archivo = request.files.get('archivo')
    if not archivo or archivo.filename == '':
        return "No seleccionaste nada"

    ruta_entrada = os.path.join(UPLOAD_FOLDER, archivo.filename)
    archivo.save(ruta_entrada)
    nombre_base, extension = os.path.splitext(archivo.filename)
    extension = extension.lower()

    # --- 1. PDF A WORD ---
    if extension == ".pdf":
        ruta_salida = os.path.join(UPLOAD_FOLDER, nombre_base + ".docx")
        cv = Converter(ruta_entrada)
        cv.convert(ruta_salida)
        cv.close()
        return send_file(ruta_salida, as_attachment=True)

    # --- 2. WORD A PDF ---
    elif extension == ".docx":
        try:
            convert(os.path.abspath(ruta_entrada))
            ruta_salida = os.path.join(UPLOAD_FOLDER, nombre_base + ".pdf")
            return send_file(ruta_salida, as_attachment=True)
        except Exception as e:
            return f"Fallo en Word: {e}"

     # EXCEL
    elif extension == ".xlsx":
        try:
            excel = comtypes.client.CreateObject("Excel.Application")
            excel.Visible = False
            libro = excel.Workbooks.Open(os.path.abspath(ruta_entrada))
            ruta_salida = os.path.abspath(os.path.join(UPLOAD_FOLDER, nombre_base + ".pdf"))
            libro.ExportAsFixedFormat(0, ruta_salida)
            libro.Close(False)
            excel.Quit()
            return send_file(ruta_salida, as_attachment=True)
        except Exception as e:
            return f"Error en Excel: {e}"

    return "Formato no soportado"

if __name__ == '__main__':
    app.run(debug=True)