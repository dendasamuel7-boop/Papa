from flask import Flask, render_template, request, send_file, after_this_request
import os
from pdf2docx import Converter
from docx2pdf import convert as docx_to_pdf # Nueva librería

app = Flask(__name__)

UPLOAD_FOLDER = '/tmp'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/converter', methods=['POST'])
def convertir():
    archivo = request.files.get('archivo')
    tipo = request.form.get('tipo_conversion') # Lee el selector de tu diseño
    
    if not archivo: 
        return "No hay archivo", 400
    
    ruta_entrada = os.path.join(UPLOAD_FOLDER, archivo.filename)
    archivo.save(ruta_entrada)
    nombre_base, _ = os.path.splitext(archivo.filename)

    try:
        # OPCIÓN A: PDF a WORD
        if tipo == "pdf_to_word":
            ruta_salida = os.path.join(UPLOAD_FOLDER, nombre_base + ".docx")
            cv = Converter(ruta_entrada)
            cv.convert(ruta_salida)
            cv.close()

        # OPCIÓN B: WORD a PDF
        elif tipo == "word_to_pdf":
            ruta_salida = os.path.join(UPLOAD_FOLDER, nombre_base + ".pdf")
            docx_to_pdf(ruta_entrada, ruta_salida)

        else:
            return "Tipo de conversión no válido", 400

        # Limpieza automática después de descargar
        @after_this_request
        def cleanup(response):
            try:
                if os.path.exists(ruta_entrada): os.remove(ruta_entrada)
                if os.path.exists(ruta_salida): os.remove(ruta_salida)
            except: pass
            return response

        return send_file(ruta_salida, as_attachment=True)

    except Exception as e:
        return f"Error en el proceso: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
