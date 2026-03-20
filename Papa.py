from flask import Flask, render_template, request, send_file
import os
from pdf2docx import Converter

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convertir', methods=['POST'])
def convertir():
    archivo = request.files.get('archivo')
    if not archivo: return "No hay archivo"
    
    ruta_entrada = os.path.join(UPLOAD_FOLDER, archivo.filename)
    archivo.save(ruta_entrada)
    nombre_base, extension = os.path.splitext(archivo.filename)

    if extension.lower() == ".pdf":
        ruta_salida = os.path.join(UPLOAD_FOLDER, nombre_base + ".docx")
        cv = Converter(ruta_entrada)
        cv.convert(ruta_salida)
        cv.close()
        return send_file(ruta_salida, as_attachment=True)
    else:
        return "En la versión web solo funciona PDF a Word. ¡Usa la versión de escritorio para Excel!"

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
