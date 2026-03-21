from flask import Flask, render_template, request, send_file, after_this_request
import os
from pdf2docx import Converter

app = Flask(__name__)

# En Render, la carpeta /tmp es la mejor para escribir archivos temporales
UPLOAD_FOLDER = '/tmp'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convertir', methods=['POST'])
def convertir():
    archivo = request.files.get('archivo')
    if not archivo: 
        return "No se seleccionó ningún archivo", 400
    
    # Limpiamos el nombre del archivo para evitar errores de rutas
    nombre_archivo = archivo.filename
    ruta_entrada = os.path.join(UPLOAD_FOLDER, nombre_archivo)
    archivo.save(ruta_entrada)
    
    nombre_base, extension = os.path.splitext(nombre_archivo)

    if extension.lower() == ".pdf":
        ruta_salida = os.path.join(UPLOAD_FOLDER, nombre_base + ".docx")
        
        try:
            # Proceso de conversión
            cv = Converter(ruta_entrada)
            cv.convert(ruta_salida)
            cv.close()

            # Función para borrar los archivos después de enviarlos al usuario
            @after_this_request
            def cleanup(response):
                try:
                    if os.path.exists(ruta_entrada): os.remove(ruta_entrada)
                    if os.path.exists(ruta_salida): os.remove(ruta_salida)
                except Exception as e:
                    print(f"Error limpiando archivos: {e}")
                return response

            return send_file(ruta_salida, as_attachment=True)
            
        except Exception as e:
            return f"Hubo un error en la conversión: {str(e)}", 500
    else:
        return "Formato no soportado. Por favor sube un PDF.", 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
