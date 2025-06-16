from flask import Flask, render_template, request, send_from_directory
import os
from procesar import procesar_archivos
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        plantilla_file = request.files['plantilla']
        datos_file = request.files['datos']

        plantilla_path = os.path.join(UPLOAD_FOLDER, "Plantilla Ajustada Inicial.xlsx")
        datos_path = os.path.join(UPLOAD_FOLDER, "Datos.xlsx")

        plantilla_file.save(plantilla_path)
        datos_file.save(datos_path)

        output_filename = procesar_archivos(plantilla_path, datos_path, OUTPUT_FOLDER)
        return send_from_directory(OUTPUT_FOLDER, output_filename, as_attachment=True)

    return render_template('index.html', current_year=datetime.now().year)

if __name__ == '__main__':
    app.run(debug=True)
