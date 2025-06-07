from flask import Flask, request, redirect, url_for, send_file, render_template
import os
from werkzeug.utils import secure_filename
from script import procesar_archivo

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

# Asegúrate de que la carpeta de carga exista
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '' or not allowed_file(file.filename):
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Obtener la placa del formulario
        placa_vehiculo = request.form['placa']

        # Obtener los valores del checkbox (last_route) y la lista desplegable (route_choice)
        last_route = 'last_route' in request.form  # Verifica si el checkbox está marcado
        route_choice = request.form['route_choice']  # El valor seleccionado de la lista desplegable
        
        # Ejecutar la función de procesamiento del archivo con los nuevos parámetros
        output_file = procesar_archivo(file_path, request.form['output_name'], placa_vehiculo, last_route, route_choice)
        
        return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

