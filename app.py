from flask import Flask, request, render_template_string
import pickle
import numpy as np
import os

app = Flask(__name__)

# 1. Cargar el "cerebro" (Corregido)
# Usamos un try/except por si el archivo no existe aún
try:
    with open('modelo_ventas.pkl', 'rb') as f:
        modelo = pickle.load(f)
except FileNotFoundError:
    modelo = None
    print("Aviso: No se encontró modelo_ventas.pkl. Asegúrate de que train.py corrió primero.")

# Diseño visual simple
HTML = '''
    <!DOCTYPE html>
    <html>
    <head><title>Predicción de Ventas</title></head>
    <body>
        <h1>Predicción de Ventas - Sistema de IA</h1>
        <p>Introduce el número correlativo del mes para obtener la proyección de ventas totales.</p>
        <form action="/predict" method="post">
            Mes (ejemplo: 55): 
            <input type="number" name="mes" required>
            <input type="submit" value="Calcular Predicción">
        </form>
        {% if prediccion %}
            <div style="margin-top: 20px; padding: 10px; background-color: #e1f5fe;">
                <h2>Resultado de la IA:</h2>
                <p>Para el mes seleccionado, se estiman ventas de: <strong>${{ prediccion }}</strong></p>
            </div>
        {% endif %}
    </body>
    </html>
'''

@app.route('/')
def home():
    return render_template_string(HTML)

# AQUÍ ESTÁ LA CORRECCIÓN: 'methods' en plural
@app.route('/predict', methods=['POST'])
def predict():
    if modelo is None:
        return "Error: El modelo no está cargado. Revisa el entrenamiento."
    
    # Obtener el dato del formulario
    mes = float(request.form['mes'])
    
    # Hacer la predicción
    prediccion = modelo.predict(np.array([[mes]]))[0]
    
    return render_template_string(HTML, prediccion=f"{prediccion:,.2f}")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)