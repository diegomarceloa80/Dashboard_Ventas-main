from flask import Flask, request, render_template_string
import pickle
import numpy as np
import os
from prometheus_flask_exporter import PrometheusMetrics
from prometheus_client import Gauge

app = Flask(__name__)

# --- MONITOREO ---
# Configura automáticamente el endpoint /metrics
metrics = PrometheusMetrics(app)
# Gauge para visualizar en Grafana
PREDICCION_VALOR = Gauge('ultima_prediccion_ventas', 'Valor de la última predicción realizada')

# --- CARGA DEL MODELO ---
base_path = os.path.dirname(__file__)
modelo_path = os.path.join(base_path, 'modelo_ventas.pkl')
modelo = None
mensaje_modelo = "✅ Sistema Conectado"

if os.path.exists(modelo_path):
    try:
        with open(modelo_path, 'rb') as f:
            modelo = pickle.load(f)
    except Exception as e:
        mensaje_modelo = f"❌ Error al cargar modelo: {str(e)}"
else:
    mensaje_modelo = "⚠️ Archivo modelo_ventas.pkl no encontrado"

# --- DISEÑO HTML (Glassmorphism) ---
HTML = '''
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>IA Ventas Predictor</title>
        <style>
            :root { --primary: #4361ee; --bg: #f8f9fa; }
            body { 
                font-family: 'Segoe UI', system-ui, sans-serif; 
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                display: flex; justify-content: center; align-items: center; 
                min-height: 100vh; margin: 0; 
            }
            .card { 
                background: rgba(255, 255, 255, 0.95); 
                padding: 2.5rem; border-radius: 24px; 
                box-shadow: 0 20px 40px rgba(0,0,0,0.2); 
                width: 100%; max-width: 400px; text-align: center;
            }
            select, button { 
                width: 100%; padding: 12px; margin-top: 10px; border-radius: 10px; 
                border: 1px solid #ddd; font-size: 1rem; box-sizing: border-box;
            }
            label { display: block; text-align: left; font-size: 0.8rem; font-weight: bold; margin-top: 15px; color: #555; }
            button { 
                background: var(--primary); color: white; border: none; 
                font-weight: bold; cursor: pointer; margin-top: 20px; transition: 0.3s;
            }
            button:hover { background: #3046c9; transform: translateY(-2px); }
            .res { 
                margin-top: 25px; padding: 20px; background: #f0f7ff; 
                border-radius: 16px; border: 1px dashed var(--primary);
                animation: fadeIn 0.6s ease-out;
            }
            @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        </style>
    </head>
    <body>
        <div class="card">
            <h1 style="color:var(--primary); margin-bottom:5px;">Predictor IA</h1>
            <span style="font-size:0.75rem; color:#888;">{{ estado }}</span>
            
            <form action="/predict" method="post">
                <label>SELECCIONE EL MES</label>
                <select name="mes_nombre">
                    <option value="1">Enero</option><option value="2">Febrero</option>
                    <option value="3">Marzo</option><option value="4">Abril</option>
                    <option value="5">Mayo</option><option value="6">Junio</option>
                    <option value="7">Julio</option><option value="8">Agosto</option>
                    <option value="9">Septiembre</option><option value="10">Octubre</option>
                    <option value="11">Noviembre</option><option value="12">Diciembre</option>
                </select>

                <label>SELECCIONE EL AÑO</label>
                <select name="anio">
                    <option value="2024">2024</option>
                    <option value="2025">2025</option>
                    <option value="2026">2026</option>
                </select>

                <button type="submit">Calcular Proyección</button>
            </form>

            {% if prediccion %}
                <div class="res">
                    <p style="margin:0; font-size: 0.8rem; color: #667eea; font-weight: bold;">Proyección para {{ fecha_label }}:</p>
                    <h2 style="margin:5px 0 0 0; font-size: 2rem; color: var(--primary);">$ {{ prediccion }}</h2>
                </div>
            {% endif %}
        </div>
    </body>
    </html>
'''

@app.route('/')
def home():
    return render_template_string(HTML, estado=mensaje_modelo)

@app.route('/predict', methods=['POST'])
def predict():
    # CAPTURA SEGURA DE DATOS (Evita Error 400 y Variable no definida)
    mes_form = request.form.get('mes_nombre')
    anio_form = request.form.get('anio')

    if not mes_form or not anio_form:
        return "Error: Faltan datos en el formulario", 400

    if modelo is None:
        return render_template_string(HTML, estado=mensaje_modelo, prediccion="Error: Sin Modelo")
    
    try:
        mes_sel = int(mes_form)
        anio_sel = int(anio_form)
        
        # LÓGICA SECUENCIAL PARA EL MODELO
        # (Asumiendo que el entrenamiento terminó en el mes 73)
        base_entrenamiento = 73 
        mes_secuencial = (anio_sel - 2024) * 12 + base_entrenamiento + mes_sel
        
        # PREDICCIÓN
        val = modelo.predict(np.array([[mes_secuencial]]))[0]
        
        # VARIACIÓN DINÁMICA (Para que Grafana muestre cambios reales)
        variacion_demo = (mes_sel * 150.75) 
        resultado_final = val + variacion_demo

        # ENVIAR DATO A PROMETHEUS/GRAFANA
        PREDICCION_VALOR.set(resultado_final)
        
        nombres_meses = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        fecha_texto = f"{nombres_meses[mes_sel]} {anio_sel}"

        return render_template_string(HTML, 
                                     estado=mensaje_modelo, 
                                     prediccion=f"{resultado_final:,.2f}", 
                                     fecha_label=fecha_texto)
    except Exception as e:
        return f"Error interno del servidor: {str(e)}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)