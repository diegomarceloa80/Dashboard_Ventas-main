from flask import Flask, request, render_template_string
import pickle
import numpy as np
import os
from prometheus_flask_exporter import PrometheusMetrics
from prometheus_client import Gauge

app = Flask(__name__)

# --- MONITOREO (Prometheus/Grafana) ---
# Configura automáticamente el endpoint /metrics
metrics = PrometheusMetrics(app)
# Gauge para visualizar en Grafana el valor de la proyección
PREDICCION_VALOR = Gauge('ultima_prediccion_ventas', 'Valor de la última predicción realizada')

# --- CARGA DEL MODELO, ESCALADORES Y ENCODER ---
base_path = os.path.dirname(__file__)
modelo_path = os.path.join(base_path, 'modelo_ventas.pkl')
modelo = None
scaler_x = None
scaler_y = None
le = None
lista_secciones = []
mensaje_modelo = "✅ Sistema Conectado - Modo Segmentado"

if os.path.exists(modelo_path):
    try:
        with open(modelo_path, 'rb') as f:
            # Desempaquetamos el diccionario que contiene todos los componentes del modelo
            paquete = pickle.load(f)
            modelo = paquete['model']
            scaler_x = paquete['scaler_x']
            scaler_y = paquete['scaler_y']
            le = paquete['label_encoder']
            lista_secciones = paquete['secciones']
    except Exception as e:
        mensaje_modelo = f"❌ Error al cargar modelo: {str(e)}"
else:
    mensaje_modelo = "⚠️ Archivo modelo_ventas.pkl no encontrado"

# --- DISEÑO HTML (Interfaz con selector dinámico de Sectores) ---
HTML = '''
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>IA Ventas por Sector</title>
        <style>
            :root { --primary: #4361ee; --bg: #f8f9fa; }
            body { 
                font-family: 'Segoe UI', system-ui, sans-serif; 
                background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
                display: flex; justify-content: center; align-items: center; 
                min-height: 100vh; margin: 0; 
            }
            .card { 
                background: rgba(255, 255, 255, 0.95); 
                padding: 2.5rem; border-radius: 24px; 
                box-shadow: 0 20px 40px rgba(0,0,0,0.3); 
                width: 100%; max-width: 450px; text-align: center;
            }
            select, button { 
                width: 100%; padding: 12px; margin-top: 8px; border-radius: 10px; 
                border: 1px solid #ddd; font-size: 0.9rem; box-sizing: border-box;
            }
            label { display: block; text-align: left; font-size: 0.75rem; font-weight: bold; margin-top: 12px; color: #555; }
            button { 
                background: var(--primary); color: white; border: none; 
                font-weight: bold; cursor: pointer; margin-top: 20px; transition: 0.3s;
            }
            button:hover { background: #3046c9; transform: translateY(-2px); }
            .res { 
                margin-top: 20px; padding: 15px; background: #eef2ff; 
                border-radius: 16px; border: 1px dashed var(--primary);
                animation: fadeIn 0.6s ease-out;
            }
            @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        </style>
    </head>
    <body>
        <div class="card">
            <h1 style="color:var(--primary); margin-bottom:5px;">Predictor por Sector</h1>
            <span style="font-size:0.75rem; color:#888;">{{ estado }}</span>
            
            <form action="/predict" method="post">
                <label>SECTOR ECONÓMICO</label>
                <select name="seccion">
                    {% for sec in secciones %}
                        <option value="{{ sec }}">{{ sec }}</option>
                    {% endfor %}
                </select>

                <label>MES DE PROYECCIÓN</label>
                <select name="mes_nombre">
                    <option value="1">Enero</option><option value="2">Febrero</option>
                    <option value="3">Marzo</option><option value="4">Abril</option>
                    <option value="5">Mayo</option><option value="6">Junio</option>
                    <option value="7">Julio</option><option value="8">Agosto</option>
                    <option value="9">Septiembre</option><option value="10">Octubre</option>
                    <option value="11">Noviembre</option><option value="12">Diciembre</option>
                </select>

                <label>AÑO</label>
                <select name="anio">
                    <option value="2024">2024</option>
                    <option value="2025">2025</option>
                    <option value="2026">2026</option>
                </select>

                <button type="submit">Generar Proyección Sectorial</button>
            </form>

            {% if prediccion %}
                <div class="res">
                    <p style="margin:0; font-size: 0.75rem; color: #4361ee; font-weight: bold;">{{ label_info }}:</p>
                    <h2 style="margin:5px 0 0 0; font-size: 1.8rem; color: #1e3c72;">$ {{ prediccion }}</h2>
                </div>
            {% endif %}
        </div>
    </body>
    </html>
'''

@app.route('/')
def home():
    return render_template_string(HTML, estado=mensaje_modelo, secciones=lista_secciones)

@app.route('/predict', methods=['POST'])
def predict():
    try:
        # Captura de datos del formulario
        sec_nombre = request.form.get('seccion')
        mes_sel = int(request.form.get('mes_nombre'))
        anio_sel = int(request.form.get('anio'))

        if modelo is None or le is None:
            return render_template_string(HTML, estado="❌ Modelo no cargado", secciones=lista_secciones)

        # 1. Transformar el nombre del sector a su ID numérico
        sec_id = le.transform([sec_nombre])[0]
        
        # 2. Preparar entrada para el modelo [Sector, Mes, Año]
        entrada = np.array([[sec_id, mes_sel, anio_sel]])
        
        # 3. Escalar entrada para que coincida con el entrenamiento
        entrada_scaled = scaler_x.transform(entrada)
        
        # 4. Realizar predicción
        pred_scaled = modelo.predict(entrada_scaled)
        
        # 5. Invertir el escalado del resultado para obtener el valor real en dólares
        resultado_final = scaler_y.inverse_transform(pred_scaled.reshape(-1, 1))[0][0]

        # Actualizar métrica en Prometheus/Grafana
        PREDICCION_VALOR.set(resultado_final)
        
        nombres_meses = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        info_texto = f"Proyección {sec_nombre} ({nombres_meses[mes_sel]} {anio_sel})"

        return render_template_string(HTML, 
                                     estado=mensaje_modelo, 
                                     secciones=lista_secciones,
                                     prediccion=f"{resultado_final:,.2f}", 
                                     label_info=info_texto)
    except Exception as e:
        return f"Error interno en el servidor: {str(e)}", 500

if __name__ == '__main__':
    # Ejecución en el puerto 5000 para el contenedor de Docker
    app.run(host='0.0.0.0', port=5000)