import pandas as pd
from kafka import KafkaProducer
import json
import time

# Conectar con el servidor Kafka
producer = KafkaProducer(
    bootstrap_servers=['localhost:9092'],
    value_serializer=lambda v: json.dumps(v).encode('utf-8'),
    api_version=(0, 10, 1)  # <--- AGREGA ESTA LÍNEA EXACTAMENTE
)

# Leer los datos que ya tenemos
df = pd.read_csv('datos_ventas.csv', encoding='latin1', sep=None, engine='python')

print("Iniciando Streaming de ventas...")

# Simular envío de datos en tiempo real
for index, row in df.iterrows():
    data = {"mes": index + 1, "venta": str(row.iloc[1])}
    producer.send('ventas_tiempo_real', value=data)
    print(f"Enviando: {data}")
    time.sleep(2) # Espera 2 segundos entre cada venta para ver el flujo