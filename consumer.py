from kafka import KafkaConsumer
import json
import sys

# Configurar el Consumidor
try:
    consumer = KafkaConsumer(
        'ventas_tiempo_real',
        bootstrap_servers=['localhost:9092'],
        auto_offset_reset='earliest',  # Lee desde el primer mensaje disponible
        enable_auto_commit=True,
        group_id='mi-grupo-ventas',
        value_deserializer=lambda m: json.loads(m.decode('utf-8'))
    )

    print("👂 ESCUCHANDO: Esperando datos desde Kafka... (Presiona Ctrl+C para salir)")

    for message in consumer:
        data = message.value
        # Imprime lo que recibe en tiempo real
        print(f"📥 RECIBIDO -> Mes: {data.get('mes')}, Venta: {data.get('venta')}")

except Exception as e:
    print(f"❌ Error de conexión: {e}")
    sys.exit(1)