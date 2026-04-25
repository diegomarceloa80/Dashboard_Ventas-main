FROM python:3.9-slim
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiamos los archivos de la app
COPY app.py .
COPY datos_ventas.csv .
# Forzamos la copia del modelo específico
COPY modelo_ventas.pkl . 

EXPOSE 5000
CMD ["python", "app.py"]