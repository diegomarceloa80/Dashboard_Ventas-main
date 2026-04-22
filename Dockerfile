FROM python:3.9-slim
WORKDIR /app

# Copiamos la lista de requerimientos
COPY requirements.txt .

# Instalamos todo lo de la lista
RUN pip install --no-cache-dir -r requirements.txt

# Copiamos el resto de los archivos
COPY . .

# Ejecutamos el entrenamiento (generará el modelo y conectará con W&B)
RUN python train.py

# Encendemos la web
EXPOSE 5000
CMD ["python", "app.py"]