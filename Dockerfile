# Usamos una imagen de Python ligera
FROM python:3.9-slim

# Establecemos la carpeta de trabajo dentro del contenedor
WORKDIR /app

# Copiamos el archivo de requerimientos primero para aprovechar el caché de Docker
COPY requirements.txt .

# Instalamos las librerías necesarias
RUN pip install --no-cache-dir -r requirements.txt

# COPIAMOS TODO EL CONTENIDO (Esto incluye app.py y el modelo_ventas.pkl)
COPY . .

# Exponemos el puerto donde corre la app
EXPOSE 5000

# Comando para arrancar la aplicación
CMD ["python", "app.py"]