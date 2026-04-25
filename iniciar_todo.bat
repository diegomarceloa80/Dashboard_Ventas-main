@echo off
title Pipeline MLOps - Automatizacion de Ventas
echo ==========================================================
echo   INICIANDO INFRAESTRUCTURA MLOPS
echo ==========================================================

:: 1. Levantar contenedores base
echo [1/5] Levantando servicios en Docker Desktop...
docker-compose up -d
if %errorlevel% neq 0 (echo ERROR en Docker Compose && pause && exit)

:: 2. Entrenamiento del Modelo (Asegurando librerias)
echo [2/5] Instalando dependencias y entrenando modelo...
pip install python-dotenv wandb pandas scikit-learn
python train.py
if %errorlevel% neq 0 (echo ERROR en el entrenamiento && pause && exit)

:: 3. Preparar Minikube
echo [3/5] Verificando Minikube...
minikube status >nul 2>&1
if %errorlevel% neq 0 (
    echo Iniciando Minikube...
    minikube start
)

:: 4. Construir Imagen en Minikube
echo [4/5] Construyendo imagen en entorno Minikube...
@FOR /f "tokens=*" %%i IN ('minikube -p minikube docker-env') DO @%%i
docker build -t app-ventas:latest .

:: 5. Desplegar en Kubernetes
echo [5/5] Aplicando configuracion de Kubernetes...
kubectl apply -f deployment.yaml
echo Esperando a que los Pods esten listos...
kubectl wait --for=condition=ready pod -l app=ventas --timeout=60s

echo ==========================================================
echo   PROCESO COMPLETADO
echo ==========================================================
minikube service app-ventas-service
pause