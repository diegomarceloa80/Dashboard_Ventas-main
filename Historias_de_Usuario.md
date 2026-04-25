# Documentación del Proyecto: Sistema de Monitoreo de Predicción de Ventas (MLOps)

## Introducción
Este proyecto implementa un ciclo de vida de Machine Learning Operacional (MLOps) utilizando una arquitectura de microservicios sobre **Kubernetes**. El sistema permite realizar predicciones de ventas a través de una API Flask y monitorear el rendimiento y comportamiento del modelo en tiempo real mediante **Prometheus** y **Grafana**.

---

## Historias de Usuario

### Historia de Usuario 1: Predicción de Ventas para el Usuario Final
**Como** Analista de Ventas,  
**quiero** ingresar datos de sectores y parámetros de mercado en una interfaz web,  
**para** obtener una predicción automática del volumen de ventas y tomar decisiones comerciales basadas en datos.

* **Criterios de Aceptación:**
    * La aplicación debe estar disponible en un navegador web (Opera/Chrome).
    * El usuario debe poder enviar datos a través de un formulario.
    * El sistema debe responder con un valor numérico de predicción procesado por el modelo de IA.
    * La respuesta debe ser rápida (baja latencia).

---

### Historia de Usuario 2: Visibilidad de Métricas de Negocio (Dashboard)
**Como** Gerente de Operaciones,  
**quiero** visualizar en un tablero gráfico el valor de las últimas predicciones realizadas,  
**para** identificar tendencias de demanda y posibles anomalías sin necesidad de revisar logs técnicos.

* **Criterios de Aceptación:**
    * Grafana debe mostrar un panel con la métrica `ultima_prediccion_ventas`.
    * Los datos deben actualizarse automáticamente cada pocos segundos.
    * El dashboard debe ser accesible de forma independiente a la aplicación de predicción.
    * La visualización debe permitir ver el histórico de los últimos 5 a 15 minutos.

---

### Historia de Usuario 3: Estabilidad y Monitoreo de Infraestructura
**Como** Ingeniero de MLOps / DevOps,  
**quiero** monitorear la salud del servicio y el tráfico de solicitudes HTTP,  
**para** asegurar que la infraestructura en Kubernetes sea estable y reaccionar ante posibles caídas del servicio.

* **Criterios de Aceptación:**
    * Prometheus debe realizar *scraping* automático del endpoint `/metrics` de la aplicación Flask.
    * Se debe poder verificar el estado del servicio mediante la métrica `up`.
    * Se deben contabilizar las solicitudes totales mediante `flask_http_request_total`.
    * El sistema debe permitir el escalamiento de pods en Kubernetes si la demanda aumenta.

---

## Detalles Técnicos de la Implementación

| Componente | Tecnología | Función |
| :--- | :--- | :--- |
| **Backend** | Python / Flask | Sirve el modelo de ML y expone métricas. |
| **Contenedor** | Docker | Empaqueta la app y sus dependencias. |
| **Orquestador** | Kubernetes (Minikube) | Gestiona el despliegue y la red de los pods. |
| **Monitoreo** | Prometheus | Recolecta y almacena las métricas temporales. |
| **Visualización** | Grafana | Crea el tablero de control para el usuario final. |

---

## Guía de Visualización de Datos
Para visualizar los datos en el entorno de desarrollo:
1. Asegurar que el túnel de Prometheus esté activo: `kubectl port-forward svc/prometheus-service -n monitoring 9090:8080`.
2. Acceder a Grafana en `localhost:3000`.
3. Consultar la métrica `ultima_prediccion_ventas` para ver el impacto del modelo de IA en tiempo real.