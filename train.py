import pandas as pd
from sklearn.ensemble import RandomForestRegressor
import pickle
import os
import wandb

# Configuración para que no pida llaves en Docker
os.environ["WANDB_MODE"] = "offline"
wandb.init(project="proyecto-ventas-final", name="entrenamiento-v1")

# 1. Cargar datos
if not os.path.exists('datos_ventas.csv'):
    print("ERROR: Archivo datos_ventas.csv no encontrado")
    wandb.finish()
else:
    df = pd.read_csv('datos_ventas.csv', encoding='latin1', sep=None, engine='python', on_bad_lines='skip')

    # 2. Identificar columnas de meses (buscamos las que tienen '/')
    columnas_ventas = [c for c in df.columns if '/' in str(c)]
    
    X_list = []
    y_list = []
    
    # 3. Limpieza y preparación de listas
    for i, col in enumerate(columnas_ventas):
        # Limpiamos la columna: quitamos comas, símbolos de dólar y espacios
        serie_limpia = df[col].astype(str).str.replace(',', '').str.replace('$', '').str.strip()
        
        # Convertimos a número (lo que no sea número será 0)
        valores_numericos = pd.to_numeric(serie_limpia, errors='coerce').fillna(0)
        total_mes = valores_numericos.sum()
        
        if total_mes > 0:
            X_list.append([i + 1]) 
            y_list.append(total_mes)

    # 4. Entrenamiento y Registro
    if len(y_list) > 0:
        modelo = RandomForestRegressor(n_estimators=100, random_state=42)
        modelo.fit(X_list, y_list)

        # Guardamos el archivo .pkl
        with open('modelo_ventas.pkl', 'wb') as f:
            pickle.dump(modelo, f)

        # Registro en W&B
        wandb.log({
            "total_meses_entrenados": len(y_list), 
            "ultimo_valor_real": y_list[-1],
            "maxima_venta_historica": max(y_list)
        })
        print(f"¡ÉXITO! Modelo entrenado con {len(y_list)} meses.")
    else:
        print("ERROR: No se encontraron datos válidos.")
    
    wandb.finish()