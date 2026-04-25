import os
import pandas as pd
import numpy as np
import pickle
import wandb
from dotenv import load_dotenv
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import StandardScaler, LabelEncoder

load_dotenv()
wandb.init(project="proyecto-ventas-final", name="entrenamiento-por-seccion")

# 1. Cargar y Limpiar Datos
df = pd.read_csv('datos_ventas.csv', encoding='latin1', sep=';', engine='python')
columnas_ventas = [c for c in df.columns if '/' in str(c)]

X_list = []
y_list = []

# 2. Ingeniería de Características por Sección
le = LabelEncoder()
df['SECCION_ID'] = le.fit_transform(df['SECCIÓN'].astype(str))

for col in columnas_ventas:
    mes, anio = map(int, col.split('/'))
    # Limpiar valores monetarios
    serie_limpia = df[col].astype(str).str.replace('.', '').str.replace(',', '.').str.replace('$', '').str.strip()
    valores = pd.to_numeric(serie_limpia, errors='coerce').fillna(0)
    
    # Agrupar por sección para ese mes específico
    ventas_por_seccion = df.assign(venta=valores).groupby('SECCION_ID')['venta'].sum()
    
    for sec_id, total in ventas_por_seccion.items():
        if total > 0:
            X_list.append([sec_id, mes, anio])
            y_list.append(total)

X = np.array(X_list)
y = np.array(y_list).reshape(-1, 1)

# 3. Escalamiento y Modelo
scaler_x = StandardScaler()
scaler_y = StandardScaler()
X_scaled = scaler_x.fit_transform(X)
y_scaled = scaler_y.fit_transform(y)

modelo = RandomForestRegressor(n_estimators=300, max_depth=15, random_state=42)
modelo.fit(X_scaled, y_scaled.ravel())

# 4. Guardar Paquete Completo
paquete = {
    'model': modelo,
    'scaler_x': scaler_x,
    'scaler_y': scaler_y,
    'label_encoder': le,
    'secciones': list(le.classes_)
}

with open('modelo_ventas.pkl', 'wb') as f:
    pickle.dump(paquete, f)

wandb.log({"total_registros": len(y_list), "num_secciones": len(le.classes_)})
wandb.finish()