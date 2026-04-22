# DASHBOARD DE VENTAS - CÁMARA DE INDUSTRIAS Y PRODUCCIÓN
# Autor: Juan C. Ambuludi

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Dict, List, Tuple
import re
import base64
import unicodedata
import os
import time
import tempfile
import shutil

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots   # ← NUEVO: para eje Y secundario

try:
    from PIL import Image
except Exception:
    Image = None

BASE_DIR = Path(__file__).resolve().parent
INDICE_EXPECT_FILE = BASE_DIR / "indice.xlsx"
INCIDENCIAS_FILE   = BASE_DIR / "incidencias.xlsx"

SRI_RECAUD_FILE = None

RECAUD_TAX_GROUPS = {
    "IVA": "IMPUESTO AL VALOR AGREGADO",
    "IR":  "IMPUESTO A LA RENTA GLOBAL",
    "ISD": "SALIDA DE DIVISAS",
    "ICE": "IMPUESTO A LOS CONSUMOS ESPECIALES",
}

HAS_GEOPANDAS = False
try:
    import geopandas as gpd
    import warnings
    warnings.filterwarnings('ignore', category=UserWarning, module='geopandas')
    HAS_GEOPANDAS = True
except Exception:
    HAS_GEOPANDAS = False


RECAUD_NETA_PDF = {
    "cur":   1_860_000_000.0,
    "prev":  1_662_000_000.0,
    "yoy":   (1860 / 1662) - 1,
    "abs":   198_000_000.0,
    "label": "Ene 2026 vs Ene 2025",
}


# =============================================================================
# SAFE READ EXCEL
# =============================================================================
def _copy_to_temp(src_path: str) -> str:
    src_path = str(src_path)
    if not os.path.exists(src_path):
        raise FileNotFoundError(f"No existe el archivo: {src_path}")
    base    = os.path.basename(src_path)
    tmp_dir = tempfile.mkdtemp(prefix="tmp_excel_")
    tmp_path = os.path.join(tmp_dir, base)
    shutil.copy2(src_path, tmp_path)
    return tmp_path


def safe_read_excel(path: str, sheet_name, engine: str = "openpyxl",
                    retries: int = 5, sleep_s: float = 0.7, **kwargs) -> pd.DataFrame:
    last_err = None
    for _ in range(max(1, retries)):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, engine=engine, **kwargs)
        except (PermissionError, OSError) as e:
            last_err = e
            time.sleep(sleep_s)
    try:
        tmp_path = _copy_to_temp(path)
        return pd.read_excel(tmp_path, sheet_name=sheet_name, engine=engine, **kwargs)
    except Exception as e:
        raise last_err if last_err is not None else e


# =============================================================================
# CONFIGURACIÓN STREAMLIT
# =============================================================================
st.set_page_config(
    page_title="Panel Ejecutivo de Ventas | Ecuador",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =============================================================================
# AUTENTICACIÓN
# =============================================================================
USUARIOS_AUTORIZADOS = {
    "jcambuludi@cip.org.ec": "clave123",
    "mfalcon@cip.org.ec":    "vision2025",
    "abonilla@cip.org.ec":   "cip2025",
    "gsalinas@cip.org.ec":    "GSalinas@",
    "pjimenez@cip.org.ec":   "vision2025",
    "afiliadocip":  "afiliadocip",
    "staffcip":              "staffcip",
}

USUARIOS_NOMBRES = {
    "jcambuludi@cip.org.ec": "Juan Ambuludi",
    "mfalcon@cip.org.ec":    "Mauricio Falcón",
    "abonilla@cip.org.ec":   "Ariana Bonilla",
    "gsalinas@cip.org.ec":    "Gabriela Salinas",
    "pjimenez@cip.org.ec":   "Pablo Jiménez",
    "afiliadocip":  "Afiliado CIP",
    "staffcip":              "Staff CIP",
}

def _normalize_user(user: str) -> str:
    user = "" if user is None else str(user).strip()
    return user.lower()


def _login_success(user: str):
    user_key = _normalize_user(user)
    st.session_state.auth = True
    st.session_state.usuario = user
    st.session_state.usuario_nombre = USUARIOS_NOMBRES.get(
        user_key, user.split("@")[0].replace("_", " ").title()
    )
    st.rerun()


def _logout():
    st.session_state.auth = False
    st.session_state.usuario = None
    st.session_state.usuario_nombre = None
    st.rerun()


if "auth" not in st.session_state:
    st.session_state.auth = False
    st.session_state.usuario = None
    st.session_state.usuario_nombre = None
if "modulo_activo" not in st.session_state:
    st.session_state.modulo_activo = None


# =============================================================================
# FLUENT DESIGN SYSTEM
# =============================================================================
@dataclass
class FluentColors:
    white:        str = "#FFFFFF"
    gray10:       str = "#FAF9F8"
    gray20:       str = "#F3F2F1"
    gray30:       str = "#EDEBE9"
    gray40:       str = "#E1DFDD"
    gray50:       str = "#D2D0CE"
    gray60:       str = "#C8C6C4"
    gray90:       str = "#605E5C"
    gray130:      str = "#323130"
    gray160:      str = "#252423"
    black:        str = "#000000"
    theme_primary:str = "#2F3A4A"
    theme_dark:   str = "#1F2937"
    theme_darker: str = "#111827"
    theme_light:  str = "#CBD5E1"
    success:      str = "#107C10"
    warning:      str = "#D97706"
    error:        str = "#D13438"
    info:         str = "#2563EB"


colors = FluentColors()
RANGE_SLIDER = dict(visible=False)

# ─── Paleta del boletín de ventas ────────────────────────────────────────────
BULLETIN_BLUE        = "#2070E0"
BULLETIN_BLUE_LIGHT  = "#3090F0"
BULLETIN_BLUE_DARK   = "#1A4FA3"
BULLETIN_GOLD        = "#E0A020"
BULLETIN_GOLD_DARK   = "#C08010"
BULLETIN_SLATE       = "#5F6B7A"
BULLETIN_SLATE_LIGHT = "#D9E2F2"
BULLETIN_RED_SOFT    = "#E07A5F"
BULLETIN_NEUTRAL     = "#F8FAFC"
BULLETIN_POSITIVE    = BULLETIN_BLUE
BULLETIN_NEGATIVE    = BULLETIN_GOLD_DARK
BULLETIN_CURRENT     = BULLETIN_BLUE_DARK
BULLETIN_PREVIOUS    = BULLETIN_SLATE_LIGHT


def bulletin_diverging_color(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return BULLETIN_PREVIOUS
    return BULLETIN_POSITIVE if float(value) >= 0 else BULLETIN_NEGATIVE

# ─── Fuente global del proyecto ──────────────────────────────────────────────
FONT_FAMILY = "'DM Sans', 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
FONT_IMPORT  = (
    "https://fonts.googleapis.com/css2?"
    "family=DM+Sans:wght@400;500;600;700;800&"
    "family=Plus+Jakarta+Sans:wght@600;700;800;900&"
    "display=swap"
)

def make_range_selector():
    return dict(
        buttons=[
            dict(count=12, label="12M", step="month", stepmode="backward"),
            dict(count=24, label="24M", step="month", stepmode="backward"),
            dict(step="all", label="Todo"),
        ],
        bgcolor=colors.gray20,
        activecolor=colors.theme_primary,
        font=dict(size=10, color=colors.gray130),
        x=0, y=1.02,
        xanchor="left", yanchor="bottom",
    )


# =============================================================================
# CSS GLOBAL — DISEÑO EJECUTIVO MEJORADO
# =============================================================================
FLUENT_CSS = f"""
<style>
@import url('{FONT_IMPORT}');

/* ── Base ─────────────────────────────────────────────────────────────────── */
html, body, .stApp, .main, .block-container,
div, p, span, label, input, textarea, button, select {{
  font-family: {FONT_FAMILY} !important;
  box-sizing: border-box;
}}

.material-symbols-outlined,
.material-symbols-rounded,
.material-symbols-sharp,
[class*="material-symbols"],
[data-testid="stIconMaterial"] {{
  font-family: "Material Symbols Outlined", "Material Symbols Rounded", "Material Symbols Sharp" !important;
}}

.stApp {{ background-color: {colors.gray10}; }}
...

#MainMenu {{visibility: hidden;}}
footer   {{visibility: hidden;}}
header   {{visibility: hidden;}}
div[data-testid="stHeader"]       {{ display: none !important; }}
div[data-testid="stToolbar"]      {{ display: none !important; }}
div[data-testid="stDecoration"]   {{ display: none !important; }}
div[data-testid="stStatusWidget"] {{ display: none !important; }}

.main, .main .block-container,
div[data-testid="stAppViewContainer"],
div[data-testid="stAppViewContainer"] > .main,
div[data-testid="stAppViewBlockContainer"],
div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlock"] > div:first-child {{
  padding-top: .15rem !important;
  margin-top: 0 !important;
}}
.main .block-container {{
  padding-top: .15rem !important;
  padding-left: clamp(1rem, 1.8vw, 2rem) !important;
  padding-right: clamp(1rem, 1.8vw, 2rem) !important;
  padding-bottom: 1.5rem !important;
  max-width: 100% !important;
  width: 100% !important;
}}

/* Ocultar barra de herramientas de dataframes */
[data-testid="stDataFrame"] [data-testid="stElementToolbar"] {{ display: none !important; }}
button[title="Download"]   {{ display: none !important; }}
[data-testid="stElementToolbar"] {{ display: none !important; }}

/* ── Hero ────────────────────────────────────────────────────────────────── */
.hero {{
  position: relative;
  width: 100%;
  overflow: hidden;
  background:
    radial-gradient(circle at 14% 26%, rgba(56,189,248,.25), transparent 44%),
    radial-gradient(circle at 86% 84%, rgba(167,139,250,.20), transparent 42%),
    linear-gradient(132deg, #020617 0%, #0b1224 52%, #172554 100%);
  border-radius: 24px;
  padding: 1.35rem 1.7rem;
  box-shadow: 0 18px 44px rgba(2,6,23,.36);
  margin: 0 0 1rem 0;
  border: 1px solid rgba(125,211,252,.24);
}}
.hero::before {{
  content: "";
  position: absolute;
  inset: 0;
  background: linear-gradient(110deg, rgba(255,255,255,.12), transparent 32%, transparent 72%, rgba(255,255,255,.06));
  pointer-events: none;
}}
.hero-wrap {{
  position: relative;
  z-index: 1;
  display: flex; align-items: center;
  justify-content: space-between; gap: 14px;
}}
.hero-kicker {{
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 2.2px;
  font-weight: 800;
  color: #dbeafe;
  margin-bottom: 8px;
}}
.hero-kicker::before {{
  content: "";
  width: 18px;
  height: 2px;
  border-radius: 999px;
  background: linear-gradient(90deg,#60a5fa,#a78bfa);
}}
.hero-title {{
  font-family: 'Plus Jakarta Sans', {FONT_FAMILY};
  font-size: clamp(1.7rem, 2.35vw, 2.26rem);
  font-weight: 900;
  color: #f8fafc;
  margin: 0; line-height: 1.15; letter-spacing: -.6px;
  text-shadow: 0 6px 20px rgba(2,6,23,.32);
}}
.hero-subtitle {{
  font-size: 13.2px; color: #cbd5e1;
  margin: 8px 0 0 0; font-weight: 600;
}}
.hero-logo {{
  display: flex; align-items: center; justify-content: center;
  width: 90px; height: 90px;
  border: 1px solid rgba(147,197,253,.44);
  background: linear-gradient(120deg,rgba(2,6,23,.88),rgba(15,23,42,.92) 56%,rgba(30,64,175,.92));
  border-radius: 18px; flex-shrink: 0; overflow: hidden;
  box-shadow: inset 0 0 0 1px rgba(255,255,255,.06), 0 10px 24px rgba(2,6,23,.34);
}}
.hero-logo img {{
  max-width: 76%; max-height: 76%;
  object-fit: contain;
  filter: brightness(1.12) contrast(1.18) saturate(1.1);
}}

/* ── Module launcher ─────────────────────────────────────────────────────── */
.launcher-badge {{
  display: inline-flex; align-items: center; gap: 7px;
  font-size: .71rem; font-weight: 800;
  letter-spacing: .17em; text-transform: uppercase;
  color: #93c5fd;
  background: rgba(59,130,246,.11);
  border: 1px solid rgba(96,165,250,.28);
  border-radius: 999px; padding: 7px 16px;
  backdrop-filter: blur(6px);
}}
.launcher-badge::before {{
  content: ""; width: 7px; height: 7px;
  border-radius: 50%; background: #60a5fa;
  box-shadow: 0 0 8px #60a5fa; flex-shrink: 0;
}}

.mod-card {{
  position: relative;
  background: linear-gradient(145deg,rgba(255,255,255,.065),rgba(255,255,255,.018));
  border: 1px solid rgba(148,163,184,.16);
  border-radius: 22px; padding: clamp(20px,2.6vw,30px);
  backdrop-filter: blur(14px);
  transition: transform .22s ease, border-color .22s ease, box-shadow .22s ease;
  overflow: hidden; height: 100%;
}}
.mod-card::before {{
  content: ""; position: absolute; top: 0; left: 0; right: 0; height: 1px;
  background: linear-gradient(90deg,transparent,rgba(148,163,184,.28),transparent);
}}
.mod-card:hover {{
  transform: translateY(-4px);
  border-color: rgba(96,165,250,.40);
  box-shadow: 0 0 0 1px rgba(96,165,250,.10),
              0 22px 44px rgba(0,0,0,.36),
              0 0 55px rgba(59,130,246,.07);
}}
.mod-icon {{
  width: 50px; height: 50px; border-radius: 15px;
  display: flex; align-items: center; justify-content: center;
  font-size: 1.45rem; margin-bottom: 16px;
  background: linear-gradient(135deg,rgba(29,78,216,.34),rgba(124,58,237,.18));
  border: 1px solid rgba(96,165,250,.24);
  box-shadow: 0 4px 14px rgba(29,78,216,.22);
}}
.mod-name {{
  font-family: 'Plus Jakarta Sans', {FONT_FAMILY};
  font-size: 1.2rem; font-weight: 800;
  color: #f1f5f9; margin: 0 0 7px 0; letter-spacing: -.25px;
}}
.mod-desc {{
  font-size: .86rem; color: #64748b;
  line-height: 1.55; margin: 0 0 18px 0; font-weight: 500;
}}
.mod-meta {{
  display: flex; align-items: center;
  gap: 9px; margin-bottom: 18px; flex-wrap: wrap;
}}
.mod-tag {{
  display: inline-flex; align-items: center; gap: 5px;
  font-size: .68rem; font-weight: 800;
  letter-spacing: .06em; text-transform: uppercase;
  padding: 3px 10px; border-radius: 999px;
}}
.mod-tag.live {{
  color: #4ade80; background: rgba(74,222,128,.10);
  border: 1px solid rgba(74,222,128,.26);
}}
.mod-tag.live::before {{
  content: ""; width: 6px; height: 6px; border-radius: 50%;
  background: #4ade80; box-shadow: 0 0 6px #4ade80;
}}
.mod-tag.count {{
  color: #94a3b8; background: rgba(148,163,184,.09);
  border: 1px solid rgba(148,163,184,.18);
}}
.mod-stats {{
  display: grid; grid-template-columns: repeat(3,1fr);
  gap: 7px; margin-bottom: 20px; padding: 13px;
  background: rgba(15,23,42,.36);
  border-radius: 13px; border: 1px solid rgba(148,163,184,.10);
}}
.mod-stat-item {{ text-align: center; }}
.mod-stat-val  {{
  font-size: .95rem; font-weight: 800; color: #e2e8f0;
  display: block; letter-spacing: -.25px;
}}
.mod-stat-lbl {{
  font-size: .62rem; font-weight: 700; color: #475569;
  text-transform: uppercase; letter-spacing: .06em;
  margin-top: 2px; display: block;
}}
.mod-card.locked {{ opacity: .50; pointer-events: none; }}
.mod-card .stButton > button {{
  width: 100%; min-height: 50px;
  font-size: .84rem; font-weight: 800;
  border-radius: 14px;
  border: 1px solid rgba(167,139,250,.52) !important;
  background: linear-gradient(120deg,#0f172a,#1d4ed8 55%,#7c3aed) !important;
  color: #f8fafc !important;
  letter-spacing: .08em; text-transform: uppercase;
  box-shadow: 0 12px 24px rgba(30,64,175,.35),
              inset 0 1px 0 rgba(255,255,255,.14) !important;
  transition: transform .2s ease, box-shadow .2s ease, filter .2s ease !important;
}}
.mod-card .stButton > button:hover {{
  transform: translateY(-2px) scale(1.01) !important;
  border-color: rgba(196,181,253,.75) !important;
  box-shadow: 0 16px 30px rgba(37,99,235,.34),
              0 0 0 2px rgba(191,219,254,.22),
              inset 0 1px 0 rgba(255,255,255,.18) !important;
  filter: saturate(1.08) !important;
}}

/* ── Filtros y etiquetas ─────────────────────────────────────────────────── */
.filter-title {{
  margin: 0 0 12px 0;
  font-size: 1.02rem;
  color: #0f172a;
  font-weight: 650;
  letter-spacing: -.1px;
  font-weight: 700;
  letter-spacing: -.15px;
}}
.section-header {{
  font-size: 14.5px; font-weight: 900; color: #0f172a;
  margin: 14px 0 10px 0; padding-bottom: 6px;
  border-bottom: 2px solid {colors.gray40};
  letter-spacing: -.1px;
}}
.card-title {{
  font-size: 12.5px; font-weight: 800;
  color: {colors.gray130}; margin: 0 0 8px 0;
}}

/* ── Tabs ────────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {{
  gap: .52rem;
  padding: .52rem;
  background: linear-gradient(135deg,#e8edf3,#dde4ed 56%,#d7dee8);
  border: 1px solid rgba(71,85,105,.18);
  border-radius: 18px;
  margin: 12px 0 14px 0;
  display: flex; flex-wrap: nowrap; width: 100%;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.92), 0 10px 22px rgba(15,23,42,.05);
}}
.stTabs [data-baseweb="tab"] {{
  height: 42px;
  background: rgba(255,255,255,.48);
  border: 1px solid rgba(148,163,184,.08); border-radius: 13px;
  padding: 0 14px; color: #1e293b;
  font-weight: 700; font-size: 11.8px;
  justify-content: center; flex: 1 1 0; min-width: 0;
  transition: all .22s ease;
}}
.stTabs [data-baseweb="tab"]:hover {{
  color: #0f172a;
  background: rgba(255,255,255,.86);
  border-color: rgba(100,116,139,.22);
}}
.stTabs [aria-selected="true"] {{
  color: #f8fafc !important;
  background: linear-gradient(135deg,#16324a,#0b1731 56%,#39437a) !important;
  border: 1px solid rgba(30,64,175,.30) !important;
  box-shadow: 0 14px 26px rgba(15,23,42,.18), inset 0 1px 0 rgba(255,255,255,.12);
  font-weight: 800 !important;
  transform: translateY(-1px);
}}

/* ── Contenedores con borde ──────────────────────────────────────────────── */
div[data-testid="stVerticalBlockBorderWrapper"] {{
  background: {colors.white};
  border: 1px solid {colors.gray40};
  border-radius: 14px;
  box-shadow: 0 2px 12px rgba(0,0,0,.05);
}}

/* ── Selectbox / Multiselect ─────────────────────────────────────────────── */
.stSelectbox > div > div,
.stMultiSelect > div > div {{
  border-radius: 12px;
  border: 1px solid #cbd5e1 !important;
  background: linear-gradient(160deg,#ffffff,#f8fbff) !important;
  font-size: 13px;
  box-shadow: 0 1px 0 rgba(255,255,255,.9), 0 7px 16px rgba(15,23,42,.05);
  transition: border-color .2s ease, box-shadow .2s ease;
}}
.stSelectbox > div > div:focus-within,
.stMultiSelect > div > div:focus-within {{
  border-color: #60a5fa !important;
  box-shadow: 0 0 0 3px rgba(96,165,250,.18), 0 8px 16px rgba(37,99,235,.12) !important;
}}

/* ── KPI cards ───────────────────────────────────────────────────────────── */
.kpi-card {{
  background: linear-gradient(150deg,#ffffff,#f8fbff);
  border-radius: 16px; padding: 13px;
  border: 1px solid #e9eef7;
  box-shadow: 0 4px 14px rgba(15,23,42,.05);
  transition: all .15s ease;
  box-sizing: border-box; width: 100%;
  overflow: hidden; min-height: 122px;
}}
[data-testid="column"] {{ min-width: 0 !important; }}
.kpi-card:hover {{
  border-color: #bfdbfe;
  box-shadow: 0 10px 22px rgba(30,64,175,.12);
  transform: translateY(-2px);
}}
.kpi-label {{
  font-size: 10px; color: #475569; font-weight: 700;
  margin: 0 0 5px 0; text-transform: uppercase;
  letter-spacing: .5px; white-space: normal;
}}
.kpi-value {{
  font-size: clamp(1.15rem, 1.5vw, 1.45rem);
  font-weight: 900; color: {colors.gray160};
  margin: 0 0 5px 0; line-height: 1.1;
  white-space: normal; overflow-wrap: anywhere;
}}
.kpi-sub {{
  font-size: 10.5px; font-weight: 600; color: {colors.gray90};
  margin-top: 5px; white-space: normal;
}}
.kpi-deltas {{
  display: flex; flex-wrap: wrap; gap: 4px; margin: 4px 0 2px 0;
}}
.kpi-delta {{
  font-size: 9.5px; font-weight: 800;
  display: inline-flex; align-items: center;
  gap: 3px; padding: 2px 7px; border-radius: 999px;
}}
.kpi-delta.positive {{ color:{colors.success}; background:rgba(16,124,16,.09); }}
.kpi-delta.negative {{ color:{colors.error};   background:rgba(209,52,56,.09); }}
.kpi-delta.neutral  {{ color:{colors.gray90};  background:{colors.gray30};     }}

/* ── Footnotes ───────────────────────────────────────────────────────────── */
.chart-footnote {{
  font-size: 10.5px; color: {colors.gray90};
  padding: 5px 2px 0 2px;
  border-top: 1px solid {colors.gray30}; margin-top: 4px;
}}
.panel-footnote {{
  font-size: 10.5px; color: {colors.gray90};
  background: {colors.gray20}; border: 1px solid {colors.gray40};
  border-radius: 10px; padding: 8px 12px;
  margin-top: 14px;
}}

/* ── Proyecciones ────────────────────────────────────────────────────────── */
.proj-note {{
  font-size: 12px; font-weight: 700; color: {colors.gray130};
  background: {colors.gray20}; border: 1px solid {colors.gray40};
  border-radius: 12px; padding: 10px 12px;
}}

/* ── Slider ──────────────────────────────────────────────────────────────── */
div[data-baseweb="slider"] {{
  padding: 4px 2px 2px 2px !important;
}}
div[data-baseweb="slider"] > div > div:first-child {{
  background: linear-gradient(90deg,#94a3b8,#64748b) !important;
  height: 6px !important; border-radius: 999px !important;
}}
div[data-baseweb="slider"] > div > div:nth-child(2) {{
  background: linear-gradient(90deg,#1e3a8a,#2563eb,#0ea5e9) !important;
  height: 6px !important; border-radius: 999px !important;
}}
div[data-baseweb="slider"] [role="slider"] {{
  width: 18px !important; height: 18px !important;
  border: 2px solid #1e3a8a !important;
  background: radial-gradient(circle at 30% 30%,#ffffff,#dbeafe) !important;
  box-shadow: 0 6px 16px rgba(30,64,175,.24) !important;
}}

/* ── Dataframe / Tabla ───────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {{
  border-radius: 12px !important;
  overflow: hidden !important;
}}
[data-testid="stDataFrame"] table {{
  font-size: 12.5px !important;
}}
[data-testid="stDataFrame"] th {{
  background: {colors.gray20} !important;
  font-weight: 800 !important;
  font-size: 11px !important;
  text-transform: uppercase !important;
  letter-spacing: .5px !important;
  color: {colors.gray130} !important;
}}

/* ── Usuario / sesión ────────────────────────────────────────────────────── */
.user-toolbar-shell {{
  width: 100%;
  display: flex;
  justify-content: flex-end;
  align-items: center;
  gap: .55rem;
  margin: .32rem 0 1rem 0;
}}
.session-pill {{
  display: inline-flex;
  align-items: center;
  gap: .5rem;
  white-space: nowrap;
  min-height: 54px;
  padding: 0 .9rem;
  border-radius: 16px;
  border: 1px solid rgba(148,163,184,.22);
  background: linear-gradient(135deg,rgba(255,255,255,.98),rgba(241,245,249,.96));
  box-shadow: 0 10px 24px rgba(15,23,42,.06), inset 0 1px 0 rgba(255,255,255,.92);
  color: #475569;
  font-size: .78rem;
  font-weight: 700;
}}
.session-pill::before {{
  content: "";
  width: 9px;
  height: 9px;
  border-radius: 999px;
  background: linear-gradient(135deg,#14b8a6,#22c55e);
  box-shadow: 0 0 0 4px rgba(20,184,166,.10);
}}
div[data-testid="stPopover"] > button {{
  min-height: 54px;
  width: 100%;
  border-radius: 16px !important;
  border: 1px solid rgba(30,64,175,.26) !important;
  background: linear-gradient(135deg,#13263a,#10203b 52%,#37406f) !important;
  color: #f8fafc !important;
  font-size: .92rem !important;
  font-weight: 700 !important;
  letter-spacing: -.01em;
  padding: .68rem .98rem !important;
  box-shadow: 0 16px 32px rgba(15,23,42,.16), inset 0 1px 0 rgba(255,255,255,.10) !important;
  transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease, filter .2s ease !important;
}}
div[data-testid="stPopover"] > button:hover {{
  transform: translateY(-1px);
  border-color: rgba(147,197,253,.46) !important;
  box-shadow: 0 18px 34px rgba(15,23,42,.18), 0 0 0 4px rgba(96,165,250,.08) !important;
  filter: saturate(1.06);
}}
div[data-testid="stPopover"] > button:focus-visible {{
  box-shadow: 0 0 0 4px rgba(96,165,250,.16), 0 16px 32px rgba(15,23,42,.18) !important;
}}
.user-popover-card {{
  background: linear-gradient(145deg,#ffffff,#f8fafc 60%,#eef2ff);
  border: 1px solid rgba(191,219,254,.85);
  border-radius: 16px;
  padding: .9rem 1rem;
  margin-bottom: .65rem;
}}
.user-popover-kicker {{
  font-size: .68rem;
  font-weight: 800;
  letter-spacing: .12em;
  text-transform: uppercase;
  color: #64748b;
  margin-bottom: .35rem;
}}
.user-popover-name {{
  font-size: 1rem;
  font-weight: 800;
  color: #0f172a;
  margin-bottom: .2rem;
}}
.user-popover-mail {{
  font-size: .78rem;
  color: #475569;
}}

.premium-topbar {{
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  border-radius: 18px;
  margin: .15rem 0 .9rem 0;
  padding: .7rem .85rem .7rem 1rem;
  background: linear-gradient(120deg,rgba(15,23,42,.88),rgba(30,64,175,.82) 52%,rgba(109,40,217,.82));
  border: 1px solid rgba(147,197,253,.34);
  box-shadow: 0 14px 28px rgba(15,23,42,.18);
}}
.premium-topbar-kicker {{
  font-size: 10px;
  font-weight: 800;
  letter-spacing: .18em;
  text-transform: uppercase;
  color: #bfdbfe;
}}
.premium-topbar-title {{
  margin-top: 2px;
  color: #f8fafc;
  font-size: 1rem;
  font-weight: 800;
}}

/* ── Filtro de año (rediseño) ────────────────────────────────────────────── */
.year-filter-label {{
  margin: 0 0 .35rem 0;
  font-size: .8rem;
  font-weight: 600;
  color: #334155;
  letter-spacing: .02em;
  text-transform: none;
}}
.year-filter-caption {{
  margin: .35rem 0 0 0;
  font-size: .8rem;
  font-weight: 600;
  color: #475569;
}}
div[data-testid="stVerticalBlockBorderWrapper"]:has(.filter-title) {{
  border: 1px solid rgba(191,219,254,.95) !important;
  background: linear-gradient(145deg,#ffffff 0%,#f3f8ff 55%,#eef5ff 100%) !important;
  box-shadow: 0 16px 30px rgba(37,99,235,.10), inset 0 1px 0 rgba(255,255,255,.88);
}}
div[data-testid="stVerticalBlockBorderWrapper"]:has(.filter-title) [data-testid="column"] {{
  padding-top: .1rem;
}}
div[data-testid="stVerticalBlockBorderWrapper"]:has(.filter-title) p,
div[data-testid="stVerticalBlockBorderWrapper"]:has(.filter-title) label {{
  color: #1e293b;
  font-size: .8rem;
  font-weight: 600;
}}

div[data-testid="stRadio"] [role="radiogroup"] {{
  background: #f8fafc;
  border: 1px solid #dbeafe;
  border-radius: 12px;
  padding: 4px;
  gap: 6px;
}}
div[data-testid="stRadio"] label[data-baseweb="radio"] {{
  background: transparent;
  border-radius: 10px;
  padding: 6px 10px !important;
  border: 1px solid transparent;
}}
div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) {{
  background: linear-gradient(135deg,#1d4ed8,#4f46e5) !important;
  border-color: rgba(37,99,235,.80) !important;
  box-shadow: 0 8px 16px rgba(37,99,235,.22);
}}
div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) span {{
  color: #ffffff !important;
  font-weight: 700 !important;
}}

@media (max-width: 768px) {{
  .stTabs [data-baseweb="tab"] {{ font-size: 10px; padding: 0 8px; }}
}}
</style>
"""

# ─── CSS de login ─────────────────────────────────────────────────────────────
def _find_background_image() -> Optional[Path]:
    preferred = [
        BASE_DIR / "fondo_dashboard.png",
        BASE_DIR / "fondo_dashboard.jpg",
        BASE_DIR / "background.png",
        BASE_DIR / "background.jpg",
        BASE_DIR / "assets" / "background.png",
        BASE_DIR / "assets" / "background.jpg",
    ]
    for candidate in preferred:
        if candidate.exists() and candidate.stat().st_size > 0:
            return candidate
    for ext in ("*.png", "*.jpg", "*.jpeg", "*.webp"):
        for file_path in BASE_DIR.glob(ext):
            lname = file_path.name.lower()
            if any(key in lname for key in ("fondo", "background", "bg")):
                return file_path
    return None


def _find_login_logo_file() -> Optional[Path]:
    for candidate in [
        BASE_DIR / "cip.png",
        BASE_DIR / "CIP.png",
        BASE_DIR / "assets" / "cip.png",
        BASE_DIR / "img" / "cip.png",
    ]:
        if candidate.exists() and candidate.stat().st_size > 0:
            return candidate
    for p in BASE_DIR.glob("*.png"):
        if "cip" in p.name.lower() and p.stat().st_size > 0:
            return p
    return None


def _login_logo_to_b64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("utf-8")


LOGIN_BG_URI = ""
_login_logo_file = _find_login_logo_file()
if _login_logo_file is not None:
    LOGIN_BG_URI = _login_logo_to_b64(_login_logo_file)

LOGIN_BG_IMAGE = _find_background_image()
if LOGIN_BG_IMAGE is None:
    LOGIN_BG_IMAGE = _login_logo_file
LOGIN_BG_IMAGE_URI = ""
if LOGIN_BG_IMAGE and LOGIN_BG_IMAGE.exists():
    LOGIN_BG_IMAGE_URI = base64.b64encode(LOGIN_BG_IMAGE.read_bytes()).decode("utf-8")

_LOGO_TAG_LOGIN = (
    f'<img src="data:image/png;base64,{LOGIN_BG_URI}" alt="CIP" '
    'class="auth-logo-img"/>'
    if LOGIN_BG_URI else ""
)

AUTH_CSS = f"""
<style>
@import url('{FONT_IMPORT}');
*, *::before, *::after {{ font-family: {FONT_FAMILY} !important; }}
.stApp {{
  position: relative; min-height: 100vh;
  background:
    linear-gradient(130deg,rgba(0,0,0,.82),rgba(0,0,0,.72)),
    radial-gradient(circle at 85% 12%,rgba(96,165,250,.18),transparent 46%),
    radial-gradient(circle at 12% 82%,rgba(167,139,250,.14),transparent 44%),
    {"url('data:image/png;base64," + LOGIN_BG_IMAGE_URI + "')," if LOGIN_BG_IMAGE_URI else ""}
    linear-gradient(145deg,#000000,#020617);
  background-size: cover, cover, cover, auto min(88vh, 860px), cover;
  background-position: center, center, center, right 36px center, center;
  background-repeat: no-repeat, no-repeat, no-repeat, no-repeat, no-repeat;
}}
[data-testid="stSidebar"] {{ display: none !important; }}
#MainMenu, footer, header {{ visibility: hidden; }}
div[data-testid="stHeader"], div[data-testid="stToolbar"],
div[data-testid="stDecoration"], div[data-testid="stStatusWidget"] {{ display: none !important; }}

.block-container {{
  position: relative; min-height: 100vh;
  display: flex; align-items: center; justify-content: flex-start;
  padding: 24px !important;
  padding-left: clamp(20px, 3.2vw, 56px) !important;
}}
div[data-testid="stForm"] {{
  width: min(520px, 94vw);
  background: linear-gradient(168deg,rgba(255,255,255,.97),rgba(248,250,252,.96));
  border: 1px solid rgba(191,219,254,.70);
  border-radius: 30px; padding: 38px 34px 28px 34px;
  box-shadow: 0 30px 60px rgba(15,23,42,.18),0 0 0 1px rgba(186,206,240,.24);
  backdrop-filter: blur(12px);
}}
div[data-testid="stForm"]::before {{
  content: "";
  display: block; width: 100px; height: 4px;
  border-radius: 999px;
  background: linear-gradient(90deg,#1d4ed8,#60a5fa,#a78bfa);
  margin: 0 auto 22px auto;
}}
.auth-logo-center {{
  display: flex; justify-content: center; margin-bottom: 18px;
}}
.auth-logo-img {{
  height: 74px;
  max-width: 260px;
  width: auto;
  object-fit: contain;
  filter: drop-shadow(0 6px 14px rgba(15,23,42,.18));
}}
.auth-title {{
  margin: 0 0 4px 0; color: {colors.gray160};
  font-size: clamp(1.6rem, 1.9vw, 2rem);
  font-weight: 900; letter-spacing: -.45px;
  line-height: 1.1; text-align: center;
}}
.auth-sub {{
  color: #475569; margin: 7px 0 20px 0;
  font-size: .90rem; font-weight: 500;
  line-height: 1.55; text-align: center;
}}
div[data-testid="stTextInput"] input,
div[data-testid="stPassword"] input {{
  border-radius: 11px !important;
  border: 1px solid {colors.gray50} !important;
  background: rgba(255,255,255,.92) !important;
  font-size: .93rem !important;
  padding: 10px 13px !important;
  transition: border-color .18s ease, box-shadow .18s ease !important;
}}
div[data-testid="stTextInput"] input:focus,
div[data-testid="stPassword"] input:focus {{
  border-color: #3b82f6 !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.15) !important;
}}
div[data-testid="stForm"] .stButton > button {{
  background: linear-gradient(135deg,#0b1228,#1d4ed8 48%,#7c3aed) !important;
  color: #fff !important;
  border: 1px solid rgba(147,197,253,.42) !important;
  border-radius: 14px !important;
  font-weight: 800 !important; font-size: .95rem !important;
  letter-spacing: .3px; min-height: 48px;
  transition: transform .18s ease, box-shadow .18s ease !important;
  box-shadow: 0 12px 26px rgba(29,78,216,.30), inset 0 1px 0 rgba(255,255,255,.22) !important;
}}
div[data-testid="stForm"] .stButton > button:hover {{
  transform: translateY(-2px) !important;
  box-shadow: 0 16px 32px rgba(29,78,216,.38), inset 0 1px 0 rgba(255,255,255,.24) !important;
}}
.auth-footer {{
  margin-top: 16px; text-align: center;
  font-size: .70rem; color: #94a3b8; font-weight: 500;
}}
  @media (max-width: 1100px) {{
  .stApp {{
    background-size: cover, cover, cover, auto min(72vh, 620px), cover;
    background-position: center, center, center, right -24px center, center;
  }}
}}
@media (max-width: 768px) {{
  .stApp {{
    background-size: cover, cover, cover, 0 0, cover;
  }}
}}
</style>
"""

# =============================================================================
# PANTALLA DE LOGIN
# =============================================================================
if not st.session_state.auth:
    st.markdown(AUTH_CSS, unsafe_allow_html=True)
    with st.form("login_form"):
        if _LOGO_TAG_LOGIN:
            st.markdown(f'<div class="auth-logo-center">{_LOGO_TAG_LOGIN}</div>', unsafe_allow_html=True)
        st.markdown("""
        <h2 class="auth-title">Bienvenido a<br>CIP Analitycs</h2>
        <p class="auth-sub">Ingresa tus credenciales institucionales para continuar</p>
        """, unsafe_allow_html=True)

        usuario = st.text_input("Usuario", placeholder="usuario o correo institucional", key="login_user")
        clave   = st.text_input("Contraseña", type="password", placeholder="••••••••", key="login_pass")
        entrar  = st.form_submit_button("Iniciar sesión", use_container_width=True)

        st.markdown('<div class="auth-footer">Acceso restringido · Solo usuarios autorizados · CIP</div>',
                    unsafe_allow_html=True)

        if entrar:
            user_key = _normalize_user(usuario)
            if user_key in USUARIOS_AUTORIZADOS and USUARIOS_AUTORIZADOS.get(user_key) == clave:
                _login_success(usuario.strip())
            else:
                st.warning("Credenciales inválidas. Verifica tus datos e inténtalo nuevamente.")
    st.stop()

# Aplicar CSS global del dashboard solo cuando hay sesión
st.markdown(FLUENT_CSS, unsafe_allow_html=True)


# =============================================================================
# PLOTLY THEME — con fuente DM Sans
# =============================================================================
def get_fluent_chart_layout():
    return {
        "template": "plotly_white",
        "font": {"family": "DM Sans, Inter, Segoe UI, sans-serif",
                 "size": 11, "color": colors.gray130},
        "title": {"text": "", "font": {"size": 13, "color": colors.gray160,
                  "family": "DM Sans, Inter, Segoe UI, sans-serif"},
                  "x": 0, "xanchor": "left"},
        "paper_bgcolor": "rgba(0,0,0,0)",
        "plot_bgcolor": "rgba(0,0,0,0)",
        "margin": {"l": 38, "r": 15, "t": 52, "b": 36},
        "xaxis": {
            "showgrid": False, "showline": True,
            "linecolor": colors.gray40, "linewidth": 1,
            "tickfont": {"size": 10, "color": colors.gray90},
            "rangeslider": RANGE_SLIDER,
        },
        "yaxis": {
            "showgrid": True, "gridcolor": colors.gray30,
            "showline": False,
            "tickfont": {"size": 10, "color": colors.gray90},
        },
        "colorway": [
            BULLETIN_BLUE, BULLETIN_GOLD, BULLETIN_BLUE_DARK, BULLETIN_BLUE_LIGHT,
            BULLETIN_GOLD_DARK, BULLETIN_SLATE, "#7c3aed", "#0891b2"
        ],
        "legend": {"font": {"size": 10, "color": colors.gray130, "family": "DM Sans, Inter"},
                   "title": {"text": ""}},
    }


def deep_merge(base: dict, override: dict) -> dict:
    out = dict(base)
    for k, v in override.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def apply_fluent_layout(fig: go.Figure, **overrides) -> go.Figure:
    merged = deep_merge(get_fluent_chart_layout(), overrides)
    fig.update_layout(**merged)
    return fig


CHART_CONFIG = {
    "displayModeBar": True,
    "modeBarButtonsToRemove": [
        "toImage", "sendDataToCloud", "editInChartStudio",
        "select2d", "lasso2d", "autoScale2d",
    ],
    "displaylogo": False,
    "scrollZoom": False,
}


def show_plotly(fig: go.Figure, key: str, footnote: str = ""):
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG, key=key)
    if footnote:
        st.markdown(f'<div class="chart-footnote">{footnote}</div>', unsafe_allow_html=True)


def show_panel_footnote():
    """Muestra UNA sola vez la nota de exclusiones al final de cada panel/sección."""
    st.markdown(
        f'<div class="panel-footnote">ⓘ {nota_excl()}</div>',
        unsafe_allow_html=True
    )


# =============================================================================
# COMPONENTES UI
# =============================================================================
def kpi_card(label, value,
             delta_acum=None, delta_acum_type="neutral",
             delta_mes=None,  delta_mes_type="neutral",
             sub=None) -> str:
    deltas_html = '<div class="kpi-deltas">'
    if delta_acum:
        icon = "▲" if delta_acum_type == "positive" else "▼" if delta_acum_type == "negative" else "●"
        deltas_html += (
            f'<span class="kpi-delta {delta_acum_type}">'
            f'{icon} {delta_acum}'
            f'<span style="font-size:7.5px;opacity:.7;margin-left:2px">Acum.</span></span>'
        )
    if delta_mes:
        icon = "▲" if delta_mes_type == "positive" else "▼" if delta_mes_type == "negative" else "●"
        deltas_html += (
            f'<span class="kpi-delta {delta_mes_type}">'
            f'{icon} {delta_mes}'
            f'<span style="font-size:7.5px;opacity:.7;margin-left:2px">Ult. mes</span></span>'
        )
    deltas_html += '</div>'
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return f"""
    <div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      {deltas_html}
      {sub_html}
    </div>
    """


def section_header(title: str) -> None:
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)


def card(title: Optional[str] = None):
    try:
        c = st.container(border=True)
    except TypeError:
        c = st.container()
    with c:
        if title:
            st.markdown(f'<div class="card-title">{title}</div>', unsafe_allow_html=True)
    return c


# =============================================================================
# FORMATEO
# =============================================================================
def fmt_usd_m(x_millones: float) -> str:
    if x_millones is None or (isinstance(x_millones, float) and np.isnan(x_millones)):
        return "—"
    x = float(x_millones)
    s = f"{x:,.1f}".replace(",","_").replace(".",",").replace("_",".")
    return f"USD {s} MM"


def fmt_pct_from_ratio(r: float, with_sign: bool = True) -> str:
    if r is None or (isinstance(r, float) and np.isnan(r)):
        return "—"
    sign = "+" if r > 0 and with_sign else ""
    return f"{sign}{r*100:.1f}%".replace(".", ",")


def fmt_pts(x: float, with_sign: bool = True) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    sign = "+" if x > 0 and with_sign else ""
    return f"{sign}{x:.1f} pts".replace(".", ",")


def fmt_num_latam(x: float, decimals: int = 1) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    txt = f"{float(x):,.{decimals}f}"
    return txt.replace(",","_").replace(".",",").replace("_",".")


def get_delta_type_ratio(value: float) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "neutral"
    return "positive" if value > 0.005 else "negative" if value < -0.005 else "neutral"


# =============================================================================
# UTIL NORMALIZACIÓN
# =============================================================================
def _norm_upper(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return s


# =============================================================================
# RUTAS
# =============================================================================
def resolve_path(primary_windows: str, fallbacks: List[str]) -> Optional[Path]:
    p = Path(primary_windows)
    if p.exists():
        return p
    for fb in fallbacks:
        q = Path(fb)
        if q.exists():
            return q
    return None


BASE_MADRE_FILE = resolve_path(
    r"C:\Users\jcambuludi\OneDrive - CAMARAS DE INDUSTRIA Y PRODUCCION\Downloads\BASE MADRE DASHBOARD VENTAS.xlsx",
    [
        str(BASE_DIR / "BASE MADRE DASHBOARD VENTAS.xlsx"),
        str(BASE_DIR / "data" / "BASE MADRE DASHBOARD VENTAS.xlsx"),
        "/mnt/data/BASE MADRE DASHBOARD VENTAS.xlsx",
        "/mnt/user-data/uploads/BASE MADRE DASHBOARD VENTAS.xlsx",
    ],
)
DATA_DIR = BASE_MADRE_FILE.parent if BASE_MADRE_FILE else BASE_DIR

SRI_RECAUD_FILE = resolve_path(
    r"C:\Users\jcambuludi\OneDrive - CAMARAS DE INDUSTRIA Y PRODUCCION\Downloads\BASE_SRI.xlsx",
    [
        str(BASE_DIR / "BASE_SRI.xlsx"),
        str(BASE_DIR / "data" / "BASE_SRI.xlsx"),
        "/mnt/data/BASE_SRI.xlsx",
        "/mnt/user-data/uploads/BASE_SRI.xlsx",
    ],
)

# =============================================================================
# CONSTANTES — EXCLUSIONES OBLIGATORIAS
# =============================================================================
MANDATORY_EXCLUDE_SECTIONS   = {"B", "D", "E", "O"}
MANDATORY_EXCLUDE_ACTIVITIES = {"G466103", "G4661O3"}


def nota_excl():
    return "Excluye secciones B, D, E y O, y la actividad G466103. Fuente: SRI · Elaboración: DT-CIP"


# =============================================================================
# MEASURE MAP
# =============================================================================
MEASURE_MAP = {
    "ventas_domesticas": "VENTAS DOMÉSTICAS",
    "ventas_totales":    "TOTAL VENTAS Y EXPORTACIONES (419)",
    "export_bienes":     "EXPORTACIONES DE BIENES (417)",
    "export_serv":       "EXPORTACIONES DE SERVICIOS (418)",
    "export_total":      "EXPORTACIONES TOTALES",
    "gravadas_411":      "VENTAS LOCALES GRAVADAS (411)",
    "tarifa_var":        "VENTAS LOCALES TARIFA VARIABLE (420)",
    "gravadas_5":        "VENTAS LOCALES 5% (435)",
}
IVA_MEASURES    = [MEASURE_MAP["gravadas_411"], MEASURE_MAP["tarifa_var"], MEASURE_MAP["gravadas_5"]]
EXPORT_MEASURES = [MEASURE_MAP["export_bienes"], MEASURE_MAP["export_serv"]]

SECTOR_ORDER = ["A","C","F","G","H","I","J","OTR"]
SECTOR_OTHER_SECTIONS = {"M","N","P","Q","R","S","T","U","V","W","X"}
SECTOR_LABEL = {
    "A":"Agricultura","C":"Manufactura","F":"Construcción",
    "G":"Comercio","H":"Transporte","I":"Alojamiento",
    "J":"Información","OTR":"Otros",
}


# =============================================================================
# LOADERS — CACHEADOS
# =============================================================================
def _detect_month_cols(df: pd.DataFrame) -> List[str]:
    rgx = re.compile(r"^(\d{4})/(\d{1,2})")
    return [str(c).strip() for c in df.columns if rgx.match(str(c).strip())]


def _col_to_ts(col: str) -> Optional[pd.Timestamp]:
    m = re.match(r"^(\d{4})/(\d{1,2})", str(col).strip())
    if not m:
        return None
    y, mm = int(m.group(1)), int(m.group(2))
    return pd.Timestamp(y, mm, 1) if 1 <= mm <= 12 else None


@st.cache_data(show_spinner=False, ttl=3600)
def load_ventas_wide(path, mtime):
    df = safe_read_excel(path, sheet_name="VENTAS AGREGADO")
    df.columns = [str(c).strip() for c in df.columns]
    for c in ["SECCIÓN","ACTIVIDAD ECONÓMICA","TIPO CONTRIBUYENTE","MeasuresLevel"]:
        if c not in df.columns:
            raise ValueError(f"Falta columna: {c}")
    df["SECCIÓN"]              = df["SECCIÓN"].astype(str).str.strip().str.upper()
    df["ACTIVIDAD ECONÓMICA"]  = df["ACTIVIDAD ECONÓMICA"].astype(str).str.strip().str.upper()
    df["TIPO CONTRIBUYENTE"]   = df["TIPO CONTRIBUYENTE"].astype(str).str.strip().str.upper()
    df["MeasuresLevel"]        = df["MeasuresLevel"].astype(str).str.strip()
    mc = _detect_month_cols(df)
    if not mc:
        raise ValueError("No se detectaron columnas mensuales.")
    mts = [_col_to_ts(c) for c in mc]
    if any(x is None for x in mts):
        raise ValueError("Columnas mensuales no parseables.")
    mc_s  = [c for _,c in sorted(zip(mts,mc))]
    mts_s = sorted(mts)
    df[mc_s] = df[mc_s].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    return df, mc_s, mts_s

def _detect_geo_value_columns(df: pd.DataFrame) -> Tuple[str, str, int, int]:
    patt = re.compile(r"^(?:TOTAL|VALOR(?:_[A-Z]+)?)_(\d{4})$")
    candidates = []
    for c in df.columns:
        c_norm = _norm_upper(c)
        m = patt.match(c_norm)
        if m:
            candidates.append((int(m.group(1)), c))
    if len(candidates) >= 2:
        candidates = sorted(candidates, key=lambda x: x[0])
        y_prev, c_prev = candidates[-2]
        y_cur,  c_cur  = candidates[-1]
        return c_prev, c_cur, y_prev, y_cur
    raise ValueError(
        "No se detectaron dos columnas de valores en GEOGRAFÍA. "
        "Se esperaban nombres como TOTAL_2025 y TOTAL_2026, o VALOR_ENE_2025 y VALOR_ENE_2026."
    )


@st.cache_data(show_spinner=False, ttl=3600)
def load_geo_cantonal(path, mtime):
    df = safe_read_excel(path, sheet_name="GEOGRAFÍA")
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {}
    for c in df.columns:
        c_norm = _norm_upper(c)
        if c_norm == "CANTON":
            rename_map[c] = "CANTON"
        elif c_norm == "TIPO":
            rename_map[c] = "TIPO"
        elif c_norm in {"CATEGORIA", "CATEGORÍA"}:
            rename_map[c] = "CATEGORÍA"
        elif c_norm == "VARIACION_PORCENTUAL":
            rename_map[c] = "VARIACION_PORCENTUAL"
        elif c_norm == "VARIACION_ABSOLUTA":
            rename_map[c] = "VARIACION_ABSOLUTA"
    if rename_map:
        df = df.rename(columns=rename_map)

    for c in ["CANTON", "TIPO", "CATEGORÍA"]:
        if c not in df.columns:
            raise ValueError(f"Falta columna: {c}")

    prev_col, cur_col, year_prev, year_cur = _detect_geo_value_columns(df)

    df["CANTON"]    = df["CANTON"].astype(str).str.strip()
    df["TIPO"]      = df["TIPO"].astype(str).str.strip().str.upper()
    df["CATEGORÍA"] = df["CATEGORÍA"].astype(str).str.strip().str.upper()

    df["GEO_VALOR_PREV"] = pd.to_numeric(df[prev_col], errors="coerce")
    df["GEO_VALOR_CUR"]  = pd.to_numeric(df[cur_col],  errors="coerce")
    df["GEO_YEAR_PREV"]  = year_prev
    df["GEO_YEAR_CUR"]   = year_cur

    if "VARIACION_ABSOLUTA" in df.columns:
        df["VARIACION_ABSOLUTA"] = pd.to_numeric(df["VARIACION_ABSOLUTA"], errors="coerce")
    else:
        df["VARIACION_ABSOLUTA"] = df["GEO_VALOR_CUR"] - df["GEO_VALOR_PREV"]

    if "VARIACION_PORCENTUAL" in df.columns:
        df["VARIACION_PORCENTUAL"] = pd.to_numeric(df["VARIACION_PORCENTUAL"], errors="coerce")
    else:
        base = pd.to_numeric(df["GEO_VALOR_PREV"], errors="coerce")
        df["VARIACION_PORCENTUAL"] = np.where(base != 0, (df["GEO_VALOR_CUR"] / base - 1.0) * 100.0, np.nan)

    return df


@st.cache_data(show_spinner=False, ttl=3600)
def load_inflacion(path, mtime):
    df = safe_read_excel(path, sheet_name="INFLACIÓN")
    df.columns = [str(c).strip() for c in df.columns]
    df["fecha"] = pd.to_datetime(df.get("Periodo", pd.Series(dtype=str)), errors="coerce")
    for col in ["Consumidor","Productor","Consumo Intermedio"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df.dropna(subset=["fecha"]).sort_values("fecha").reset_index(drop=True)


@st.cache_data(show_spinner=False, ttl=3600)
def load_confianza(path, mtime):
    df = safe_read_excel(path, sheet_name="CONFIANZA")
    df.columns = [str(c).strip() for c in df.columns]
    df["fecha"] = pd.to_datetime(df.get("Periodo", pd.Series(dtype=str)), errors="coerce")
    df = df.dropna(subset=["fecha"]).sort_values("fecha").reset_index(drop=True)
    col_global = "Global" if "Global" in df.columns else None
    col_pres   = next((c for c in ["Presente","Situación Presente (3)"] if c in df.columns), None)
    col_fut    = next((c for c in ["Futuro","Situación Futura (4)"]     if c in df.columns), None)
    return pd.DataFrame({
        "fecha":   df["fecha"],
        "Global":  pd.to_numeric(df[col_global], errors="coerce") if col_global else np.nan,
        "Presente":pd.to_numeric(df[col_pres],   errors="coerce") if col_pres   else np.nan,
        "Futuro":  pd.to_numeric(df[col_fut],    errors="coerce") if col_fut    else np.nan,
    })


@st.cache_data(show_spinner=False, ttl=3600)
def load_indice_expectativas(path, mtime):
    df = safe_read_excel(path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    val_col = next((c for c in df.columns if "expectativas" in _norm_upper(c).lower()), None)
    if val_col is None:
        raise ValueError("No se encontró columna de índice de expectativas en 'indice.xlsx'.")
    out = pd.DataFrame({
        "fecha":               pd.to_datetime(df.get("Periodo", pd.Series(dtype=str)), errors="coerce"),
        "Indice_expectativas": pd.to_numeric(df[val_col], errors="coerce"),
    })
    return out.dropna(subset=["fecha"]).sort_values("fecha").reset_index(drop=True)


def _parse_incidencias_period(value) -> pd.Timestamp:
    txt = str(value).strip().lower()
    txt = re.sub(r"\s*\(p\)\s*", "", txt)
    m = re.match(r"^([a-záéíóú]{3})-(\d{2})$", txt)
    if not m:
        return pd.NaT
    mm_key, yy = m.group(1), int(m.group(2))
    month_map = {
        "ene": 1, "feb": 2, "mar": 3, "abr": 4,
        "may": 5, "jun": 6, "jul": 7, "ago": 8,
        "sep": 9, "oct": 10, "nov": 11, "dic": 12,
    }
    month = month_map.get(mm_key)
    if month is None:
        return pd.NaT
    year = 2000 + yy
    return pd.Timestamp(year=year, month=month, day=1)


@st.cache_data(show_spinner=False, ttl=3600)
def load_incidencias_inpp(path, mtime):
    df = safe_read_excel(path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    id_cols    = ["CPC", "Descripción"]
    month_cols = [c for c in df.columns if c not in id_cols]
    long_df    = df.melt(id_vars=id_cols, value_vars=month_cols,
                         var_name="Periodo", value_name="Incidencia")
    long_df["fecha"] = long_df["Periodo"].apply(_parse_incidencias_period)
    long_df["Incidencia"] = pd.to_numeric(long_df["Incidencia"], errors="coerce")
    long_df["CPC"]        = long_df["CPC"].astype(str).str.strip()
    long_df["Descripción"] = long_df["Descripción"].astype(str).str.strip()
    return (long_df.dropna(subset=["fecha", "Incidencia"])
                   .sort_values(["fecha", "Descripción"])
                   .reset_index(drop=True))


@st.cache_data(show_spinner=False, ttl=3600)
def load_proyecciones(path, mtime):
    df = safe_read_excel(path, sheet_name="PROYECCIONES")
    df.columns = [str(c).strip() for c in df.columns]
    df["fecha"] = pd.to_datetime(df.get("Periodo", pd.Series(dtype=str)), errors="coerce")
    return df.dropna(subset=["fecha"]).sort_values("fecha").reset_index(drop=True)


@st.cache_data(show_spinner=False, ttl=3600)
def load_anios_moviles(path, mtime):
    df = safe_read_excel(path, sheet_name="AÑOS MÓVILES")
    df.columns = [str(c).strip() for c in df.columns]
    df["fecha"] = pd.to_datetime(df.get("Periodo", pd.Series(dtype=str)), errors="coerce")
    return df.dropna(subset=["fecha"]).sort_values("fecha").reset_index(drop=True)


# =============================================================================
# FILTROS OBLIGATORIOS + LÓGICA
# =============================================================================
def _mandatory_mask_ventas(dfw):
    sec = dfw["SECCIÓN"].astype(str).str.strip().str.upper()
    act = dfw["ACTIVIDAD ECONÓMICA"].astype(str).str.strip().str.upper()
    return (~sec.isin(MANDATORY_EXCLUDE_SECTIONS)) & (~act.isin(MANDATORY_EXCLUDE_ACTIVITIES))


def _bucket_sector(sec):
    s = _norm_upper(sec)
    if s in MANDATORY_EXCLUDE_SECTIONS: return None
    if s in {"A","C","F","G","H","I","J"}: return s
    if s in SECTOR_OTHER_SECTIONS: return "OTR"
    return None


def _filter_contrib(df, contrib, col="TIPO CONTRIBUYENTE"):
    c = _norm_upper(contrib)
    if c == "TODOS": return df
    return df[df[col].astype(str).str.strip().str.upper() == c].copy()


def filter_ventas_wide(dfw, contrib, exclude_sections_user, include_commerce, only_commerce):
    df  = dfw.copy()
    m0  = _mandatory_mask_ventas(df)
    df  = _filter_contrib(df, contrib, col="TIPO CONTRIBUYENTE")
    m   = m0.reindex(df.index, fill_value=False)
    exc = set([str(s).strip().upper() for s in (exclude_sections_user or []) if str(s).strip()])
    exc |= MANDATORY_EXCLUDE_SECTIONS
    if only_commerce:
        m = m & (df["SECCIÓN"].astype(str).str.upper() == "G")
    else:
        if not include_commerce: exc.add("G")
        if exc: m = m & (~df["SECCIÓN"].astype(str).str.upper().isin(list(exc)))
    return df.loc[m].copy()


def series_from_wide(df, measure, month_cols, month_ts):
    d = df[df["MeasuresLevel"] == measure]
    if d.empty: return pd.Series(dtype=float)
    vals = d[month_cols].sum(axis=0).astype(float)
    return pd.Series(vals.values, index=pd.Index(month_ts)).sort_index()


def build_ventas_panel_from_wide(dfw_f, month_cols, month_ts):
    s_dom  = series_from_wide(dfw_f, MEASURE_MAP["ventas_domesticas"], month_cols, month_ts)
    s_tot  = series_from_wide(dfw_f, MEASURE_MAP["ventas_totales"],    month_cols, month_ts)
    s_iva  = None
    for m in IVA_MEASURES:
        sm   = series_from_wide(dfw_f, m, month_cols, month_ts)
        s_iva = sm if s_iva is None else s_iva.add(sm, fill_value=0.0)
    s_eb   = series_from_wide(dfw_f, MEASURE_MAP["export_bienes"], month_cols, month_ts)
    s_es   = series_from_wide(dfw_f, MEASURE_MAP["export_serv"],   month_cols, month_ts)
    s_et   = series_from_wide(dfw_f, MEASURE_MAP["export_total"],  month_cols, month_ts)
    if len(s_et) == 0:
        s_et = s_eb.add(s_es, fill_value=0.0)
    idx = pd.Index(month_ts, name="fecha")
    df  = pd.DataFrame(index=idx).reset_index()
    df["ventas_domesticas"] = s_dom.reindex(idx).fillna(0.0).values
    df["ventas_gravadas"]   = (s_iva.reindex(idx).fillna(0.0).values if s_iva is not None else 0.0)
    df["exportaciones"]     = s_et.reindex(idx).fillna(0.0).values
    df["export_bienes"]     = s_eb.reindex(idx).fillna(0.0).values
    df["export_serv"]       = s_es.reindex(idx).fillna(0.0).values
    df["ventas_totales"]    = (s_tot.reindex(idx).fillna(0.0).values
                               if len(s_tot) > 0
                               else df["ventas_domesticas"] + df["exportaciones"])
    return df


def get_last_available_month_from_ts(month_ts):
    return pd.Timestamp(max(month_ts)).replace(day=1)


def calc_ytd_and_last_month(df_ts, last_month, months_use=None, reference_year=None):
    out     = {}
    metrics = ["ventas_domesticas","ventas_gravadas","exportaciones","ventas_totales",
               "export_bienes","export_serv"]
    df      = df_ts.copy()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.to_period("M").dt.to_timestamp()
    df      = df.dropna(subset=["fecha"])
    ytd_end = pd.Timestamp(last_month).to_period("M").to_timestamp()
    year    = int(reference_year) if reference_year is not None else int(ytd_end.year)
    ytd_end = pd.Timestamp(year, int(ytd_end.month), 1)
    months_set = {
        int(m) for m in (months_use if months_use else range(1, int(ytd_end.month)+1))
        if 1 <= int(m) <= int(ytd_end.month)
    } or set(range(1, int(ytd_end.month)+1))
    cur_mask  = (df["fecha"].dt.year == year)     & (df["fecha"].dt.month.isin(months_set))
    prev_mask = (df["fecha"].dt.year == year-1)   & (df["fecha"].dt.month.isin(months_set))
    last_cur  = ytd_end
    last_prev = last_cur - pd.DateOffset(years=1)
    for m in metrics:
        if m not in df.columns: continue
        ytd_c  = float(df.loc[cur_mask,  m].sum())
        ytd_p  = float(df.loc[prev_mask, m].sum())
        ytd_yoy = (ytd_c / ytd_p - 1.0) if ytd_p > 0 else np.nan
        lc = float(df.loc[df["fecha"] == last_cur,  m].sum())
        lp = float(df.loc[df["fecha"] == last_prev, m].sum())
        last_yoy = (lc / lp - 1.0) if lp > 0 else np.nan
        out[m] = dict(ytd_cur=ytd_c, ytd_prev=ytd_p, ytd_yoy=ytd_yoy,
                      last_cur=lc, last_prev=lp, last_yoy=last_yoy)
    return out


# =============================================================================
# SECTORES
# =============================================================================
def _months_effective(months_sel, end_month):
    if not months_sel: return list(range(1, end_month+1))
    return sorted([m for m in months_sel if 1 <= int(m) <= int(end_month)])


def _month_cols_for_year(month_cols, month_ts, year, months_use):
    return [c for c,ts in zip(month_cols, month_ts)
            if int(ts.year) == int(year) and int(ts.month) in set(months_use)]


def sector_year_sum_wide(dfw_all, month_cols, month_ts, contrib, measures,
                          year, months_use, include_commerce, only_commerce, exclude_sections_user):
    df    = filter_ventas_wide(dfw_all, contrib, exclude_sections_user, include_commerce, only_commerce)
    df    = df[df["MeasuresLevel"].isin(measures)].copy()
    cy    = _month_cols_for_year(month_cols, month_ts, year, months_use)
    if not cy:
        return pd.Series({k: 0.0 for k in SECTOR_ORDER}, dtype=float)
    df["bucket"] = df["SECCIÓN"].apply(_bucket_sector)
    df = df.dropna(subset=["bucket"])
    agg = df.groupby("bucket")[cy].sum().sum(axis=1)
    out = pd.Series({k: 0.0 for k in SECTOR_ORDER}, dtype=float)
    for k, v in agg.items():
        if str(k) in out.index:
            out[str(k)] = float(v) if pd.notna(v) else 0.0
    return out


def build_sector_panel_wide(dfw_all, month_cols, month_ts, contrib, measures,
                              year_cur, end_month, months_sel,
                              include_commerce, only_commerce, exclude_sections_user):
    months_use = _months_effective(months_sel, end_month)
    y, y1, y2  = int(year_cur), year_cur-1, year_cur-2
    s_y  = sector_year_sum_wide(dfw_all, month_cols, month_ts, contrib, measures, y,  months_use, include_commerce, only_commerce, exclude_sections_user)
    s_y1 = sector_year_sum_wide(dfw_all, month_cols, month_ts, contrib, measures, y1, months_use, include_commerce, only_commerce, exclude_sections_user)
    s_y2 = sector_year_sum_wide(dfw_all, month_cols, month_ts, contrib, measures, y2, months_use, include_commerce, only_commerce, exclude_sections_user)
    yoy_y1 = np.where(s_y2.values > 0, (s_y1.values/s_y2.values - 1.0)*100.0, np.nan)
    yoy_y  = np.where(s_y1.values > 0, (s_y.values/s_y1.values  - 1.0)*100.0, np.nan)
    return pd.DataFrame({
        "sector_code": SECTOR_ORDER,
        "sector": [f"{SECTOR_LABEL[c]} ({c})" if c!="OTR" else "Otros (OTR)" for c in SECTOR_ORDER],
        f"Nivel_{y2}_MM": s_y2.values / 1e6,
        f"Nivel_{y1}_MM": s_y1.values / 1e6,
        f"Nivel_{y}_MM":  s_y.values / 1e6,
        f"YoY_{y1}": yoy_y1,
        f"YoY_{y}":  yoy_y,
        "Abs_MM":    (s_y.values - s_y1.values) / 1e6,
        "Nivel_MM":  s_y.values / 1e6,
    })


def _compute_axis_range(series_list, *, min_pad=0.0, pct_pad=0.12, fallback=(-1.0, 1.0)):
    vals = []
    for s in series_list:
        arr = pd.to_numeric(pd.Series(s), errors="coerce").dropna().tolist()
        vals.extend(arr)
    if not vals:
        return list(fallback)
    vmin, vmax = float(min(vals)), float(max(vals))
    if np.isclose(vmin, vmax):
        pad = max(abs(vmax) * pct_pad, min_pad, 1.0)
        return [vmin - pad, vmax + pad]
    span = vmax - vmin
    pad = max(span * pct_pad, min_pad)
    return [vmin - pad, vmax + pad]


def plot_sector_relative(df, y_prev, y_cur, title, key, footnote=""):
    col_prev = f"YoY_{y_prev}"
    col_cur  = f"YoY_{y_cur}"
    lvl_prev = f"Nivel_{y_prev}_MM"
    lvl_cur  = f"Nivel_{y_cur}_MM"
    lvl_base = f"Nivel_{y_prev-1}_MM"

    y_prev_vals = pd.to_numeric(df.get(col_prev, np.nan), errors="coerce")
    y_cur_vals  = pd.to_numeric(df.get(col_cur,  np.nan), errors="coerce")

    series_for_axis = []
    if y_prev_vals.notna().any():
        series_for_axis.append(y_prev_vals)
    if y_cur_vals.notna().any():
        series_for_axis.append(y_cur_vals)

    y_range = _compute_axis_range(
        series_for_axis,
        min_pad=6.0,
        pct_pad=0.18,
        fallback=(-10.0, 10.0)
    )

    fig = go.Figure()

    if y_prev_vals.notna().any():
        fig.add_trace(go.Scatter(
            x=df["sector"], y=y_prev_vals, mode="markers+text", name=str(y_prev),
            text=[("" if pd.isna(v) else f"{v:.1f}%") for v in y_prev_vals],
            textposition="bottom center",
            marker=dict(size=10, color=BULLETIN_BLUE, line=dict(width=1.2, color="#ffffff")),
            cliponaxis=False,
            customdata=np.column_stack([
                pd.to_numeric(df.get(lvl_prev, np.nan), errors="coerce"),
                pd.to_numeric(df.get(lvl_base, np.nan), errors="coerce"),
            ]),
            hovertemplate=(
                "<b>%{x}</b>"
                f"<br>{y_prev} vs {y_prev-1}: %{{y:.2f}}%"
                f"<br>{y_prev}: %{{customdata[0]:,.2f}} MM"
                f"<br>{y_prev-1}: %{{customdata[1]:,.2f}} MM"
                "<extra></extra>"
            )
        ))

    if y_cur_vals.notna().any():
        fig.add_trace(go.Scatter(
            x=df["sector"], y=y_cur_vals, mode="markers+text", name=str(y_cur),
            text=[("" if pd.isna(v) else f"{v:.1f}%") for v in y_cur_vals],
            textposition="top center",
            marker=dict(size=10, color=BULLETIN_GOLD, line=dict(width=1.2, color="#ffffff")),
            cliponaxis=False,
            customdata=np.column_stack([
                pd.to_numeric(df.get(lvl_cur, np.nan), errors="coerce"),
                pd.to_numeric(df.get(lvl_prev, np.nan), errors="coerce"),
            ]),
            hovertemplate=(
                "<b>%{x}</b>"
                f"<br>{y_cur} vs {y_prev}: %{{y:.2f}}%"
                f"<br>{y_cur}: %{{customdata[0]:,.2f}} MM"
                f"<br>{y_prev}: %{{customdata[1]:,.2f}} MM"
                "<extra></extra>"
            )
        ))

    fig.add_hline(y=0, line_width=1, line_dash="solid", opacity=0.35)
    apply_fluent_layout(
        fig, height=390, hovermode="x unified",
        yaxis=dict(title="Variación interanual (%)", range=y_range),
        xaxis=dict(tickangle=90, rangeslider=RANGE_SLIDER),
        legend=dict(orientation="h", yanchor="bottom", y=1.04, xanchor="center", x=0.5),
        margin=dict(l=38, r=15, t=80, b=90),
        title=dict(text=title, x=0.5, xanchor="center", font=dict(size=12)),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor=colors.gray30, ticksuffix="%")
    show_plotly(fig, key, footnote=footnote)


def plot_sector_absolute(df, title, key, sort_asc=True, footnote=""):
    if sort_asc:
        dfb = df.sort_values("Abs_MM", ascending=True, kind="stable").copy()
    else:
        dfb = df.copy()

    abs_vals = pd.to_numeric(dfb["Abs_MM"], errors="coerce")
    y_range = _compute_axis_range([abs_vals], min_pad=20.0, pct_pad=0.16, fallback=(-50.0, 50.0))
    bar_colors = [
        BULLETIN_GOLD_DARK if (pd.notna(v) and v >= 0) else BULLETIN_RED_SOFT
        for v in abs_vals
    ]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=dfb["sector"], y=abs_vals,
        text=[("" if pd.isna(v) else fmt_num_latam(v,1)) for v in abs_vals],
        textposition="outside",
        marker=dict(color=bar_colors, line=dict(color="#ffffff", width=1.2)),
        name="Δ (MM)",
        cliponaxis=False,
        hovertemplate=(
            "<b>%{x}</b>"
            "<br>Variación absoluta: %{y:,.2f} MM"
            "<extra></extra>"
        )
    ))
    fig.add_hline(y=0, line_width=1, line_dash="solid", opacity=0.35)
    apply_fluent_layout(
        fig, height=390,
        yaxis=dict(title="Variación absoluta (USD millones)", range=y_range),
        xaxis=dict(tickangle=90, rangeslider=RANGE_SLIDER),
        showlegend=False,
        title=dict(text=title, x=0, font=dict(size=12)),
        margin=dict(l=48, r=15, t=55, b=95)
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor=colors.gray30)
    show_plotly(fig, key, footnote=footnote)

@st.cache_data(show_spinner=False, ttl=3600)
def load_sri_principal(path, mtime):
    df = safe_read_excel(path, sheet_name="PRINCIPAL")
    df.columns = [str(c).strip() for c in df.columns]

    req = ["TIPO CONTRIBUYENTE", "GRUPO", "MeasuresLevel"]
    for c in req:
        if c not in df.columns:
            raise ValueError(f"Falta columna en PRINCIPAL: {c}")

    df["TIPO CONTRIBUYENTE"] = df["TIPO CONTRIBUYENTE"].astype(str).str.strip().str.upper()
    df["GRUPO"] = df["GRUPO"].astype(str).str.strip().str.upper()
    df["MeasuresLevel"] = df["MeasuresLevel"].astype(str).str.strip().str.upper()

    mc = _detect_month_cols(df)
    if not mc:
        raise ValueError("No se detectaron columnas mensuales en PRINCIPAL.")

    mts = [_col_to_ts(c) for c in mc]
    if any(x is None for x in mts):
        raise ValueError("Hay columnas mensuales no parseables en PRINCIPAL.")

    mc_s = [c for _, c in sorted(zip(mts, mc))]
    mts_s = sorted(mts)

    df[mc_s] = df[mc_s].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    return df, mc_s, mts_s


@st.cache_data(show_spinner=False, ttl=3600)
def load_sri_secciones(path, mtime):
    df = safe_read_excel(path, sheet_name="SECCIONES")

    if "SECCIONES" not in df.columns:
        raise ValueError("No se encontró la columna SECCIONES en la hoja SECCIONES.")

    df["SECCIONES"] = df["SECCIONES"].astype(str).str.strip().str.upper()

    date_cols = [c for c in df.columns if hasattr(c, "year") and hasattr(c, "month")]
    if not date_cols:
        raise ValueError("No se detectaron columnas de fecha en SECCIONES.")

    date_cols = sorted(date_cols)
    df[date_cols] = df[date_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    return df, date_cols


def _filter_recaud_contrib(df, contrib):
    c = _norm_upper(contrib)
    if c in {"TODOS", "CONSOLIDADO", "TOTAL"}:
        return df.copy()
    return df[df["TIPO CONTRIBUYENTE"].astype(str).str.upper() == c].copy()


def sri_measure_series(df, month_cols, month_ts, measure, contrib="Todos", grupos=None):
    d = _filter_recaud_contrib(df, contrib)
    d = d[d["MeasuresLevel"] == str(measure).strip().upper()].copy()

    if grupos:
        gs = {str(g).strip().upper() for g in grupos}
        d = d[d["GRUPO"].isin(gs)].copy()

    vals = d[month_cols].sum(axis=0).astype(float)
    return pd.Series(vals.values, index=pd.Index(month_ts)).sort_index()


def build_recaudacion_series(df, month_cols, month_ts, contrib="Todos", grupos=None):
    s_bruta = sri_measure_series(df, month_cols, month_ts, "VALOR RECAUDADO", contrib, grupos)
    s_nc    = sri_measure_series(df, month_cols, month_ts, "VALOR NOTAS CRÉDITO", contrib, grupos)
    s_comp  = sri_measure_series(df, month_cols, month_ts, "VALOR COMPENSACIONES", contrib, grupos)
    s_tbc   = sri_measure_series(df, month_cols, month_ts, "VALOR TBC", contrib, grupos)

    s_neta = (
        s_bruta
        .subtract(s_nc, fill_value=0.0)
        .subtract(s_comp, fill_value=0.0)
        .subtract(s_tbc, fill_value=0.0)
    )

    idx = pd.Index(month_ts, name="fecha")
    out = pd.DataFrame(index=idx).reset_index()
    out["recaud_bruta"] = s_bruta.reindex(idx).fillna(0.0).values
    out["recaud_neta"] = s_neta.reindex(idx).fillna(0.0).values
    out["notas_credito"] = s_nc.reindex(idx).fillna(0.0).values
    out["compensaciones"] = s_comp.reindex(idx).fillna(0.0).values
    out["tbc"] = s_tbc.reindex(idx).fillna(0.0).values
    return out


def recaud_tax_snapshot(df, month_cols, month_ts, contrib, year_cur, month_cur, months_use):
    year_prev = int(year_cur) - 1
    rows = []

    for label, grupo in RECAUD_TAX_GROUPS.items():
        s = build_recaudacion_series(df, month_cols, month_ts, contrib=contrib, grupos=[grupo])

        cur_ytd = float(s.loc[(s["fecha"].dt.year == year_cur) & (s["fecha"].dt.month.isin(months_use)), "recaud_bruta"].sum())
        prev_ytd = float(s.loc[(s["fecha"].dt.year == year_prev) & (s["fecha"].dt.month.isin(months_use)), "recaud_bruta"].sum())

        cur_m = float(s.loc[(s["fecha"].dt.year == year_cur) & (s["fecha"].dt.month == month_cur), "recaud_bruta"].sum())
        prev_m = float(s.loc[(s["fecha"].dt.year == year_prev) & (s["fecha"].dt.month == month_cur), "recaud_bruta"].sum())

        rows.append({
            "impuesto": label,
            "ytd_prev_mm": prev_ytd / 1e6,
            "ytd_cur_mm": cur_ytd / 1e6,
            "ytd_abs_mm": (cur_ytd - prev_ytd) / 1e6,
            "ytd_var": (cur_ytd / prev_ytd - 1.0) if prev_ytd > 0 else np.nan,
            "mes_prev_mm": prev_m / 1e6,
            "mes_cur_mm": cur_m / 1e6,
            "mes_abs_mm": (cur_m - prev_m) / 1e6,
            "mes_var": (cur_m / prev_m - 1.0) if prev_m > 0 else np.nan,
        })

    return pd.DataFrame(rows)


def build_sector_panel_sri(df_sec, year_cur, months_use):
    year_cur = int(year_cur)
    year_prev = year_cur - 1
    year_base = year_cur - 2

    date_cols = [c for c in df_sec.columns if hasattr(c, "year") and hasattr(c, "month")]

    def _sum_for_year(y):
        cols = [c for c in date_cols if int(c.year) == int(y) and int(c.month) in set(months_use)]
        if not cols:
            return pd.Series(0.0, index=df_sec.index, dtype=float)
        return df_sec[cols].sum(axis=1).astype(float)

    tmp = df_sec.copy()
    tmp["bucket"] = tmp["SECCIONES"].apply(_bucket_sector)
    tmp = tmp.dropna(subset=["bucket"]).copy()

    year_vectors = {
        year_base: _sum_for_year(year_base),
        year_prev: _sum_for_year(year_prev),
        year_cur:  _sum_for_year(year_cur),
    }
    year_aggs = {
        y: tmp.assign(_value=vec).groupby("bucket")["_value"].sum()
        for y, vec in year_vectors.items()
    }

    out = pd.DataFrame({"sector_code": SECTOR_ORDER})
    out["sector"] = [f"{SECTOR_LABEL[c]} ({c})" if c != "OTR" else "Otros (OTR)" for c in SECTOR_ORDER]

    for y in [year_base, year_prev, year_cur]:
        out[f"Nivel_{y}_MM"] = [float(year_aggs[y].get(k, 0.0)) / 1e6 for k in SECTOR_ORDER]

    out[f"YoY_{year_prev}"] = [
        ((float(year_aggs[year_prev].get(k, 0.0)) / float(year_aggs[year_base].get(k, 0.0))) - 1.0) * 100.0
        if float(year_aggs[year_base].get(k, 0.0)) > 0 else np.nan
        for k in SECTOR_ORDER
    ]
    out[f"YoY_{year_cur}"] = [
        ((float(year_aggs[year_cur].get(k, 0.0)) / float(year_aggs[year_prev].get(k, 0.0))) - 1.0) * 100.0
        if float(year_aggs[year_prev].get(k, 0.0)) > 0 else np.nan
        for k in SECTOR_ORDER
    ]
    out["Nivel_MM"] = out[f"Nivel_{year_cur}_MM"]
    out["Abs_MM"] = out[f"Nivel_{year_cur}_MM"] - out[f"Nivel_{year_prev}_MM"]
    return out


def plot_tax_grouped_hbars(df_tax, prev_col, cur_col, var_col, title, key, year_prev, year_cur):
    d = df_tax.sort_values(cur_col, ascending=True).copy()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=d["impuesto"],
        x=d[prev_col],
        name=str(year_prev),
        orientation="h",
        marker_color=BULLETIN_PREVIOUS,
        text=[fmt_num_latam(v, 0) for v in d[prev_col]],
        textposition="inside",
    ))
    fig.add_trace(go.Bar(
        y=d["impuesto"],
        x=d[cur_col],
        name=str(year_cur),
        orientation="h",
        marker_color=BULLETIN_CURRENT,
        text=[
            f"{fmt_num_latam(v, 0)} · {x*100:.0f}%"
            if pd.notna(x) else fmt_num_latam(v, 0)
            for v, x in zip(d[cur_col], d[var_col])
        ],
        textposition="outside",
    ))

    apply_fluent_layout(
        fig,
        height=360,
        barmode="group",
        xaxis=dict(title="USD millones"),
        yaxis=dict(title=""),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(l=40, r=15, t=55, b=30),
        title=dict(text=title, x=0, font=dict(size=12)),
    )
    show_plotly(fig, key)


def plot_recaud_comparison(cur_value, prev_value, cur_label, prev_label, title, key, color_cur, color_prev=BULLETIN_PREVIOUS):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=[prev_label, cur_label],
        y=[prev_value, cur_value],
        marker_color=[color_prev, color_cur],
        text=[fmt_num_latam(prev_value, 0), fmt_num_latam(cur_value, 0)],
        textposition="inside",
        name=title,
    ))
    apply_fluent_layout(
        fig,
        height=320,
        showlegend=False,
        yaxis=dict(title="USD millones"),
        margin=dict(l=38, r=15, t=50, b=36),
        title=dict(text=title, x=0, font=dict(size=12)),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor=colors.gray30)
    show_plotly(fig, key)

def plot_sector_absolute(df, title, key, sort_asc=True, footnote=""):
    if sort_asc:
        dfb = df.sort_values("Abs_MM", ascending=True, kind="stable").copy()
    else:
        dfb = df.copy()

    abs_vals = pd.to_numeric(dfb["Abs_MM"], errors="coerce")
    y_range = _compute_axis_range([abs_vals], min_pad=20.0, pct_pad=0.16, fallback=(-50.0, 50.0))
    bar_colors = [
        BULLETIN_GOLD_DARK if (pd.notna(v) and v >= 0) else BULLETIN_RED_SOFT
        for v in abs_vals
    ]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=dfb["sector"], y=abs_vals,
        text=[("" if pd.isna(v) else fmt_num_latam(v,1)) for v in abs_vals],
        textposition="outside",
        marker=dict(color=bar_colors, line=dict(color="#ffffff", width=1.2)),
        name="Δ (MM)",
        cliponaxis=False,
        hovertemplate=(
            "<b>%{x}</b>"
            "<br>Variación absoluta: %{y:,.2f} MM"
            "<extra></extra>"
        )
    ))
    fig.add_hline(y=0, line_width=1, line_dash="solid", opacity=0.35)
    apply_fluent_layout(
        fig, height=390,
        yaxis=dict(title="Variación absoluta (USD millones)", range=y_range),
        xaxis=dict(tickangle=90, rangeslider=RANGE_SLIDER),
        showlegend=False,
        title=dict(text=title, x=0, font=dict(size=12)),
        margin=dict(l=48, r=15, t=55, b=95)
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor=colors.gray30)
    show_plotly(fig, key, footnote=footnote)


# =============================================================================
# RECAUDACIÓN — helpers
# =============================================================================
def _filter_contrib_recaud(df, contrib):
    c = _norm_upper(contrib)
    if c == "TODOS": return df.copy()
    return df[df["TIPO CONTRIBUYENTE"].astype(str).str.upper() == c].copy()


# =============================================================================
# MAPA CANTONAL
# =============================================================================
MAP_SCALE_YOY = [[0.0,"lightsalmon"],[0.5,"whitesmoke"],[1.0,"royalblue"]]
MAP_SCALE_NIVEL = [[0.0,"whitesmoke"],[0.35,"#A8C4E8"],[0.65,"#4169E1"],[1.0,"#1A2F6B"]]
MAP_COLOR_MISSING = "lightgrey"


def construir_escala_cero(vmin, vmax,
                          neutral="whitesmoke", neg="lightsalmon", pos="royalblue"):
    if vmax <= vmin:
        return [[0.0, neutral],[1.0, pos]]
    if vmin < 0 < vmax:
        p = float(np.clip((0.0-vmin)/(vmax-vmin), 0.0, 1.0))
        return [[0.0,neg],[p,neutral],[1.0,pos]]
    if vmin >= 0: return [[0.0,neutral],[1.0,pos]]
    return [[0.0,neg],[1.0,neutral]]


def find_shp_file(base_dir, preferred_type="canton"):
    if preferred_type == "provincia":
        prefs = ["nxprovincias","provincias","provincia","ecu_provincias","nx_provincias"]
    else:
        prefs = ["nxcantones","cantones","canton","nx_cantones","ecu_cantones"]
    for stem in prefs:
        shp = base_dir / f"{stem}.shp"
        if shp.exists(): return shp
    hits = sorted(list(base_dir.glob("*.shp")))
    return hits[0] if hits else None


def canon_canton_key(s):
    s = _norm_upper(s).replace("Ð","N")
    s = re.sub(r"[^\w\s]"," ",s).replace("_"," ")
    s = re.sub(r"\s+"," ",s).strip()
    for art in ["EL ","LA ","LOS ","LAS "]:
        if s.startswith(art): s = s[len(art):]; break
    for a,b in [("CORONEL","CRNEL"),("GENERAL","GNRAL")]:
        s = s.replace(a,b)
    return s.replace(" ","")


@st.cache_data(show_spinner=False, ttl=7200)
def load_canton_geojson_cached(shp_path, mtime, simplify_tol=0.003):
    if not HAS_GEOPANDAS: return None
    try:
        gdf = gpd.read_file(shp_path)
        if gdf.crs is not None and gdf.crs.to_epsg() != 4326:
            gdf = gdf.to_crs(epsg=4326)
        name_cols = ["DPA_DESCAN","CANTON","DPA_CANTON","NOMBRE","NOM_CANTON"]
        name_col  = next((c for c in name_cols if c in gdf.columns), None)
        if name_col is None:
            name_col = next((c for c in gdf.columns
                             if any(k in c.upper() for k in ["CANTON","DESCAN","NOMBRE"])), None)
        if name_col is None: return None
        gdf = gdf[[name_col,"geometry"]].copy()
        gdf["CANTON_NORM"] = gdf[name_col].apply(lambda x: canon_canton_key(str(x)))
        gdf = gdf.dropna(subset=["CANTON_NORM"]).drop_duplicates(subset=["CANTON_NORM"]).reset_index(drop=True)
        if not gdf.geometry.is_valid.all():
            gdf.geometry = gdf.geometry.buffer(0)
        if simplify_tol:
            gdf["geometry"] = gdf["geometry"].simplify(simplify_tol, preserve_topology=True)
        return gdf.__geo_interface__
    except Exception:
        return None


@st.cache_data(show_spinner=False, ttl=7200)
def load_province_canton_mapping(shp_path, mtime):
    if not HAS_GEOPANDAS: return {}
    try:
        gdf = gpd.read_file(shp_path)
        if gdf.crs is not None and gdf.crs.to_epsg() != 4326:
            gdf = gdf.to_crs(epsg=4326)
        prov_col   = next((c for c in ["DPA_DESPRO","PROVINCIA","DPA_PROV","NOM_PROV","PROVINCE"] if c in gdf.columns), None)
        canton_col = next((c for c in ["DPA_DESCAN","CANTON","DPA_CANTON","NOMBRE","NOM_CANTON"]   if c in gdf.columns), None)
        if canton_col is None: return {}
        gdf["CANTON_NORM"]  = gdf[canton_col].apply(lambda x: canon_canton_key(str(x)))
        canton_centroid = {}
        for _,row in gdf.iterrows():
            try:
                c = row.geometry.centroid
                canton_centroid[row["CANTON_NORM"]] = (float(c.x), float(c.y))
            except Exception: pass
        if prov_col is None:
            return {"prov_cantons":{},"canton_centroid":canton_centroid,"prov_centroid":{},"prov_display":{}}
        prov_cantons: Dict[str, List[str]] = {}
        prov_centroid: Dict[str, tuple]    = {}
        prov_display: Dict[str, str]       = {}
        gdf["PROV_NORM"]    = gdf[prov_col].apply(lambda x: canon_canton_key(str(x)))
        gdf["PROV_DISPLAY"] = gdf[prov_col].apply(lambda x: str(x).strip().title())
        for prov_norm, grp in gdf.groupby("PROV_NORM"):
            prov_cantons[prov_norm] = list(grp["CANTON_NORM"].unique())
            try:
                import shapely.ops
                merged = shapely.ops.unary_union(grp.geometry)
                c      = merged.centroid
                prov_centroid[prov_norm] = (float(c.x), float(c.y))
            except Exception:
                lons = [r.geometry.centroid.x for _,r in grp.iterrows()]
                lats = [r.geometry.centroid.y for _,r in grp.iterrows()]
                prov_centroid[prov_norm] = (float(np.mean(lons)), float(np.mean(lats)))
            prov_display[prov_norm] = grp["PROV_DISPLAY"].iloc[0]
        return {"prov_cantons":prov_cantons,"canton_centroid":canton_centroid,
                "prov_centroid":prov_centroid,"prov_display":prov_display}
    except Exception:
        return {}


@st.cache_data(show_spinner=False, ttl=7200)
def load_prov_outline_traces(shp_path, mtime, simplify_tol=0.0015):
    if not HAS_GEOPANDAS: return []
    try:
        gdf = gpd.read_file(shp_path)
        if gdf.crs is not None and gdf.crs.to_epsg() != 4326:
            gdf = gdf.to_crs(epsg=4326)
        if not gdf.geometry.is_valid.all():
            gdf.geometry = gdf.geometry.buffer(0)
        if simplify_tol:
            gdf["geometry"] = gdf["geometry"].simplify(simplify_tol, preserve_topology=True)
        prov_geojson = gdf.__geo_interface__
        traces = []
        for feat in prov_geojson.get("features",[]):
            geom = feat.get("geometry",{})
            rings = []
            if geom.get("type") == "Polygon":
                rings = geom.get("coordinates",[])
            elif geom.get("type") == "MultiPolygon":
                rings = [ring for poly in geom.get("coordinates",[]) for ring in poly]
            for ring in rings:
                traces.append(go.Scattermapbox(
                    lon=[c[0] for c in ring], lat=[c[1] for c in ring],
                    mode="lines", line=dict(width=1.4, color="rgba(50,49,48,0.82)"),
                    showlegend=False, hoverinfo="skip"
                ))
        return traces
    except Exception:
        return []


def build_cantonal_map_figure(geojson_obj, df_map, metric_label, colorscale, zmid=None, height=520, value_year_label="Valor actual"):
    df_map = df_map.copy()

    # Resolver nombre visible del cantón de forma robusta
    display_col = next(
        (c for c in ["cant_display", "Cantón", "CANTON", "Canton", "canton"] if c in df_map.columns),
        None
    )

    if display_col is not None:
        df_map["cant_display"] = df_map[display_col]
    else:
        df_map["cant_display"] = df_map["CANTON_NORM"]

    df_map["cant_display"] = df_map["cant_display"].astype(str).replace("nan", np.nan)
    df_map["cant_display"] = df_map["cant_display"].fillna(df_map["CANTON_NORM"])

    # Asegurar columnas numéricas requeridas
    for col in ["color_val", "rank", "valor_2025", "share", "yoy"]:
        if col not in df_map.columns:
            df_map[col] = np.nan
        df_map[col] = pd.to_numeric(df_map[col], errors="coerce")

    df_na = df_map[df_map["color_val"].isna()].copy()
    df_ok = df_map[df_map["color_val"].notna()].copy()

    fig = go.Figure()

    if len(df_na):
        fig.add_trace(go.Choroplethmapbox(
            geojson=geojson_obj,
            locations=df_na["CANTON_NORM"],
            z=np.zeros(len(df_na)),
            featureidkey="properties.CANTON_NORM",
            colorscale=[[0, MAP_COLOR_MISSING], [1, MAP_COLOR_MISSING]],
            showscale=False,
            marker_opacity=0.65,
            marker_line_width=0.25,
            customdata=np.column_stack([
                df_na["cant_display"].values
            ]),
            hovertemplate="<b>%{customdata[0]}</b><br>Sin dato<extra></extra>",
            name="Sin dato"
        ))

    if len(df_ok):
        z = df_ok["color_val"].astype(float).values
        z_fin = pd.to_numeric(df_ok["color_val"], errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()

        if len(z_fin) >= 10:
            lo = float(np.nanpercentile(z_fin, 15))
            hi = float(np.nanpercentile(z_fin, 85))
        else:
            lo = float(np.nanmin(z_fin)) if len(z_fin) else 0.0
            hi = float(np.nanmax(z_fin)) if len(z_fin) else 1.0

        if (z_fin == 0).any() and lo > 0:
            lo = 0.0
        if (z_fin == 0).any() and hi < 0:
            hi = 0.0
        if lo == hi:
            hi = lo + 1.0

        map_cs = construir_escala_cero(lo, hi) if zmid == 0.0 else colorscale
        map_zmid = None if zmid == 0.0 else zmid

        fig.add_trace(go.Choroplethmapbox(
            geojson=geojson_obj,
            locations=df_ok["CANTON_NORM"],
            z=z,
            featureidkey="properties.CANTON_NORM",
            colorscale=map_cs,
            zmin=lo,
            zmax=hi,
            zmid=map_zmid,
            marker_opacity=0.75,
            marker_line_width=0.25,
            customdata=np.column_stack([
                df_ok["cant_display"].values,
                df_ok["rank"].values,
                df_ok["valor_2025"].values,
                df_ok["share"].values,
                df_ok["yoy"].values
            ]),
            hovertemplate=(
                "<b>%{customdata[0]}</b>"
                "<br>Posición: %{customdata[1]:.0f}"
                f"<br>{value_year_label}: %{{customdata[2]:,.0f}}"
                "<br>Participación: %{customdata[3]:.2%}"
                "<br>Variación: %{customdata[4]:.2f}%"
                "<extra></extra>"
            ),
            colorbar=dict(title=metric_label, thickness=14, len=0.5),
            name=metric_label
        ))

    fig.update_layout(
        height=height,
        margin=dict(l=0, r=0, t=0, b=0),
        mapbox=dict(style="carto-positron", zoom=5.5, center=dict(lat=-1.5, lon=-78.5)),
        uirevision="geo-map"
    )
    return fig

# =============================================================================
# HEADER / HERO
# =============================================================================
def img_to_base64(path):
    with open(path,"rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def logo_to_base64_without_black(path):
    if Image is None: return img_to_base64(str(path))
    try:
        img  = Image.open(path).convert("RGBA")
        data = np.array(img)
        rgb  = data[:,:,:3]
        mask = (rgb[:,:,0] < 38) & (rgb[:,:,1] < 38) & (rgb[:,:,2] < 38)
        data[:,:,3][mask] = 0
        from io import BytesIO
        buff = BytesIO()
        Image.fromarray(data,"RGBA").save(buff, format="png")
        return base64.b64encode(buff.getvalue()).decode("utf-8")
    except Exception:
        return img_to_base64(str(path))


def _find_cip_logo():
    candidates = [BASE_DIR/"cip.png", BASE_DIR/"CIP.png", BASE_DIR/"cip.png",
                  BASE_DIR/"assets"/"cip.png", BASE_DIR/"img"/"cip.png"]
    for p in candidates:
        if p.exists(): return p
    for p in BASE_DIR.glob("*.png"):
        if "cip" in p.name.lower(): return p
    return None


def get_last_month_label(dt):
    meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    ts    = pd.to_datetime(dt, errors="coerce")
    if pd.isna(ts): return "—"
    m = int(ts.month)
    return f"{meses[m-1]} {ts.year}" if 1 <= m <= 12 else str(ts.year)


hero_slot = st.empty()


def render_hero(subtitle):
    logo = _find_cip_logo()
    logo_html = (
        f'<div class="hero-logo"><img src="data:image/png;base64,'
        f'{logo_to_base64_without_black(logo)}" alt="CIP"/></div>'
        if logo
        else f'<div class="hero-logo"><span style="font-size:12px;color:{colors.gray90};font-weight:700;">CIP</span></div>'
    )
    hero_slot.markdown(f"""
    <div class="hero">
      <div class="hero-wrap">
        <div class="hero-copy">
          <div class="hero-kicker">Dashboard ejecutivo</div>
          <div class="hero-title">Panorama Ejecutivo de Ventas</div>
          <div class="hero-subtitle">{subtitle}</div>
        </div>
        {logo_html}
      </div>
    </div>
    """, unsafe_allow_html=True)


# =============================================================================
# MÓDULO HUB — pantalla post-login REDISEÑADA
# (se muestra ANTES de cargar el Excel para que sea rápida)
# =============================================================================
if st.session_state.modulo_activo is None:
    hero_slot.empty()
    logo     = _find_cip_logo()
    logo_tag = (
        f'<img src="data:image/png;base64,{logo_to_base64_without_black(logo)}" '
        'alt="CIP" style="height:68px;width:auto;"/>'
        if logo
        else '<div style="font-size:1rem;font-weight:900;color:#93c5fd;letter-spacing:.08em;">CIP</div>'
    )

    st.markdown(f"""
    <style>
    @import url('{FONT_IMPORT}');
    *, *::before, *::after {{ font-family: {FONT_FAMILY} !important; }}
    .stApp {{
      background:
        radial-gradient(ellipse at 16% 20%, rgba(59,130,246,.20) 0%, transparent 50%),
        radial-gradient(ellipse at 80% 70%, rgba(139,92,246,.14) 0%, transparent 44%),
        linear-gradient(158deg,#020617 0%,#0a0f1e 44%,#060c1a 100%) !important;
    }}
    .stApp::before {{
      content:""; position:fixed; inset:0;
      background-image: radial-gradient(circle,rgba(96,165,250,.06) 1px,transparent 1px);
      background-size: 44px 44px; pointer-events:none; z-index:0;
    }}
    .hero {{ display: none !important; }}
    .main .block-container {{
      padding-top: 0 !important;
      padding-left: clamp(14px,3.5vw,60px) !important;
      padding-right: clamp(14px,3.5vw,60px) !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    # Top-bar
    st.markdown(f"""
    <div style="display:flex;align-items:center;justify-content:space-between;
                padding: clamp(18px,2.6vw,34px) 0 0 0;
                margin-bottom: clamp(26px,3.8vw,50px);">
      <div style="display:flex;align-items:center;gap:15px;">
        {logo_tag}
        <div style="width:1px;height:34px;background:rgba(148,163,184,.16);"></div>
        <div style="font-size:.70rem;font-weight:700;color:#475569;
                    letter-spacing:.14em;text-transform:uppercase;">
          Dirección Técnica
        </div>
      </div>
      <div class="launcher-badge">Portal analítico</div>
    </div>
    """, unsafe_allow_html=True)

    # Hero copy
    st.markdown("""
    <div style="margin-bottom: clamp(26px,4vw,46px);">
      <div style="font-size:.72rem;font-weight:800;letter-spacing:.22em;
                  text-transform:uppercase;color:#60a5fa;margin-bottom:12px;">
        Centro analítico de gestión
      </div>
      <h1 style="font-family:'Plus Jakarta Sans',sans-serif;
                 font-size:clamp(1.9rem,3.8vw,3.1rem);font-weight:900;
                 color:#f8fafc;letter-spacing:-1.1px;line-height:1.08;
                 margin:0 0 13px 0;">
        CIP
        <span style="background:linear-gradient(90deg,#60a5fa,#a78bfa);
                     -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                     background-clip:text;"> Analytics</span>
      </h1>
    </div>
    """, unsafe_allow_html=True)

    # Separador de módulos
    st.markdown("""
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;">
      <div style="font-size:.67rem;font-weight:800;letter-spacing:.16em;
                  text-transform:uppercase;color:#475569;">Módulos disponibles</div>
      <div style="flex:1;height:1px;background:rgba(148,163,184,.10);"></div>
    </div>
    """, unsafe_allow_html=True)

    # Grid de módulos — 3 columnas
    col_v, col_c, col_e = st.columns([1,1,1], gap="medium")

    with col_v:
        st.markdown("""
        <div class="mod-card">
          <div class="mod-icon">📊</div>
          <div class="mod-name">Ventas y Recaudación</div>
          <div class="mod-desc">
            Análisis de ventas domésticas, recaudación tributaria, inflación,
            confianza del consumidor y proyecciones econométricas.
          </div>
          <div class="mod-meta">
            <div class="mod-tag live">ACTUALIZADO A ENERO DE 2026</div>
            <div class="mod-tag count">7 secciones</div>
          </div>
          <div class="mod-stats">
            <div class="mod-stat-item">
              <span class="mod-stat-val">SRI</span>
              <span class="mod-stat-lbl">Fuente</span>
            </div>
            <div class="mod-stat-item">
              <span class="mod-stat-val">+20</span>
              <span class="mod-stat-lbl">Gráficos</span>
            </div>
            <div class="mod-stat-item">
              <span class="mod-stat-val">v2.3</span>
              <span class="mod-stat-lbl">Versión</span>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("🚀  Ingresar al módulo", use_container_width=True,
                     key="btn_ventas_launcher", type="primary"):
            st.session_state.modulo_activo = "ventas"
            st.rerun()

    with col_c:
        st.markdown("""
        <div class="mod-card locked">
          <div class="mod-icon" style="font-size:1.35rem;">🏭</div>
          <div class="mod-name">Comercio Exterior</div>
          <div class="mod-desc">
            Importaciones, exportaciones no petroleras, balanza comercial
            y participación sectorial por partida arancelaria.
          </div>
          <div class="mod-meta">
            <div class="mod-tag"
                 style="color:#f59e0b;background:rgba(245,158,11,.08);
                        border-color:rgba(245,158,11,.24);">Próximamente</div>
          </div>
          <div class="mod-stats">
            <div class="mod-stat-item">
              <span class="mod-stat-val">BCE</span>
              <span class="mod-stat-lbl">Fuente</span>
            </div>
            <div class="mod-stat-item">
              <span class="mod-stat-val">—</span>
              <span class="mod-stat-lbl">Gráficos</span>
            </div>
            <div class="mod-stat-item">
              <span class="mod-stat-val">Q3 25</span>
              <span class="mod-stat-lbl">ETA</span>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)
        st.button("🔒  No disponible aún", use_container_width=True,
                  key="btn_comercio", disabled=True)

    with col_e:
        st.markdown("""
        <div class="mod-card locked">
          <div class="mod-icon" style="font-size:1.35rem;">📈</div>
          <div class="mod-name">Empleo &amp; Productividad</div>
          <div class="mod-desc">
            Mercado laboral, empleo adecuado, subempleo y tasas de ocupación
            por sector y región.
          </div>
          <div class="mod-meta">
            <div class="mod-tag"
                 style="color:#f59e0b;background:rgba(245,158,11,.08);
                        border-color:rgba(245,158,11,.24);">Próximamente</div>
          </div>
          <div class="mod-stats">
            <div class="mod-stat-item">
              <span class="mod-stat-val">INEC</span>
              <span class="mod-stat-lbl">Fuente</span>
            </div>
            <div class="mod-stat-item">
              <span class="mod-stat-val">—</span>
              <span class="mod-stat-lbl">Gráficos</span>
            </div>
            <div class="mod-stat-item">
              <span class="mod-stat-val">Q4 25</span>
              <span class="mod-stat-lbl">ETA</span>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)
        st.button("🔒  No disponible aún", use_container_width=True,
                  key="btn_empleo", disabled=True)

    # Footer
    user_name_mod = st.session_state.get("usuario_nombre","")
    st.markdown(f"""
    <div style="margin-top:clamp(26px,4vw,46px);padding-top:15px;
                border-top:1px solid rgba(148,163,184,.10);
                display:flex;align-items:center;justify-content:space-between;
                flex-wrap:wrap;gap:8px;">
      <div style="font-size:.70rem;color:#1e293b;font-weight:500;">
        <strong style="color:#334155;">Panel Ejecutivo de Ventas</strong>
        &nbsp;·&nbsp; Cámara de Industrias y Producción &nbsp;·&nbsp; Dirección Técnica
      </div>
      <div style="font-size:.70rem;color:#1e293b;font-weight:500;">
        Sesión como <strong style="color:#334155;">{user_name_mod}</strong>
        &nbsp;·&nbsp; v2.3
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.stop()


# =============================================================================
# CARGA DE DATOS — solo se ejecuta cuando se entra a un módulo
# =============================================================================
if BASE_MADRE_FILE is None or not BASE_MADRE_FILE.exists():
    st.error("No se encontró BASE MADRE DASHBOARD VENTAS.xlsx. Revisa la ruta.")
    st.stop()

base_mtime = BASE_MADRE_FILE.stat().st_mtime
with st.spinner("Cargando datos…"):
    ventas_wide, ventas_month_cols, ventas_month_ts = load_ventas_wide(str(BASE_MADRE_FILE), base_mtime)
    if SRI_RECAUD_FILE is None or not SRI_RECAUD_FILE.exists():
        st.error("No se encontró BASE_SRI.xlsx.")
        st.stop()
    sri_mtime = SRI_RECAUD_FILE.stat().st_mtime
    sri_principal, sri_month_cols, sri_month_ts = load_sri_principal(str(SRI_RECAUD_FILE), sri_mtime)
    sri_secciones, sri_sec_date_cols = load_sri_secciones(str(SRI_RECAUD_FILE), sri_mtime)

last_month_sales = get_last_available_month_from_ts(ventas_month_ts)
last_month_rec = get_last_available_month_from_ts(sri_month_ts)
render_hero(f"Ecuador • Corte {get_last_month_label(last_month_sales)} • Valores en USD millones • Dirección Técnica CIP")


user_name = st.session_state.get("usuario_nombre") or st.session_state.usuario
_user_left, _user_right = st.columns([7.6, 2.4], vertical_alignment="center")
with _user_right:
    _spacer, _popover_col = st.columns([0.8, 2.2], vertical_alignment="center")
    with _popover_col:
        try:
            with st.popover(f"◈  {user_name}   ▾"):
                st.markdown(
                    f"""<div class="user-popover-card">
                            <div class="user-popover-kicker">Panel ejecutivo</div>
                            <div class="user-popover-name">{user_name}</div>
                            <div class="user-popover-mail">{st.session_state.usuario}</div>
                        </div>""",
                    unsafe_allow_html=True
                )
                if st.button("Cerrar sesión", use_container_width=True, key="logout_top"):
                    _logout()
        except Exception:
            if st.button(f"◈  {user_name} · Salir", use_container_width=True, key="logout_fallback"):
                _logout()

# =============================================================================
# FILTROS GENERALES
# =============================================================================
month_labels = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
_all_years   = sorted({ts.year for ts in ventas_month_ts})
_all_years   = sorted(set(_all_years) | {ts.year for ts in sri_month_ts})
_year_min    = _all_years[0] if _all_years else 2018
_year_max    = max(last_month_sales.year, last_month_rec.year)

with st.container(border=True):
    st.markdown('<div class="filter-title">Filtros del panel</div>', unsafe_allow_html=True)
    fa, fb, fc, fd = st.columns([1.2, 2.5, 1.4, 1.6])

    with fa:
        years_options = list(range(_year_max, _year_min - 1, -1))
        st.markdown('<p class="year-filter-label">Año</p>', unsafe_allow_html=True)
        selected_year = st.selectbox(
            "Año",
            options=years_options,
            index=0,
            key="year_filter_single",
            label_visibility="collapsed"
        )
        selected_years = [int(selected_year)]
        st.markdown(
            f'<p class="year-filter-caption">Período activo: {selected_year}</p>',
            unsafe_allow_html=True
        )
    with fb:
        selected_years_set = set(int(y) for y in selected_years)
        months_sales_sel = [ts.month for ts in ventas_month_ts if ts.year in selected_years_set]
        months_rec_sel = [ts.month for ts in sri_month_ts if ts.year in selected_years_set]
        last_month_sel = max(months_sales_sel + months_rec_sel) if (months_sales_sel or months_rec_sel) else int(last_month_sales.month)
        month_slider = st.slider(
            "Rango de meses (máx. 12)",
            1,
            12,
            (1, min(12, last_month_sel)),
            format="%d",
            key="month_range_slider"
        )
        _m0, _m1 = month_slider
        st.caption(f"Meses activos: {month_labels[_m0-1]} → {month_labels[_m1-1]}")
    with fc:
        contributor = st.selectbox("Segmento",
            ["Todos","Sociedades","Personas naturales"], index=0, key="contrib_filter")
    with fd:
        view = st.selectbox("Cobertura sectorial",
            ["Total","Solo comercio","Sin comercio"], index=0, key="view_filter")

selected_years = sorted(set(int(y) for y in selected_years))
selected_year  = max(selected_years)
_date_start    = pd.Timestamp(min(selected_years), _m0, 1).date()
_date_end      = pd.Timestamp(max(selected_years), min(_m1,12), 1).date()
date_range     = (_date_start, _date_end)
months         = list(range(_m0, _m1+1))
sectors_exclude= []
include_commerce = view in ["Total","Solo comercio"]
only_commerce    = (view == "Solo comercio")

def _period_mask(df_dates: pd.Series, years_sel: List[int], months_sel: List[int]) -> pd.Series:
    years_set = set(int(y) for y in years_sel)
    months_set = set(int(m) for m in months_sel)
    return df_dates.dt.year.isin(years_set) & df_dates.dt.month.isin(months_set)


def _selected_period_bounds(years_sel: List[int], months_sel: List[int]):
    y0, y1 = min(int(y) for y in years_sel), max(int(y) for y in years_sel)
    m0, m1 = min(int(m) for m in months_sel), max(int(m) for m in months_sel)
    start = pd.Timestamp(y0, m0, 1)
    end = pd.Timestamp(y1, m1, 1) + pd.offsets.MonthEnd(0)
    return start, end


def _series_window_until_cutoff(df: pd.DataFrame, date_col: str,
                                years_sel: List[int], months_sel: List[int],
                                fallback_tail: int = 24, min_unique_dates: int = 2):
    if df.empty:
        return df.copy(), False
    start, end = _selected_period_bounds(years_sel, months_sel)
    work = df.copy()
    work[date_col] = pd.to_datetime(work[date_col], errors="coerce")
    work = work.dropna(subset=[date_col]).sort_values(date_col)
    selected = work[(work[date_col] >= start) & (work[date_col] <= end)].copy()
    if selected[date_col].nunique() >= min_unique_dates:
        return selected.reset_index(drop=True), False
    hist = work[work[date_col] <= end].copy()
    if hist.empty:
        hist = selected.copy()
    if fallback_tail and not hist.empty:
        last_dates = list(pd.Series(hist[date_col].dropna().sort_values().unique()).tail(fallback_tail))
        hist = hist[hist[date_col].isin(last_dates)].copy()
    return hist.reset_index(drop=True), True


def _time_axis_for_series(df: pd.DataFrame, date_col: str = "fecha") -> dict:
    dates = pd.Series(pd.to_datetime(df[date_col], errors="coerce").dropna().sort_values().unique())
    axis = dict(type="date", tickformat="%b\n%Y", hoverformat="%b-%Y", showgrid=False)
    if len(dates) == 1:
        dt = pd.Timestamp(dates.iloc[0])
        axis["range"] = [dt - pd.Timedelta(days=20), dt + pd.Timedelta(days=20)]
        axis["dtick"] = 7 * 24 * 60 * 60 * 1000
    elif len(dates) <= 8:
        axis["dtick"] = "M1"
    else:
        axis["dtick"] = "M3"
    return axis


# =============================================================================
# PANEL PRINCIPAL — datos filtrados
# =============================================================================
ventas_f_wide = filter_ventas_wide(ventas_wide, contrib=contributor,
    exclude_sections_user=sectors_exclude,
    include_commerce=include_commerce, only_commerce=only_commerce)

df_ts = build_ventas_panel_from_wide(ventas_f_wide, ventas_month_cols, ventas_month_ts)
df_ts["fecha"] = pd.to_datetime(df_ts["fecha"]).dt.to_period("M").dt.to_timestamp()

df_filtered = df_ts[
    _period_mask(df_ts["fecha"], selected_years, months)
].copy()

data_max = pd.Timestamp(df_ts["fecha"].max()).replace(day=1)
last_month_effective = (
    pd.Timestamp(df_filtered["fecha"].max()).replace(day=1)
    if not df_filtered.empty
    else min(last_month_sales, pd.Timestamp(selected_year, _m1, 1), data_max)
)

months_for_kpi = [m for m in months if m <= int(pd.Timestamp(last_month_effective).month)]
panel          = calc_ytd_and_last_month(
    df_ts,
    last_month=last_month_effective,
    months_use=months_for_kpi,
    reference_year=int(pd.Timestamp(last_month_effective).year),
)


# =============================================================================
# KPIs
# =============================================================================
def _mm(x):       return x / 1e6
def _safe(d, k):  return d.get(k, np.nan) if d else np.nan


st.markdown(
    f'<div class="kpi-panel"><div class="kpi-panel-title">',
    unsafe_allow_html=True
)
k1, k2, k3, k4 = st.columns(4, gap="small")
vd = panel.get("ventas_domesticas", {}); vg = panel.get("ventas_gravadas", {})
ex = panel.get("exportaciones",     {}); vt = panel.get("ventas_totales",  {})

with k1:
    st.markdown(kpi_card("Ventas y exportaciones",
        fmt_usd_m(_mm(_safe(vt,"ytd_cur"))),
        delta_acum=fmt_pct_from_ratio(_safe(vt,"ytd_yoy")),
        delta_acum_type=get_delta_type_ratio(_safe(vt,"ytd_yoy")),
        delta_mes=fmt_pct_from_ratio(_safe(vt,"last_yoy")),
        delta_mes_type=get_delta_type_ratio(_safe(vt,"last_yoy")),
        sub=f"Ult. mes: {fmt_usd_m(_mm(_safe(vt,'last_cur')))}"
    ), unsafe_allow_html=True)
with k2:
    st.markdown(kpi_card("Ventas domésticas",
        fmt_usd_m(_mm(_safe(vd,"ytd_cur"))),
        delta_acum=fmt_pct_from_ratio(_safe(vd,"ytd_yoy")),
        delta_acum_type=get_delta_type_ratio(_safe(vd,"ytd_yoy")),
        delta_mes=fmt_pct_from_ratio(_safe(vd,"last_yoy")),
        delta_mes_type=get_delta_type_ratio(_safe(vd,"last_yoy")),
        sub=f"Ult. mes: {fmt_usd_m(_mm(_safe(vd,'last_cur')))}"
    ), unsafe_allow_html=True)
with k3:
    st.markdown(kpi_card("Ventas gravadas con IVA",
        fmt_usd_m(_mm(_safe(vg,"ytd_cur"))),
        delta_acum=fmt_pct_from_ratio(_safe(vg,"ytd_yoy")),
        delta_acum_type=get_delta_type_ratio(_safe(vg,"ytd_yoy")),
        delta_mes=fmt_pct_from_ratio(_safe(vg,"last_yoy")),
        delta_mes_type=get_delta_type_ratio(_safe(vg,"last_yoy")),
        sub=f"Ult. mes: {fmt_usd_m(_mm(_safe(vg,'last_cur')))}"
    ), unsafe_allow_html=True)
with k4:
    st.markdown(kpi_card("Exportaciones de bienes y servicios",
        fmt_usd_m(_mm(_safe(ex,"ytd_cur"))),
        delta_acum=fmt_pct_from_ratio(_safe(ex,"ytd_yoy")),
        delta_acum_type=get_delta_type_ratio(_safe(ex,"ytd_yoy")),
        delta_mes=fmt_pct_from_ratio(_safe(ex,"last_yoy")),
        delta_mes_type=get_delta_type_ratio(_safe(ex,"last_yoy")),
        sub=f"Ult. mes: {fmt_usd_m(_mm(_safe(ex,'last_cur')))}"
    ), unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)


# =============================================================================
# TABS — sin "Tendencia anual" ni "Exportaciones"
# =============================================================================
(tab_resumen, tab_sectores, tab_recaud,
 tab_proy, tab_geo, tab_inflacion, tab_confianza) = st.tabs([
    "Visión general", "Estructura sectorial",
    "Recaudación", "Proyecciones",
    "Territorio", "Precios", "Confianza",
])


# =============================================================================
# TAB: VISIÓN GENERAL
# — Años móviles PRIMERO, luego tendencia mensual
# =============================================================================
with tab_resumen:

    # ── 1. AÑOS MÓVILES CON ENFOQUE EJECUTIVO (movido desde "Tendencia anual") ──
    section_header("Años móviles con enfoque ejecutivo")

    am   = load_anios_moviles(str(BASE_MADRE_FILE), base_mtime).copy()
    am_w = am.tail(24).copy()

    col_am1, col_am2 = st.columns(2, gap="large")
    with col_am1:
        with card("Año móvil • Ventas domésticas y gravadas (variación %)"):
            fig = go.Figure()
            last_points_am1 = []
            if "Ventas domésticas.1" in am_w.columns:
                ser_dom_am = am_w[["fecha", "Ventas domésticas.1"]].dropna().copy()
                ser_dom_am["Ventas domésticas.1"] = ser_dom_am["Ventas domésticas.1"].astype(float) * 100.0
                fig.add_trace(go.Scatter(
                    x=ser_dom_am["fecha"],
                    y=ser_dom_am["Ventas domésticas.1"],
                    mode="lines+markers", name="Ventas domésticas",
                    line=dict(width=2.8, color=BULLETIN_GOLD),
                    marker=dict(size=6, symbol="circle-open", color=BULLETIN_GOLD, line=dict(width=2, color=BULLETIN_GOLD_DARK))
                ))
                if len(ser_dom_am):
                    last_points_am1.append(("Ventas domésticas", pd.Timestamp(ser_dom_am["fecha"].iloc[-1]), float(ser_dom_am["Ventas domésticas.1"].iloc[-1]), BULLETIN_GOLD))
            if "% Ventas gravadas" in am_w.columns:
                ser_iva_am = am_w[["fecha", "% Ventas gravadas"]].dropna().copy()
                ser_iva_am["% Ventas gravadas"] = pd.to_numeric(ser_iva_am["% Ventas gravadas"], errors="coerce") * 100.0
                ser_iva_am = ser_iva_am.dropna(subset=["% Ventas gravadas"])
                fig.add_trace(go.Scatter(
                    x=ser_iva_am["fecha"],
                    y=ser_iva_am["% Ventas gravadas"],
                    mode="lines+markers", name="Ventas gravadas (IVA)",
                    line=dict(width=2.8, color=BULLETIN_BLUE),
                    marker=dict(size=6, symbol="circle-open", color=BULLETIN_BLUE, line=dict(width=2, color=BULLETIN_BLUE_DARK))
                ))
                if len(ser_iva_am):
                    last_points_am1.append(("Ventas gravadas (IVA)", pd.Timestamp(ser_iva_am["fecha"].iloc[-1]), float(ser_iva_am["% Ventas gravadas"].iloc[-1]), BULLETIN_BLUE))
            fig.add_hline(y=0, line_width=1, line_dash="solid", opacity=0.35)
            apply_fluent_layout(fig, height=370, hovermode="x unified",
                yaxis=dict(title="Variación interanual (%)"),
                legend=dict(orientation="h", yanchor="top", y=-0.14, xanchor="center", x=0.5),
                xaxis=dict(rangeslider=RANGE_SLIDER),
                margin=dict(l=38,r=92,t=40,b=65))
            fig.update_xaxes(showgrid=False)
            fig.update_yaxes(showgrid=True, gridcolor=colors.gray30)
            yshift_am1 = {"Ventas domésticas": -16, "Ventas gravadas (IVA)": 16}
            for label_name, x_last, y_last, color in last_points_am1:
                fig.add_annotation(
                    x=x_last, y=y_last,
                    text=f"{y_last:.1f}%".replace(".", ","),
                    showarrow=False, xanchor="left",
                    xshift=10, yshift=yshift_am1.get(label_name, 0),
                    font=dict(size=10.5, color=color),
                    bgcolor="rgba(255,255,255,0.86)",
                    bordercolor="rgba(148,163,184,0.22)",
                    borderwidth=1, borderpad=4
                )
            show_plotly(fig, "am_yoy_dom_iva")

    with col_am2:
        with card("Año móvil • Exportaciones de bienes y servicios y ventas totales (variación %)"):
            fig = go.Figure()
            last_points_am2 = []
            if "Total exportaciones" in am_w.columns:
                ser_exp_am = am_w[["fecha", "Total exportaciones"]].dropna().copy()
                ser_exp_am["Total exportaciones"] = pd.to_numeric(ser_exp_am["Total exportaciones"], errors="coerce") * 100.0
                ser_exp_am = ser_exp_am.dropna(subset=["Total exportaciones"])
                fig.add_trace(go.Scatter(
                    x=ser_exp_am["fecha"],
                    y=ser_exp_am["Total exportaciones"],
                    mode="lines+markers", name="Exportaciones de bienes y servicios",
                    line=dict(width=2.6, color=BULLETIN_BLUE_LIGHT),
                    marker=dict(size=6, symbol="circle-open", color=BULLETIN_BLUE_LIGHT, line=dict(width=2, color=BULLETIN_BLUE_DARK))
                ))
                if len(ser_exp_am):
                    last_points_am2.append(("Exportaciones de bienes y servicios", pd.Timestamp(ser_exp_am["fecha"].iloc[-1]), float(ser_exp_am["Total exportaciones"].iloc[-1]), BULLETIN_BLUE_LIGHT))
            if "Ventas totales" in am_w.columns:
                ser_tot_am = am_w[["fecha", "Ventas totales"]].dropna().copy()
                ser_tot_am["Ventas totales"] = pd.to_numeric(ser_tot_am["Ventas totales"], errors="coerce") * 100.0
                ser_tot_am = ser_tot_am.dropna(subset=["Ventas totales"])
                fig.add_trace(go.Scatter(
                    x=ser_tot_am["fecha"],
                    y=ser_tot_am["Ventas totales"],
                    mode="lines+markers", name="Ventas totales",
                    line=dict(width=2.8, color=BULLETIN_BLUE_DARK),
                    marker=dict(size=6, symbol="circle-open", color=BULLETIN_BLUE_DARK, line=dict(width=2, color=BULLETIN_BLUE_DARK))
                ))
                if len(ser_tot_am):
                    last_points_am2.append(("Ventas totales", pd.Timestamp(ser_tot_am["fecha"].iloc[-1]), float(ser_tot_am["Ventas totales"].iloc[-1]), BULLETIN_BLUE_DARK))
            fig.add_hline(y=0, line_width=1, line_dash="solid", opacity=0.35)
            apply_fluent_layout(fig, height=370, hovermode="x unified",
                yaxis=dict(title="Variación interanual (%)"),
                legend=dict(orientation="h", yanchor="top", y=-0.14, xanchor="center", x=0.5),
                xaxis=dict(rangeslider=RANGE_SLIDER),
                margin=dict(l=38,r=92,t=40,b=65))
            fig.update_xaxes(showgrid=False)
            fig.update_yaxes(showgrid=True, gridcolor=colors.gray30)
            yshift_am2 = {"Exportaciones de bienes y servicios": -16, "Ventas totales": 16}
            for label_name, x_last, y_last, color in last_points_am2:
                fig.add_annotation(
                    x=x_last, y=y_last,
                    text=f"{y_last:.1f}%".replace(".", ","),
                    showarrow=False, xanchor="left",
                    xshift=10, yshift=yshift_am2.get(label_name, 0),
                    font=dict(size=10.5, color=color),
                    bgcolor="rgba(255,255,255,0.86)",
                    bordercolor="rgba(148,163,184,0.22)",
                    borderwidth=1, borderpad=4
                )
            show_plotly(fig, "am_yoy_exp_tot")

    # ── 2. TENDENCIAS MENSUALES ───────────────────────────────────────────────
    section_header("Tendencias y desempeño mensual")

    def _yoy_pct(series: pd.Series) -> pd.Series:
        return series.pct_change(12) * 100.0

    df_plot = df_filtered.sort_values("fecha").copy()
    df_plot["fecha"] = pd.to_datetime(df_plot["fecha"]).dt.to_period("M").dt.to_timestamp()
    df_plot_full = df_ts.sort_values("fecha").copy()
    df_plot_full["fecha"] = pd.to_datetime(df_plot_full["fecha"]).dt.to_period("M").dt.to_timestamp()

    s_dom_full = pd.Series(df_plot_full["ventas_domesticas"].values, index=df_plot_full["fecha"]) / 1e6
    s_iva_full = pd.Series(df_plot_full["ventas_gravadas"].values,   index=df_plot_full["fecha"]) / 1e6
    s_exp_full = pd.Series(df_plot_full["exportaciones"].values,     index=df_plot_full["fecha"]) / 1e6
    s_tot_full = pd.Series(df_plot_full["ventas_totales"].values,    index=df_plot_full["fecha"]) / 1e6

    s_dom = pd.Series(df_plot["ventas_domesticas"].values, index=df_plot["fecha"]) / 1e6
    s_iva = pd.Series(df_plot["ventas_gravadas"].values,   index=df_plot["fecha"]) / 1e6
    s_exp = pd.Series(df_plot["exportaciones"].values,     index=df_plot["fecha"]) / 1e6
    s_tot = pd.Series(df_plot["ventas_totales"].values,    index=df_plot["fecha"]) / 1e6
    x_month_labels = [get_last_month_label(ts) for ts in s_dom.index]

    yoy_dom_full = _yoy_pct(s_dom_full)
    yoy_iva_full = _yoy_pct(s_iva_full)
    yoy_exp_full = _yoy_pct(s_exp_full)
    yoy_tot_full = _yoy_pct(s_tot_full)
    yoy_dom = yoy_dom_full.reindex(s_dom.index)

    last_dt = pd.Timestamp(last_month_effective).to_period("M").to_timestamp()
    if last_dt not in s_dom_full.index:
        last_dt = pd.Timestamp(df_plot_full["fecha"].max()).to_period("M").to_timestamp()
    last_label = get_last_month_label(last_dt)
    last_dom   = float(s_dom_full.get(last_dt, np.nan))
    last_exp   = float(s_exp_full.get(last_dt, np.nan))
    last_tot   = float(s_tot_full.get(last_dt, np.nan))

    last_yoy = {
        "Ventas domésticas":     float(yoy_dom_full.get(last_dt, np.nan)),
        "Ventas gravadas (IVA)": float(yoy_iva_full.get(last_dt, np.nan)),
        "Exportaciones de bienes y servicios": float(yoy_exp_full.get(last_dt, np.nan)),
        "Ventas totales":        float(yoy_tot_full.get(last_dt, np.nan)),
    }

    row1 = st.columns([2.2, 1.1], gap="large")
    row2 = st.columns([2.2, 1.1], gap="large")

    with row1[0]:
        with card(f"Evolución mensual (USD millones) • Corte {last_label}"):
            fig1 = make_subplots(specs=[[{"secondary_y": True}]])
            draw_mode = "lines+markers" if len(s_dom.index) <= 1 else "lines+markers"

            fig1.add_trace(go.Scatter(
                x=x_month_labels, y=s_dom.values,
                name="Ventas domésticas", mode=draw_mode,
                line=dict(width=2.8, color=BULLETIN_GOLD),
                marker=dict(size=7, color=BULLETIN_GOLD, line=dict(width=1.1, color="#ffffff"))
            ), secondary_y=False)
            fig1.add_trace(go.Scatter(
                x=x_month_labels, y=s_iva.values,
                name="Ventas gravadas (IVA)", mode=draw_mode,
                line=dict(width=2.8, color=BULLETIN_BLUE),
                marker=dict(size=7, color=BULLETIN_BLUE, line=dict(width=1.1, color="#ffffff"))
            ), secondary_y=False)
            fig1.add_trace(go.Scatter(
                x=x_month_labels, y=s_tot.values,
                name="Ventas y exportaciones", mode=draw_mode,
                line=dict(width=3.2, color=BULLETIN_BLUE_DARK, dash="solid"),
                marker=dict(size=7, color=BULLETIN_BLUE_DARK, line=dict(width=1.1, color="#ffffff"))
            ), secondary_y=False)
            fig1.add_trace(go.Scatter(
                x=x_month_labels, y=s_exp.values,
                name="Exportaciones de bienes y servicios (eje der.)", mode=draw_mode,
                line=dict(width=2.6, color=BULLETIN_BLUE_LIGHT, dash="dot"),
                marker=dict(size=7, color=BULLETIN_BLUE_LIGHT, line=dict(width=1.1, color="#ffffff"))
            ), secondary_y=True)

            fig1.update_yaxes(
                title_text="USD millones (ventas)", secondary_y=False,
                showgrid=True, gridcolor=colors.gray30, tickfont=dict(size=10, color=colors.gray90),
                tickformat="~s"
            )
            fig1.update_yaxes(
                title_text="USD millones (exportaciones)", secondary_y=True,
                showgrid=False, tickfont=dict(size=10, color=BULLETIN_BLUE_LIGHT),
                tickformat="~s"
            )
            fig1.update_xaxes(
                showgrid=False, showline=True, linecolor=colors.gray40, linewidth=1,
                tickfont=dict(size=10, color=colors.gray90), rangeslider=RANGE_SLIDER
            )
            fig1.update_layout(
                height=380, hovermode="x unified",
                legend=dict(
                    orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    font=dict(size=10, color=colors.gray130)
                ),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="DM Sans, Inter, Segoe UI, sans-serif", size=11, color=colors.gray130),
                margin=dict(l=52, r=60, t=60, b=36),
            )
            show_plotly(fig1, "resumen_lineas_mm")

    with row1[1]:
        with card(f"Composición del total ventas y exportaciones • Acumulado a {last_label}"):
            comp_dom = max(0.0, np.nan_to_num(_mm(_safe(vd, "ytd_cur"))))
            comp_exp = max(0.0, np.nan_to_num(_mm(_safe(ex, "ytd_cur"))))
            comp_tot = max(0.0, np.nan_to_num(_mm(_safe(vt, "ytd_cur"))))
            donut_vals = [comp_dom, comp_exp]
            fig3 = go.Figure(data=[go.Pie(
                labels=["Ventas domésticas", "Exportaciones de bienes y servicios"],
                values=donut_vals, hole=0.64, sort=False,
                textinfo="percent", textfont=dict(size=11),
                hovertemplate="%{label}<br>%{value:,.1f} MM<br>%{percent}<extra></extra>",
                marker=dict(
                    colors=[BULLETIN_GOLD, BULLETIN_BLUE],
                    line=dict(color="#ffffff", width=2)
                )
            )])
            fig3.update_layout(
                height=380, margin=dict(l=8, r=8, t=18, b=8), showlegend=True,
                legend=dict(
                    orientation="h", yanchor="bottom", y=-0.08,
                    xanchor="center", x=0.5, font=dict(size=10)
                ),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="DM Sans, Inter, Segoe UI", color=colors.gray130, size=11)
            )
            fig3.add_annotation(
                x=0.5, y=0.53,
                text=(
                    f"<span style='font-size:11px;color:{colors.gray90}'>Acumulado</span>"
                    f"<br><b>{fmt_num_latam(comp_tot, 0)}</b>"
                    "<br><span style='font-size:11px'>MM</span>"
                ),
                showarrow=False, font=dict(size=17, color=colors.gray160), align="center"
            )
            show_plotly(
                fig3, "resumen_donut_comp",
                footnote=""
            )

    with row2[0]:
        with card("Crecimiento interanual • Ventas domésticas (%)"):
            yoy_window = yoy_dom_full.dropna()
            yoy_window = yoy_window.loc[yoy_window.index >= (last_dt - pd.DateOffset(months=22))]
            bar_colors = [bulletin_diverging_color(v) for v in yoy_window.values]
            fig2 = go.Figure()
            yoy_x_labels = [get_last_month_label(ts) for ts in yoy_window.index]
            fig2.add_trace(go.Bar(
                x=yoy_x_labels, y=yoy_window.values,
                marker=dict(color=bar_colors),
                hovertemplate="%{x}<br>Variación: %{y:.1f}%<extra></extra>",
                name="Variación"
            ))
            fig2.add_hline(y=0, line_width=1, line_dash="solid", opacity=0.35)
            apply_fluent_layout(fig2, height=310, hovermode="x unified",
                yaxis=dict(title="Variación interanual (%)"),
                xaxis=dict(rangeslider=RANGE_SLIDER))
            fig2.update_xaxes(showgrid=False)
            fig2.update_yaxes(showgrid=True, gridcolor=colors.gray30)
            show_plotly(fig2, "resumen_yoy_bar")

    with row2[1]:
        with card(f"Variación por componente • {last_label}"):
            comp_names = list(last_yoy.keys())
            comp_vals  = [last_yoy[k] for k in comp_names]
            order      = np.argsort(np.nan_to_num(comp_vals))[::-1]
            comp_names = [comp_names[i] for i in order]
            comp_vals  = [comp_vals[i]  for i in order]
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(
                x=comp_vals, y=comp_names, orientation="h",
                marker=dict(color=[bulletin_diverging_color(v) for v in comp_vals]),
                hovertemplate="%{y}<br>Variación: %{x:.1f}%<extra></extra>",
                name="Variación"
            ))
            fig4.add_vline(x=0, line_width=1, line_dash="solid", opacity=0.35)
            apply_fluent_layout(fig4, height=310,
                xaxis=dict(title="Variación interanual (%)"), yaxis=dict(title=""),
                margin=dict(l=130,r=12,t=30,b=30), showlegend=False)
            fig4.update_xaxes(showgrid=True, gridcolor=colors.gray30)
            fig4.update_yaxes(showgrid=False)
            show_plotly(fig4, "resumen_yoy_comp")

    # Nota de exclusiones — UNA SOLA VEZ al final del panel
    show_panel_footnote()


# =============================================================================
# TAB: ESTRUCTURA SECTORIAL
# =============================================================================
with tab_sectores:
    section_header("Desempeño sectorial")

    with card("Configuración"):
        c1, c2 = st.columns([2.2, 1.2])
        with c1:
            sector_metric = st.selectbox("Tipo de ventas",
                ["Ventas domésticas","Ventas gravadas con IVA","Exportaciones de bienes y servicios","Ventas y exportaciones"],
                index=0, key="sector_metric_select")
        with c2:
            sector_sort_note = st.selectbox("Orden barras",
                ["Ascendente","Orden fijo"], index=0, key="sector_sort_select")

    if sector_metric == "Ventas domésticas":
        measures_sec = [MEASURE_MAP["ventas_domesticas"]]; title_prefix = "Ventas domésticas"
    elif sector_metric == "Ventas gravadas con IVA":
        measures_sec = IVA_MEASURES; title_prefix = "Ventas gravadas con IVA"
    elif sector_metric == "Exportaciones de bienes y servicios":
        measures_sec = EXPORT_MEASURES; title_prefix = "Exportaciones de bienes y servicios"
    else:
        measures_sec = [MEASURE_MAP["ventas_totales"]]; title_prefix = "Ventas y exportaciones"

    end_month_s = int(pd.Timestamp(last_month_effective).month)
    year_cur_s  = int(pd.Timestamp(last_month_effective).year)
    year_prev_s = year_cur_s - 1

    df_sec = build_sector_panel_wide(
        ventas_wide, ventas_month_cols, ventas_month_ts,
        contrib=contributor, measures=measures_sec,
        year_cur=year_cur_s, end_month=end_month_s, months_sel=months,
        include_commerce=include_commerce, only_commerce=only_commerce,
        exclude_sections_user=sectors_exclude
    )

    meses_abr     = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    periodo_txt_s = f"Enero – {meses_abr[end_month_s-1]}"
    sort_asc      = (sector_sort_note == "Ascendente")

    cL, cR = st.columns([1.2,1.2], gap="large")
    with cL:
        with card(f"Variación relativa • {periodo_txt_s}"):
            plot_sector_relative(df_sec, y_prev=year_prev_s, y_cur=year_cur_s,
                title=f"{title_prefix} – Variación relativa (%)",
                key=f"sectores_rel_{_norm_upper(title_prefix)}",
                footnote=f"")

    with cR:
        with card(f"Variación absoluta (USD millones) • {periodo_txt_s}"):
            plot_sector_absolute(df_sec,
                title=f"{title_prefix} – Variación absoluta (USD millones)",
                key=f"sectores_abs_{_norm_upper(title_prefix)}",
                sort_asc=sort_asc,
                footnote=f"")

    with card("Treemap de participación y crecimiento sectorial"):
        df_t = df_sec.copy()
        df_t["YoY"]       = df_t[f"YoY_{year_cur_s}"]
        df_t["YoY_label"] = df_t["YoY"].apply(lambda v: "-" if pd.isna(v) else f"{v:+.1f}%")
        total_nivel       = float(df_t["Nivel_MM"].sum()) if len(df_t) else np.nan
        df_t["Share"]     = np.where(total_nivel > 0, df_t["Nivel_MM"]/total_nivel, np.nan)
        fig_tm = px.treemap(
            df_t, path=["sector"], values="Nivel_MM", color="YoY",
            color_continuous_scale=[
                [0.0,"#991b1b"],[0.25,"#ef4444"],[0.5,"#f8fafc"],[0.75,"#3b82f6"],[1.0,"#1e3a8a"]
            ],
            custom_data=["YoY_label","YoY","Share","Nivel_MM"],
        )
        fig_tm.update_traces(
            texttemplate=(
                "<b style='font-size:14px'>%{label}</b>"
                "<br><span style='font-size:11px'>%{value:,.1f} MM</span>"
                "<br><span style='font-size:12px'>%{customdata[0]}</span>"
            ),
            textposition="middle center",
            textfont=dict(size=12, color="#0f172a", family="DM Sans, Inter, sans-serif"),
            marker=dict(line=dict(color="#ffffff",width=2.5), cornerradius=6),
            hovertemplate=(
                "<b>%{label}</b>"
                "<br>Nivel: %{customdata[3]:,.2f} MM"
                "<br>Participación: %{customdata[2]:.2%}"
                "<br>Variación interanual: %{customdata[1]:+.2f}%"
                "<extra></extra>"
            ),
            root_color="#f1f5f9",
        )
        fig_tm.update_layout(
            height=500, margin=dict(l=6,r=6,t=44,b=6),
            coloraxis_colorbar=dict(
                title=dict(text="YoY (%)", font=dict(size=11)),
                tickformat=",.2f", thickness=13, len=0.75,
                bgcolor="rgba(248,250,252,.92)", bordercolor="#e2e8f0",
                borderwidth=1, tickfont=dict(size=10),
            ),
            font=dict(family="DM Sans, Inter, sans-serif"),
        )
        show_plotly(fig_tm, "sectores_treemap",
            footnote="El tamaño de cada bloque representa el nivel de ventas (MM), mientras que el color refleja la variación interanual (%).")

    # Nota — al final del panel
    show_panel_footnote()


# =============================================================================
# TAB: RECAUDACIÓN TRIBUTARIA — revisión completa de cálculos y visualización
# =============================================================================
with tab_recaud:
    section_header("Recaudación tributaria")

    contrib_opts = ["Todos", "Sociedades", "Personas naturales"]

    with card("Configuración"):
        contrib_recaud = st.selectbox(
            "Tipo de contribuyente",
            options=contrib_opts,
            index=0,
            key="recaud_sri_contrib_select",
        )

    rec_ts = build_recaudacion_series(
        sri_principal,
        sri_month_cols,
        sri_month_ts,
        contrib=contrib_recaud
    )
    rec_ts["fecha"] = pd.to_datetime(rec_ts["fecha"]).dt.to_period("M").dt.to_timestamp()

    rec_f = rec_ts[
        _period_mask(rec_ts["fecha"], selected_years, months)
    ].copy()

    _years_avail = sorted({ts.year for ts in sri_month_ts})
    year_boletin = selected_year if selected_year in _years_avail else int(last_month_rec.year)
    
    _months_y = [ts.month for ts in sri_month_ts if ts.year == year_boletin]
    end_month_bol = min(_m1, max(_months_y)) if _months_y else int(last_month_rec.month)
    rec_months_use = [m for m in months if m <= end_month_bol] or list(range(1, end_month_bol + 1))

    year_prev = year_boletin - 1
    month_cur = end_month_bol
    meses_abr = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    periodo_txt = f"Enero – {meses_abr[end_month_bol-1]} {year_boletin}"

    def _sum_period(df, col, year, months_use):
        return float(df.loc[(df["fecha"].dt.year == year) & (df["fecha"].dt.month.isin(months_use)), col].sum())

    def _sum_month(df, col, year, month):
        return float(df.loc[(df["fecha"].dt.year == year) & (df["fecha"].dt.month == month), col].sum())

    bruta_ytd_cur = _sum_period(rec_ts, "recaud_bruta", year_boletin, rec_months_use)
    bruta_ytd_prev = _sum_period(rec_ts, "recaud_bruta", year_prev, rec_months_use)
    bruta_ytd_var = (bruta_ytd_cur / bruta_ytd_prev - 1.0) if bruta_ytd_prev > 0 else np.nan
    bruta_ytd_abs = bruta_ytd_cur - bruta_ytd_prev

    bruta_mes_cur = _sum_month(rec_ts, "recaud_bruta", year_boletin, month_cur)
    bruta_mes_prev = _sum_month(rec_ts, "recaud_bruta", year_prev, month_cur)
    bruta_mes_var = (bruta_mes_cur / bruta_mes_prev - 1.0) if bruta_mes_prev > 0 else np.nan
    bruta_mes_abs = bruta_mes_cur - bruta_mes_prev

    neta_ytd_cur = _sum_period(rec_ts, "recaud_neta", year_boletin, rec_months_use)
    neta_ytd_prev = _sum_period(rec_ts, "recaud_neta", year_prev, rec_months_use)
    neta_ytd_var = (neta_ytd_cur / neta_ytd_prev - 1.0) if neta_ytd_prev > 0 else np.nan
    neta_ytd_abs = neta_ytd_cur - neta_ytd_prev

    neta_mes_cur = _sum_month(rec_ts, "recaud_neta", year_boletin, month_cur)
    neta_mes_prev = _sum_month(rec_ts, "recaud_neta", year_prev, month_cur)
    neta_mes_var = (neta_mes_cur / neta_mes_prev - 1.0) if neta_mes_prev > 0 else np.nan
    neta_mes_abs = neta_mes_cur - neta_mes_prev

    rk1, rk2, rk3, rk4 = st.columns(4)
    with rk1:
        st.markdown(kpi_card(
            f"Recaudación bruta acumulada • {periodo_txt}",
            fmt_usd_m(bruta_ytd_cur / 1e6),
            delta_acum=fmt_pct_from_ratio(bruta_ytd_var),
            delta_acum_type=get_delta_type_ratio(bruta_ytd_var),
            sub=f"Δ absoluta: {fmt_usd_m(bruta_ytd_abs / 1e6)}"
        ), unsafe_allow_html=True)

    with rk2:
        st.markdown(kpi_card(
            f"Recaudación bruta mensual • {meses_abr[month_cur-1]} {year_boletin}",
            fmt_usd_m(bruta_mes_cur / 1e6),
            delta_mes=fmt_pct_from_ratio(bruta_mes_var),
            delta_mes_type=get_delta_type_ratio(bruta_mes_var),
            sub=f"Δ absoluta: {fmt_usd_m(bruta_mes_abs / 1e6)}"
        ), unsafe_allow_html=True)

    with rk3:
        st.markdown(kpi_card(
            f"Recaudación neta acumulada • {periodo_txt}",
            fmt_usd_m(neta_ytd_cur / 1e6),
            delta_acum=fmt_pct_from_ratio(neta_ytd_var),
            delta_acum_type=get_delta_type_ratio(neta_ytd_var),
            sub=f"Δ absoluta: {fmt_usd_m(neta_ytd_abs / 1e6)}"
        ), unsafe_allow_html=True)

    with rk4:
        st.markdown(kpi_card(
            f"Recaudación neta mensual • {meses_abr[month_cur-1]} {year_boletin}",
            fmt_usd_m(neta_mes_cur / 1e6),
            delta_mes=fmt_pct_from_ratio(neta_mes_var),
            delta_mes_type=get_delta_type_ratio(neta_mes_var),
            sub=f"Δ absoluta: {fmt_usd_m(neta_mes_abs / 1e6)}"
        ), unsafe_allow_html=True)

    comparacion_rec = st.radio(
        "Vista de comparación",
        ["Acumulado", "Mensual"],
        horizontal=True,
        key="recaud_comparacion_radio",
    )
    modo_acumulado = (comparacion_rec == "Acumulado")
    periodo_comp = periodo_txt if modo_acumulado else f"{meses_abr[month_cur-1]} {year_boletin}"
    color_comp = BULLETIN_CURRENT

    c1, c2 = st.columns(2, gap="large")
    with c1:
        with card(f"Total recaudación bruta (USD millones) • {periodo_comp}"):
            plot_recaud_comparison(
                cur_value=(bruta_ytd_cur if modo_acumulado else bruta_mes_cur) / 1e6,
                prev_value=(bruta_ytd_prev if modo_acumulado else bruta_mes_prev) / 1e6,
                cur_label=str(year_boletin),
                prev_label=str(year_prev),
                title=f"Comparación {comparacion_rec.lower()}",
                key=f"recaud_bruta_{comparacion_rec.lower()}_cmp",
                color_cur=color_comp,
            )
    with c2:
        with card(f"Total recaudación neta (USD millones) • {periodo_comp}"):
            plot_recaud_comparison(
                cur_value=(neta_ytd_cur if modo_acumulado else neta_mes_cur) / 1e6,
                prev_value=(neta_ytd_prev if modo_acumulado else neta_mes_prev) / 1e6,
                cur_label=str(year_boletin),
                prev_label=str(year_prev),
                title=f"Comparación {comparacion_rec.lower()}",
                key=f"recaud_neta_{comparacion_rec.lower()}_cmp",
                color_cur=color_comp,
            )

    df_tax = recaud_tax_snapshot(
        sri_principal,
        sri_month_cols,
        sri_month_ts,
        contrib=contrib_recaud,
        year_cur=year_boletin,
        month_cur=month_cur,
        months_use=rec_months_use
    )

    with card(f"Recaudación por tipo de impuesto (USD millones) • {periodo_comp}"):
        if modo_acumulado:
            plot_tax_grouped_hbars(
                df_tax,
                prev_col="ytd_prev_mm",
                cur_col="ytd_cur_mm",
                var_col="ytd_var",
                title="Comparación acumulada",
                key="recaud_tax_cmp",
                year_prev=year_prev,
                year_cur=year_boletin,
            )
        else:
            plot_tax_grouped_hbars(
                df_tax,
                prev_col="mes_prev_mm",
                cur_col="mes_cur_mm",
                var_col="mes_var",
                title="Comparación mensual",
                key="recaud_tax_cmp",
                year_prev=year_prev,
                year_cur=year_boletin,
            )

    section_header("Recaudación sectorial de sociedades")

    df_sec = build_sector_panel_sri(
        sri_secciones,
        year_cur=year_boletin,
        months_use=rec_months_use
    )

    sec1, sec2 = st.columns(2, gap="large")
    with sec1:
        with card(f"Variación relativa • {periodo_txt}"):
            plot_sector_relative(
                df_sec,
                y_prev=year_prev,
                y_cur=year_boletin,
                title="Sociedades – variación relativa (%)",
                key="recaud_sector_rel_sri"
            )

    with sec2:
        with card(f"Variación absoluta (USD millones) • {periodo_txt}"):
            plot_sector_absolute(
                df_sec,
                title="Sociedades – variación absoluta (USD millones)",
                key="recaud_sector_abs_sri",
                sort_asc=True
            )

    st.markdown(
        """
        <div class="panel-footnote">
        ⓘ <b>Nota metodológica de recaudación:</b> la recaudación <b>bruta</b> se construye con la suma de <i>VALOR RECAUDADO</i>; 
        la recaudación <b>neta</b> se calcula como <i>VALOR RECAUDADO - VALOR NOTAS CRÉDITO - VALOR COMPENSACIONES - VALOR TBC</i>. 
        La comparación acumulada enfrenta el mismo conjunto de meses del año vigente frente al año previo, mientras que la comparación mensual usa exclusivamente el mes de corte. 
        La desagregación por impuesto se obtiene del campo <i>GRUPO</i> de la hoja PRINCIPAL y la desagregación sectorial proviene de la hoja SECCIONES, correspondiente a sociedades.
        Fuente: SRI · Elaboración: DT-CIP
        </div>
        """,
        unsafe_allow_html=True
    )

# =============================================================================
# TAB: TERRITORIO
# =============================================================================
with tab_geo:
    section_header("Distribución geográfica cantonal")

    geo             = load_geo_cantonal(str(BASE_MADRE_FILE), base_mtime)
    shp_canton_path = find_shp_file(DATA_DIR, preferred_type="canton")
    shp_prov_path   = find_shp_file(DATA_DIR, preferred_type="provincia")

    with card("Configuración del mapa"):
        g1, g2, g3, g4 = st.columns([1.6,1.2,1.6,1.6])
        geo_tipo_map = {
            "Ventas domésticas":"VENTAS DOMÉSTICAS",
            "Ventas gravadas":"VENTAS GRAVADAS",
            "Total ventas y exportaciones":"TOTAL VENTAS Y EXPORTACIONES",
            "Exportaciones de bienes y servicios":"EXPORTACIONES",
        }
        geo_cat_map = {
            "Total":"TOTAL",
            "Exclusión de comercio":"EXCLUSIÓN DE COMERCIO",
            "Comercio":"COMERCIO",
        }
        with g1:
            geo_tipo_label = st.selectbox("Indicador", list(geo_tipo_map.keys()), index=2, key="geo_tipo_select")
        with g2:
            geo_cat_label  = st.selectbox("Segmento",  list(geo_cat_map.keys()),  index=0, key="geo_cat_select")
        with g3:
            _geo_year_prev = int(pd.to_numeric(geo.get("GEO_YEAR_PREV", pd.Series([year_prev_s])), errors="coerce").dropna().iloc[0]) if len(geo) else int(year_prev_s)
            _geo_year_cur  = int(pd.to_numeric(geo.get("GEO_YEAR_CUR",  pd.Series([year_cur_s])),  errors="coerce").dropna().iloc[0]) if len(geo) else int(year_cur_s)
            map_metric = st.selectbox("Variable del color",
                [f"Variación {_geo_year_cur} vs {_geo_year_prev} (%)", f"Variación absoluta {_geo_year_cur} vs {_geo_year_prev} (MM)", f"Nivel de ventas {_geo_year_cur} (MM)"],
                index=0, key="geo_metric_select")
        with g4:
            _prov_mapping_raw = {}
            if HAS_GEOPANDAS and shp_canton_path and shp_canton_path.exists():
                _prov_mapping_raw = load_province_canton_mapping(str(shp_canton_path), shp_canton_path.stat().st_mtime)
            _prov_display  = _prov_mapping_raw.get("prov_display",{})
            _prov_opts     = ["Todas las provincias"] + sorted(_prov_display.values())
            selected_prov  = st.selectbox("Provincia (zoom)", _prov_opts, index=0, key="geo_prov_select")

    geo_tipo = geo_tipo_map[geo_tipo_label]
    geo_cat  = geo_cat_map[geo_cat_label]
    gdf_base = geo[(geo["TIPO"]==geo_tipo) & (geo["CATEGORÍA"]==geo_cat)].copy()
    geo_year_prev = int(gdf_base["GEO_YEAR_PREV"].dropna().iloc[0]) if len(gdf_base) else int(year_prev_s)
    geo_year_cur  = int(gdf_base["GEO_YEAR_CUR"].dropna().iloc[0])  if len(gdf_base) else int(year_cur_s)
    gdf_base["canton_key"]  = gdf_base["CANTON"].apply(canon_canton_key)
    gdf_base["valor_2025"]  = pd.to_numeric(gdf_base["GEO_VALOR_CUR"],          errors="coerce")
    gdf_base["valor_2024"]  = pd.to_numeric(gdf_base["GEO_VALOR_PREV"],         errors="coerce")
    gdf_base["var_abs"]     = pd.to_numeric(gdf_base["VARIACION_ABSOLUTA"],     errors="coerce")
    gdf_base["yoy"]         = pd.to_numeric(gdf_base["VARIACION_PORCENTUAL"],   errors="coerce")
    if gdf_base["yoy"].dropna().abs().median() <= 2.0:
        gdf_base["yoy"] = gdf_base["yoy"] * 100.0

    total_2025 = float(pd.to_numeric(gdf_base["valor_2025"], errors="coerce").dropna().sum())
    gdf_base["share"] = np.where(total_2025 > 0, gdf_base["valor_2025"]/total_2025, np.nan)
    gdf_base = gdf_base.sort_values("valor_2025", ascending=False).reset_index(drop=True)
    gdf_base["rank"] = np.arange(1, len(gdf_base)+1)

    _prov_cantons  = _prov_mapping_raw.get("prov_cantons",{})
    _prov_centroid = _prov_mapping_raw.get("prov_centroid",{})
    _sel_prov_key  = None
    _map_center    = dict(lat=-1.5, lon=-78.5)
    _map_zoom      = 5.5
    if selected_prov != "Todas las provincias" and _prov_display:
        _sel_prov_key = next((k for k,v in _prov_display.items() if v == selected_prov), None)
    if _sel_prov_key and _sel_prov_key in _prov_centroid:
        clon, clat = _prov_centroid[_sel_prov_key]
        _map_center = dict(lat=clat, lon=clon)
        _map_zoom   = 8.0

    if HAS_GEOPANDAS and shp_canton_path and shp_canton_path.exists():
        geojson_obj = load_canton_geojson_cached(str(shp_canton_path), shp_canton_path.stat().st_mtime, 0.0045)
        if geojson_obj is None:
            with card("Mapa cantonal"):
                st.warning("No se pudo convertir el shapefile. Se muestra la tabla.")
        else:
            base_map = gdf_base.set_index("canton_key", drop=False)
            keys     = sorted(list(set(
                f["properties"].get("CANTON_NORM")
                for f in geojson_obj.get("features",[])
                if f.get("properties",{}).get("CANTON_NORM")
            )))
            df_map = pd.DataFrame({"CANTON_NORM": keys})
            for col in ["CANTON","valor_2025","valor_2024","var_abs","yoy","share","rank"]:
                df_map[col] = df_map["CANTON_NORM"].map(base_map[col].to_dict() if col in base_map.columns else {})
            
            df_map["cant_display"] = df_map["CANTON"].fillna(df_map["CANTON_NORM"])
            if _sel_prov_key and _sel_prov_key in _prov_cantons:
                df_map = df_map[df_map["CANTON_NORM"].isin(set(_prov_cantons[_sel_prov_key]))].copy()

            if map_metric.startswith(f"Variación {geo_year_cur}"):
                df_map["color_val"] = pd.to_numeric(df_map["yoy"], errors="coerce")
                scale, zmid, label  = MAP_SCALE_YOY, 0.0, "Variación (%)"
            elif "absoluta" in map_metric:
                df_map["color_val"] = pd.to_numeric(df_map.get("var_abs"), errors="coerce") / 1e6
                scale, zmid, label  = MAP_SCALE_YOY, 0.0, "Variación absoluta (MM)"
            else:
                df_map["color_val"] = pd.to_numeric(df_map["valor_2025"],errors="coerce") / 1e6
                scale, zmid, label  = MAP_SCALE_NIVEL, None, f"Ventas {geo_year_cur} (MM)"

            fig_map = build_cantonal_map_figure(geojson_obj, df_map, label, scale, zmid, height=520, value_year_label=f"Valor {geo_year_cur}")
            fig_map.update_layout(mapbox=dict(style="carto-positron", zoom=_map_zoom, center=_map_center))
            if shp_prov_path and shp_prov_path.exists():
                for tr in load_prov_outline_traces(str(shp_prov_path), shp_prov_path.stat().st_mtime):
                    fig_map.add_trace(tr)
            with card("Mapa cantonal"):
                show_plotly(fig_map, "geo_map",
                    footnote="Fuente: SRI · Elaboración: DT-CIP")
    else:
        with card("Mapa cantonal"):
            st.info("Mapa no disponible (shapefile no detectado o geopandas no instalado).")

    rank_title = f"Ranking cantonal{' — ' + selected_prov if selected_prov != 'Todas las provincias' else ''}"
    with card(f"Ranking cantonal por ventas {geo_year_cur}"):
        _rank_df = gdf_base.copy()
        if _sel_prov_key and _sel_prov_key in _prov_cantons:
            _rank_df = _rank_df[_rank_df["canton_key"].isin(set(_prov_cantons[_sel_prov_key]))].copy()
        disp = _rank_df.head(50)[["CANTON","valor_2025","yoy"]].copy()
        disp[f"Ventas {geo_year_cur} (MM)"] = (disp["valor_2025"]/1e6).round(1)
        disp["Variación (%)"]    = pd.to_numeric(disp["yoy"], errors="coerce").round(1)
        disp = disp[["CANTON", f"Ventas {geo_year_cur} (MM)", "Variación (%)"]].rename(columns={"CANTON":"Cantón"})
        st.dataframe(disp, hide_index=True, use_container_width=True, height=520)

    show_panel_footnote()


# =============================================================================
# TAB: PRECIOS / INFLACIÓN
# =============================================================================
with tab_inflacion:
    section_header("Índices de precios con variación anual")

    infl = load_inflacion(str(BASE_MADRE_FILE), base_mtime)
    p_exact = infl[_period_mask(infl["fecha"], selected_years, months)].copy()
    p_plot, p_used_fallback = _series_window_until_cutoff(
        infl, "fecha", selected_years, months, fallback_tail=24, min_unique_dates=2
    )

    if len(p_exact) == 0:
        st.warning("No hay datos de inflación en el período seleccionado.")
    else:
        def to_pct(x):
            if x is None or (isinstance(x, float) and np.isnan(x)):
                return np.nan
            x = float(x)
            return x * 100.0 if abs(x) <= 1.5 else x

        ultimo_p = p_exact.iloc[-1]
        pen_src = p_plot.dropna(subset=["fecha"]).sort_values("fecha")
        pen_p = pen_src.iloc[-2] if len(pen_src) > 1 else None

        with card(f"Último dato • {get_last_month_label(pd.Timestamp(ultimo_p['fecha']))}"):
            ip1, ip2, ip3 = st.columns(3)
            for col_widget, key_col, label, sub_text in [
                (ip1, "Consumidor", "Inflación al consumidor", "Variación mensual en pp"),
                (ip2, "Productor", "Inflación del productor", "Variación mensual en pp"),
                (ip3, "Consumo Intermedio", "Consumo intermedio", "Insumos para la producción"),
            ]:
                val = to_pct(ultimo_p.get(key_col))
                prev_val = to_pct(pen_p.get(key_col)) if pen_p is not None else np.nan
                delta = val - prev_val if pd.notna(val) and pd.notna(prev_val) else None
                with col_widget:
                    st.markdown(kpi_card(
                        label,
                        f"{val:.2f}%" if val is not None and not np.isnan(val) else "—",
                        delta_mes=f"{delta:+.2f} pp" if delta is not None and not np.isnan(delta) else None,
                        delta_mes_type=(
                            "positive" if delta and delta > 0 else
                            "negative" if delta and delta < 0 else "neutral"
                        ),
                        sub=sub_text
                    ), unsafe_allow_html=True)

        with card("Variación anual de precios por categoría"):
            fig_inf = go.Figure()
            line_mode = "lines+markers" if p_plot["fecha"].nunique() <= 2 else "lines+markers"
            infl_colors = {
                "Consumidor": BULLETIN_GOLD,
                "Productor": BULLETIN_BLUE,
                "Consumo intermedio": BULLETIN_BLUE_DARK,
            }
            last_points_inf = []
            for col_name, name in [("Consumidor", "Consumidor"), ("Productor", "Productor"),
                                   ("Consumo Intermedio", "Consumo intermedio")]:
                if col_name in p_plot.columns:
                    ser = p_plot[["fecha", col_name]].copy()
                    ser[col_name] = ser[col_name].apply(to_pct)
                    ser = ser.dropna(subset=[col_name])
                    if ser.empty:
                        continue
                    color = infl_colors.get(name, BULLETIN_SLATE)
                    fig_inf.add_trace(go.Scatter(
                        x=ser["fecha"], y=ser[col_name],
                        name=name, mode=line_mode,
                        line=dict(width=2.8, color=color),
                        marker=dict(size=6.5, color=color, line=dict(width=1.0, color="#ffffff")),
                        hovertemplate="%{x|%b-%Y}<br>%{y:.2f}%<extra></extra>"
                    ))
                    last_points_inf.append((name, pd.Timestamp(ser["fecha"].iloc[-1]), float(ser[col_name].iloc[-1]), color))

            apply_fluent_layout(
                fig_inf, height=440, hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                yaxis=dict(title="Variación anual (%)"),
                xaxis=_time_axis_for_series(p_plot, "fecha"),
                margin=dict(l=38, r=95, t=60, b=44)
            )

            yshift_map = {"Consumidor": -16, "Productor": 12, "Consumo intermedio": 34}
            for name, x_last, y_last, color in last_points_inf:
                fig_inf.add_annotation(
                    x=x_last, y=y_last,
                    text=f"{y_last:.2f}%".replace(".", ","),
                    showarrow=False, xanchor="left", align="left",
                    xshift=10, yshift=yshift_map.get(name, 0),
                    font=dict(size=10.5, color=color),
                    bgcolor="rgba(255,255,255,0.86)",
                    bordercolor="rgba(148,163,184,0.22)",
                    borderwidth=1, borderpad=4
                )

            show_plotly(
                fig_inf, "inflacion_evolucion",
                footnote=(
                    "Fuente: INEC · Elaboración: DT-CIP. "
                    + ("Se muestra la historia reciente hasta el mes de corte para preservar una serie legible."
                       if p_used_fallback else
                       "Serie mostrada dentro del rango temporal seleccionado.")
                )
            )

    inc = load_incidencias_inpp(str(INCIDENCIAS_FILE), INCIDENCIAS_FILE.stat().st_mtime)
    if not inc.empty:
        section_header("Incidencias INPP")

        categorias_inpp = sorted([
            cat for cat in inc["Descripción"].dropna().unique().tolist()
            if "variación mensual" not in str(cat).lower()
        ])

        with card("Configuración de incidencias"):
            cat_sel = st.multiselect(
                "Categorías INPP",
                options=categorias_inpp,
                default=categorias_inpp,
                key="inpp_cat_sel"
            )

        inc_filtered = inc[inc["Descripción"].isin(cat_sel if cat_sel else categorias_inpp)].copy()
        inc_f, inc_used_fallback = _series_window_until_cutoff(
            inc_filtered, "fecha", selected_years, months, fallback_tail=24, min_unique_dates=2
        )

        inc_f = inc_f[
            ~inc_f["Descripción"].astype(str).str.contains("variación mensual", case=False, na=False)
        ].copy()

        if inc_f.empty:
            st.info("No hay incidencias INPP disponibles para el período seleccionado.")
        else:
            snap_date = pd.Timestamp(inc_f["fecha"].max())
            inc_snap = inc_f[inc_f["fecha"] == snap_date].copy()
            inc_snap["Incidencia"] = pd.to_numeric(inc_snap["Incidencia"], errors="coerce")
            inc_snap = inc_snap.dropna(subset=["Incidencia"]).copy()
            inc_snap["peso_abs"] = inc_snap["Incidencia"].abs()
            inc_snap["incidencia_fmt"] = inc_snap["Incidencia"].map(lambda x: f"{x:+.2f} pp")
            inc_snap["peso_abs_fmt"] = inc_snap["peso_abs"].map(lambda x: f"{x:.2f}")
            inc_snap = inc_snap.sort_values("peso_abs", ascending=False)

            with card(f"Incidencias INPP por categoría • {get_last_month_label(snap_date)}"):
                fig_inc = px.treemap(
                    inc_snap,
                    path=["Descripción"],
                    values="peso_abs",
                    color="Incidencia",
                    color_continuous_scale=[
                        [0.0, BULLETIN_GOLD_DARK],
                        [0.5, BULLETIN_NEUTRAL],
                        [1.0, BULLETIN_BLUE]
                    ],
                    color_continuous_midpoint=0,
                    custom_data=["Incidencia", "peso_abs", "incidencia_fmt", "peso_abs_fmt"]
                )
                fig_inc.update_traces(
                    text=inc_snap["incidencia_fmt"],
                    textinfo="label+text",
                    texttemplate=(
                        "<b>%{label}</b>"
                        "<br>%{text}"
                    ),
                    textfont=dict(size=12, color="#0f172a"),
                    marker=dict(line=dict(color="#ffffff", width=2)),
                    hovertemplate=(
                        "<b>%{label}</b>"
                        "<br>Incidencia: %{customdata[2]}"
                        "<br>Peso absoluto: %{customdata[3]}"
                        "<extra></extra>"
                    ),
                    root_color="#f8fafc",
                )
                fig_inc.update_layout(
                    height=480,
                    margin=dict(l=6, r=6, t=38, b=6),
                    coloraxis_colorbar=dict(
                        title=dict(text="Incidencia (pp)", font=dict(size=11)),
                        thickness=13, len=0.76,
                        bgcolor="rgba(248,250,252,.92)",
                        bordercolor="#e2e8f0", borderwidth=1,
                        tickfont=dict(size=10),
                    ),
                    font=dict(family="DM Sans, Inter, sans-serif"),
                )
                show_plotly(
                    fig_inc, "inpp_incidencias",
                    footnote=(
                        "Fuente: INEC · Elaboración: DT-CIP. "
                        + f"Se muestra el último corte disponible dentro del rango activo ({get_last_month_label(snap_date)}). "
                        + "El tamaño del bloque representa el peso absoluto de la incidencia y el color su signo e intensidad."
                    )
                )

    show_panel_footnote()


# =============================================================================
# TAB: CONFIANZA
# =============================================================================
with tab_confianza:
    section_header("Confianza del consumidor")

    conf = load_confianza(str(BASE_MADRE_FILE), base_mtime)
    idx_exp = load_indice_expectativas(str(INDICE_EXPECT_FILE), INDICE_EXPECT_FILE.stat().st_mtime)

    with card("Filtros de confianza"):
        cf1, cf2 = st.columns([1.3, 1.2])
        with cf1:
            conf_window = st.selectbox(
                "Ventana temporal",
                ["Todo el período", "Últimos 24 meses", "Últimos 12 meses"],
                index=0, key="conf_window_select"
            )
        with cf2:
            show_expect = st.toggle("Mostrar índice de expectativas económicas", value=True, key="conf_show_expect")

    conf_exact = conf[_period_mask(conf["fecha"], selected_years, months)].copy()
    conf_f, conf_used_fallback = _series_window_until_cutoff(
        conf, "fecha", selected_years, months, fallback_tail=24, min_unique_dates=2
    )

    if conf_window != "Todo el período" and not conf_f.empty:
        n = 24 if "24" in conf_window else 12
        last_dates = list(pd.Series(conf_f["fecha"].dropna().sort_values().unique()).tail(n))
        conf_f = conf_f[conf_f["fecha"].isin(last_dates)].copy()

    if conf_exact.empty and conf_f.empty:
        st.warning("No hay información de confianza disponible para el período seleccionado.")
    elif conf_f["fecha"].nunique() < 2:
        st.warning("No hay suficiente historia para mostrar una serie legible de confianza.")
    else:
        ultimo_src = conf_exact if not conf_exact.empty else conf_f
        ultimo_c = ultimo_src.sort_values("fecha").iloc[-1]
        penultimo_c = conf_f.sort_values("fecha").iloc[-2]
        var_m = float(ultimo_c["Global"] - penultimo_c["Global"])
        var_p = float(ultimo_c["Presente"] - penultimo_c["Presente"])
        var_f = float(ultimo_c["Futuro"] - penultimo_c["Futuro"])

        idx_f = idx_exp[idx_exp["fecha"].isin(conf_f["fecha"] if len(conf_f) else idx_exp["fecha"])].copy()

        ck1, ck2, ck3 = st.columns(3, gap="small")
        for col_w, label, val, var in [
            (ck1, "ICC Global", ultimo_c["Global"], var_m),
            (ck2, "Situación presente", ultimo_c["Presente"], var_p),
            (ck3, "Expectativas futuras", ultimo_c["Futuro"], var_f),
        ]:
            with col_w:
                st.markdown(kpi_card(
                    label,
                    f"{val:.1f}" if pd.notna(val) else "—",
                    delta_mes=fmt_pts(var),
                    delta_mes_type="positive" if var > 0 else "negative" if var < 0 else "neutral",
                    sub=f"Variación mensual: {var:+.1f} pts"
                ), unsafe_allow_html=True)

        with card("Evolución mensual: confianza y expectativas de la economía"):
            fig_conf = go.Figure()
            line_mode = "lines+markers" if conf_f["fecha"].nunique() <= 2 else "lines+markers"

            conf_specs = [
                ("Global", "Global", BULLETIN_BLUE_DARK, 3.0),
                ("Presente", "Situación presente", BULLETIN_GOLD_DARK, 2.5),
                ("Futuro", "Expectativas (ICC)", BULLETIN_BLUE, 2.5),
            ]
            last_conf_points = []

            for col_name, label_name, color, width in conf_specs:
                fig_conf.add_trace(go.Scatter(
                    x=conf_f["fecha"], y=conf_f[col_name],
                    name=label_name, mode=line_mode,
                    line=dict(width=width, color=color),
                    marker=dict(size=6, color=color, line=dict(width=1.0, color="#ffffff")),
                    hovertemplate="%{x|%b-%Y}<br>%{y:.1f} pts<extra></extra>"
                ))
                ser_valid = conf_f[["fecha", col_name]].dropna()
                if len(ser_valid):
                    last_conf_points.append((label_name, pd.Timestamp(ser_valid["fecha"].iloc[-1]), float(ser_valid[col_name].iloc[-1]), color))

            if show_expect and len(idx_f):
                fig_conf.add_trace(go.Scatter(
                    x=idx_f["fecha"], y=idx_f["Indice_expectativas"],
                    name="Índice de expectativas económicas", mode=line_mode,
                    line=dict(width=2.8, color=BULLETIN_SLATE, dash="dot"),
                    marker=dict(size=5, color=BULLETIN_SLATE, line=dict(width=1.0, color="#ffffff")),
                    hovertemplate="%{x|%b-%Y}<br>%{y:.1f} pts<extra></extra>"
                ))
                ser_expect = idx_f[["fecha", "Indice_expectativas"]].dropna()
                if len(ser_expect):
                    last_conf_points.append((
                        "Índice de expectativas económicas",
                        pd.Timestamp(ser_expect["fecha"].iloc[-1]),
                        float(ser_expect["Indice_expectativas"].iloc[-1]),
                        BULLETIN_SLATE,
                    ))

            fig_conf.add_hline(y=50, line_width=1, line_dash="dash", opacity=0.35)

            apply_fluent_layout(
                fig_conf, height=450, hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                yaxis=dict(title="Puntos"),
                xaxis=_time_axis_for_series(conf_f, "fecha"),
                margin=dict(l=38, r=135, t=60, b=44)
            )

            yshift_conf = {
                "Global": -18,
                "Situación presente": 18,
                "Expectativas (ICC)": 38,
                "Índice de expectativas económicas": 58,
            }
            for label_name, x_last, y_last, color in last_conf_points:
                fig_conf.add_annotation(
                    x=x_last, y=y_last,
                    text=f"{y_last:.1f}",
                    showarrow=False, xanchor="left",
                    xshift=10, yshift=yshift_conf.get(label_name, 0),
                    font=dict(size=10.5, color=color),
                    bgcolor="rgba(255,255,255,0.86)",
                    bordercolor="rgba(148,163,184,0.22)",
                    borderwidth=1, borderpad=4
                )

            show_plotly(
                fig_conf, "confianza_evolucion",
                footnote=(
                    "Fuente: BCE · Elaboración: DT-CIP. Índice < 50: pesimismo; > 50: optimismo. "
                    + ("Se usa la historia reciente hasta el mes de corte para evitar un eje temporal sin serie."
                       if conf_used_fallback else
                       "Serie mostrada dentro de la ventana temporal seleccionada.")
                )
            )

            st.markdown(
                """
                <div class="proj-note" style="margin-top:10px;">
                  <b>Interpretación:</b> el índice de expectativas económicas resume la percepción de los hogares sobre la evolución futura de la economía. 
                  Un valor superior a 50 puntos sugiere un sesgo optimista, mientras que un valor inferior a 50 refleja expectativas predominantemente pesimistas.
                </div>
                """,
                unsafe_allow_html=True
            )

    show_panel_footnote()


# =============================================================================
# TAB: PROYECCIONES — Fanchart mejorado
# =============================================================================
with tab_proy:
    section_header("Proyecciones de variación interanual")

    proy = load_proyecciones(str(BASE_MADRE_FILE), base_mtime).copy()

    with card("Configuración"):
        serie_proy = st.selectbox("Indicador a proyectar",
            ["Ventas domésticas","Ventas gravadas IVA","Exportaciones de bienes y servicios","Ventas y exportaciones"],
            index=0, key="serie_proy_select")

    col_yoy_map = {
        "Ventas domésticas":       "Ventas domesticas.1",
        "Ventas gravadas IVA":     "Ventas gravadas.1",
        "Exportaciones de bienes y servicios": "Exportaciones.1",
        "Ventas y exportaciones":  "Total domesticas y exportadores.1",
    }
    col_sel = col_yoy_map[serie_proy]

    if col_sel not in proy.columns:
        st.error(f"No existe la columna esperada en PROYECCIONES: {col_sel}")
    else:
        proy["yoy"] = pd.to_numeric(proy[col_sel], errors="coerce") * 100.0
        hist = proy[proy["fecha"] <= last_month_sales].dropna(subset=["yoy"]).tail(24)
        fut  = proy[proy["fecha"] >  last_month_sales].dropna(subset=["yoy"])

        with card(f"Proyección de variación interanual: {serie_proy}"):
            figp = go.Figure()

            hist_plot = hist.sort_values("fecha").copy()
            fut_plot  = fut.sort_values("fecha").copy()

            # ── Señal histórica ────────────────────────────────────────────
            figp.add_trace(go.Scatter(
                x=hist_plot["fecha"], y=hist_plot["yoy"],
                mode="lines+markers", name="Histórico",
                line=dict(width=2.9, color=BULLETIN_SLATE),
                marker=dict(size=6, color=BULLETIN_SLATE,
                            line=dict(color=BULLETIN_BLUE_DARK, width=1.2)),
                hovertemplate="%{x|%b-%Y}<br><b>Histórico</b>: %{y:.1f}%<extra></extra>"
            ))

            if len(fut_plot):
                last_actual_date = hist_plot["fecha"].iloc[-1]
                last_actual_y    = float(hist_plot["yoy"].iloc[-1])
                x_proj = pd.Index([last_actual_date]).append(pd.Index(fut_plot["fecha"]))
                y_proj = np.r_[last_actual_y, fut_plot["yoy"].values.astype(float)]

                hist_recent = hist_plot["yoy"].tail(18).astype(float)
                sigma_level = float(hist_recent.std(ddof=0)) if len(hist_recent) > 1 else 0.0
                sigma_step  = float(hist_recent.diff().dropna().std(ddof=0)) if len(hist_recent) > 3 else 0.0
                base_sigma  = max(sigma_level * 0.42, sigma_step * 0.85, 0.65)
                steps       = np.arange(len(x_proj), dtype=float)
                widths_base = base_sigma * np.sqrt(steps)

                bands = [
                    (1.645, "rgba(59,130,246,0.12)", "Intervalo 90%"),
                    (1.036, "rgba(59,130,246,0.20)", "Intervalo 70%"),
                    (0.674, "rgba(59,130,246,0.30)", "Intervalo 50%"),
                ]

                x_band = list(x_proj) + list(x_proj[::-1])
                for z_value, fill_rgba, label_band in bands:
                    band_width = widths_base * z_value
                    upper = y_proj + band_width
                    lower = y_proj - band_width
                    y_band = list(upper) + list(lower[::-1])
                    figp.add_trace(go.Scatter(
                        x=x_band,
                        y=y_band,
                        mode="lines",
                        fill="toself",
                        fillcolor=fill_rgba,
                        line=dict(color="rgba(59,130,246,0)", width=0),
                        hoverinfo="skip",
                        name=label_band,
                    ))

                upper_90 = y_proj + widths_base * 1.645
                lower_90 = y_proj - widths_base * 1.645
                figp.add_trace(go.Scatter(
                    x=x_proj, y=upper_90,
                    mode="lines", name="Límite superior 90%",
                    line=dict(width=1.2, color="rgba(59,130,246,0.42)", dash="dot"),
                    hoverinfo="skip", showlegend=False
                ))
                figp.add_trace(go.Scatter(
                    x=x_proj, y=lower_90,
                    mode="lines", name="Límite inferior 90%",
                    line=dict(width=1.2, color="rgba(59,130,246,0.42)", dash="dot"),
                    hoverinfo="skip", showlegend=False
                ))

                if len(fut_plot):
                    figp.add_vrect(
                        x0=fut_plot["fecha"].iloc[0],
                        x1=fut_plot["fecha"].iloc[-1],
                        fillcolor="rgba(15,23,42,0.035)",
                        line_width=0,
                        layer="below"
                    )

                proj_text = [""] + [f"{float(v):+.1f}%".replace(".", ",") for v in fut_plot["yoy"]]
                proj_pos  = ["top center"] + [
                    ("top center" if i % 2 == 0 else "bottom center")
                    for i in range(len(fut_plot))
                ]
                figp.add_trace(go.Scatter(
                    x=x_proj, y=y_proj,
                    mode="lines+markers+text", name="Trayectoria central",
                    text=proj_text,
                    textposition=proj_pos,
                    textfont=dict(size=10.5, color=BULLETIN_BLUE_DARK),
                    line=dict(width=3.2, dash="dot", color=BULLETIN_BLUE),
                    marker=dict(size=7.5, color="#ffffff", symbol="diamond",
                                line=dict(color=BULLETIN_BLUE_DARK, width=2)),
                    cliponaxis=False,
                    hovertemplate="%{x|%b-%Y}<br><b>Proyección central</b>: %{y:.1f}%<extra></extra>"
                ))

                figp.add_shape(
                    type="line",
                    x0=last_actual_date,
                    x1=last_actual_date,
                    y0=0,
                    y1=1,
                    xref="x",
                    yref="paper",
                    line=dict(width=1.4, dash="dash", color="rgba(100,116,139,0.52)"),
                    layer="above"
                )
                figp.add_annotation(
                    x=last_actual_date,
                    y=1.02,
                    xref="x",
                    yref="paper",
                    text="Inicio de proyección",
                    showarrow=False,
                    font=dict(size=10, color="#64748b"),
                    xanchor="left",
                    yanchor="bottom",
                    bgcolor="rgba(255,255,255,0.82)",
                    bordercolor="rgba(100,116,139,0.18)",
                    borderwidth=1,
                    borderpad=4
                )

            figp.add_hline(y=0, line_width=1.15, line_dash="solid", opacity=0.28)

            apply_fluent_layout(
                figp,
                height=470,
                hovermode="x unified",
                legend=dict(
                    orientation="h", yanchor="top", y=-0.14,
                    xanchor="center", x=0.5, font=dict(size=10),
                    traceorder="reversed"
                ),
                yaxis=dict(
                    title="Variación interanual (%)",
                    showgrid=True,
                    gridcolor=colors.gray30,
                    zeroline=False
                ),
                xaxis=dict(
                    rangeslider=RANGE_SLIDER,
                    showgrid=True,
                    gridcolor=colors.gray20,
                    tickformat="%b-%Y"
                ),
                margin=dict(l=38, r=16, t=42, b=84),
                plot_bgcolor="rgba(248,250,252,0.58)"
            )
            show_plotly(figp, "proj_fanchart")

        st.markdown("""
        <div class="proj-note" style="margin-top:14px;">
          <b>Nota metodológica:</b>
          La trayectoria central se acompaña de intervalos de 50%, 70% y 90% que se ensanchan con el horizonte para hacer visible la incertidumbre proyectada.
        </div>""", unsafe_allow_html=True)

    show_panel_footnote()


# =============================================================================
# FOOTER
# =============================================================================
st.markdown(
    f"""
    <div style="text-align:center;padding:14px 0 6px 0;
                color:{colors.gray90};font-size:10.5px;
                border-top:1px solid {colors.gray40};margin-top:16px;">
      <strong>Panel Ejecutivo de Ventas</strong>
      &nbsp;·&nbsp; Fuente: SRI Ecuador
      &nbsp;·&nbsp; Dirección Técnica — CIP
      &nbsp;·&nbsp; v2.3
    </div>
    """,
    unsafe_allow_html=True
)
