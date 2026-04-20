import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path
from datetime import datetime

st.set_page_config(
    page_title="Análisis de Demanda B2B",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR  = Path(__file__).parent
RUTA_MAPA = BASE_DIR / "mapa_demanda.html"

# CSS: quitar chrome de Streamlit + badge de fecha flotante
st.markdown("""
<style>
    .block-container { padding: 0 !important; max-width: 100% !important; }
    #MainMenu, footer, header { visibility: hidden; height: 0 !important; }
    .stDeployButton { display: none; }
    [data-testid="stToolbar"] { display: none; }
    [data-testid="stDecoration"] { display: none; }
    [data-testid="stSidebar"] { display: none; }
    [data-testid="collapsedControl"] { display: none; }
    iframe { border: none !important; width: 100% !important; }
    .main .block-container { padding: 0 !important; }

    #badge-fecha {
        position: fixed;
        bottom: 10px;
        left: 50%;
        transform: translateX(-50%);
        z-index: 99999;
        background: rgba(255,255,255,0.92);
        border: 1px solid #ddd;
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 11px;
        font-family: Arial, sans-serif;
        color: #555;
        box-shadow: 0 1px 6px rgba(0,0,0,0.12);
        pointer-events: none;
        white-space: nowrap;
    }
    #badge-fecha span { color: #1565C0; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner="Cargando mapa...")
def cargar_mapa(ruta: str, mtime: float) -> str:
    with open(ruta, "r", encoding="utf-8") as f:
        return f.read()

if not RUTA_MAPA.exists():
    st.error("⚠️ El mapa aún no ha sido generado.")
    st.stop()

mtime     = RUTA_MAPA.stat().st_mtime
fecha_str = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")

# Badge flotante centrado abajo — fuera del iframe del mapa
st.markdown(
    f'<div id="badge-fecha">📡 Datos actualizados al <span>{fecha_str}</span></div>',
    unsafe_allow_html=True,
)

html_content = cargar_mapa(str(RUTA_MAPA), mtime)
components.html(html_content, height=900, scrolling=False)
