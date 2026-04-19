import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path

st.set_page_config(
    page_title="Análisis de Demanda B2B",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# CSS: quitar todo el chrome de Streamlit para que el mapa ocupe toda la pantalla
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
</style>
""", unsafe_allow_html=True)

BASE_DIR  = Path(__file__).parent
RUTA_MAPA = BASE_DIR / "mapa_demanda.html"

@st.cache_data(show_spinner="Cargando mapa...")
def cargar_mapa(ruta: str, mtime: float) -> str:
    with open(ruta, "r", encoding="utf-8") as f:
        return f.read()

if not RUTA_MAPA.exists():
    st.error("⚠️ El mapa aún no ha sido generado.")
    st.stop()

mtime = RUTA_MAPA.stat().st_mtime
html_content = cargar_mapa(str(RUTA_MAPA), mtime)

components.html(html_content, height=900, scrolling=False)
