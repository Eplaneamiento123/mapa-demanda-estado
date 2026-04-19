import streamlit as st
import streamlit.components.v1 as components
from datetime import datetime
from pathlib import Path

# ==========================================
# CONFIGURACIÓN
# ==========================================
st.set_page_config(
    page_title="Análisis de Demanda B2B - Fibra Óptica",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Quitar márgenes/paddings de Streamlit para que el mapa use toda la pantalla
st.markdown("""
<style>
    .block-container { padding-top: 1rem; padding-bottom: 0; max-width: 100%; }
    #MainMenu, footer, header { visibility: hidden; }
    .stDeployButton { display: none; }
    iframe { border: none !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# RUTAS (mismas carpetas donde el script genera los archivos)
# ==========================================
BASE_DIR = Path(__file__).parent
RUTA_MAPA  = BASE_DIR / "mapa_demanda.html"
RUTA_EXCEL = BASE_DIR / "resultados_demanda.xlsx"

# ==========================================
# UTILIDADES CON CACHÉ
# ==========================================
@st.cache_data(show_spinner=False)
def cargar_mapa(ruta: str, mtime: float) -> str:
    """Lee el HTML del mapa. El mtime como parámetro invalida la caché si el archivo cambia."""
    with open(ruta, "r", encoding="utf-8") as f:
        return f.read()

@st.cache_data(show_spinner=False)
def cargar_excel_bytes(ruta: str, mtime: float) -> bytes:
    """Lee el Excel en bytes para el botón de descarga."""
    with open(ruta, "rb") as f:
        return f.read()

def formatear_fecha(ts: float) -> str:
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M")

# ==========================================
# SIDEBAR CON INFO
# ==========================================
with st.sidebar:
    st.title("📡 Demanda B2B")
    st.caption("Análisis de fibra óptica - Mapa interactivo")

    st.divider()
    st.subheader("Estado de archivos")

    if RUTA_MAPA.exists():
        mtime_mapa = RUTA_MAPA.stat().st_mtime
        tamaño_mb  = RUTA_MAPA.stat().st_size / (1024 * 1024)
        st.success(f"✅ Mapa disponible")
        st.caption(f"📅 Generado: {formatear_fecha(mtime_mapa)}")
        st.caption(f"📦 Tamaño: {tamaño_mb:.1f} MB")
    else:
        st.error("❌ Mapa no generado")

    if RUTA_EXCEL.exists():
        mtime_excel = RUTA_EXCEL.stat().st_mtime
        st.success(f"✅ Reporte Excel disponible")
        st.caption(f"📅 Generado: {formatear_fecha(mtime_excel)}")

        excel_bytes = cargar_excel_bytes(str(RUTA_EXCEL), mtime_excel)
        st.download_button(
            label="⬇️ Descargar Excel",
            data=excel_bytes,
            file_name=f"resultados_demanda_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning("⚠️ Excel no disponible")

    st.divider()
    st.caption(
        "💡 Use el control de capas (arriba a la derecha) para mostrar/ocultar entidades, "
        "nodos, tipo de fondo, etc."
    )

# ==========================================
# CUERPO: MAPA
# ==========================================
if not RUTA_MAPA.exists():
    st.warning("⚠️ El mapa aún no ha sido generado.")
    st.info("Ejecuta el script `P_DEMANDA_ESTADO.py` con la VPN corporativa conectada para actualizar el mapa y el reporte.")
    st.stop()

mtime_mapa = RUTA_MAPA.stat().st_mtime

with st.spinner("Cargando mapa interactivo…"):
    html_content = cargar_mapa(str(RUTA_MAPA), mtime_mapa)

# Altura dinámica: ocupa casi toda la ventana
components.html(html_content, height=900, scrolling=False)
