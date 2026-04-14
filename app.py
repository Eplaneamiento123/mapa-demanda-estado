import streamlit as st
import streamlit.components.v1 as components
import os

st.set_page_config(page_title="Mapa Demanda Fibra", layout="wide")

ruta = "mapa_demanda.html"

if os.path.exists(ruta):
    with open(ruta, "r", encoding="utf-8") as f:
        html_content = f.read()
    components.html(html_content, height=800, scrolling=True)
else:
    st.warning("El mapa aún no ha sido generado. Ejecuta el proceso de actualización.")