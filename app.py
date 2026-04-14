import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Mapa Demanda Fibra", layout="wide")

with open("index.html", "r", encoding="utf-8") as f:
    html_content = f.read()

components.html(html_content, height=800, scrolling=True)