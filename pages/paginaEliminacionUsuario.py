import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import io
import plotly.graph_objects as go
import requests
import zipfile
import login as login

archivo = __file__.split("/")[-1]
login.generarLogin(archivo)
if 'correo_electronico' in st.session_state:
    st.header('Información | :orange[Página de Eliminación de Usuario]')
    
    uploaded_file = st.file_uploader("Elige un archivo de .TXT (Minuto a Minuto)", type=["txt"])
    uploaded_file2 = st.file_uploader("Elige un archivo de .TXT (Hora a Hora)", type=["txt"])