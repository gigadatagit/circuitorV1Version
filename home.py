import streamlit as st
import login as login

archivo = __file__.split("/")[-1]
login.generarLogin(archivo)
if 'usuario' in st.session_state:
    st.header('Página :orange[principal]')
    
    st.markdown("""
    # 🚀 **:orange[Plataforma Integral de Gestión de Datos - Circuitor VX]**
    Bienvenido a **Plataforma Integral de Gestión de Datos - Circuitor VX**, la herramienta diseñada para **automatizar**, **transformar** y **visualizar** información de manera eficiente. 📊✨  

    ### 🔑 **Características principales**
    - **📥 Procesamiento avanzado de datos:** Carga archivos en formato TXT y optimiza la transformación y gestión de información.
    - **📄 Generación automatizada de informes:** Convierte datos en documentos Word y Excel estructurados, reduciendo tiempos de elaboración a solo **3 minutos**.
    - **📊 Visualización dinámica:** Explora gráficos interactivos con Plotly para analizar tendencias y tomar decisiones informadas en tiempo real.

    ### 🎯 **Nuestro objetivo**
    Ofrecerte una solución **todo-en-uno** que simplifique la administración de datos, mejore el análisis y potencie la productividad.  

    #### 🔍 **¡Transforma tus datos en conocimiento con eficiencia y precisión!**
    """)
