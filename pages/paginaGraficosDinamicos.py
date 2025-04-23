import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import io
import plotly.graph_objects as go
import requests
import zipfile
import login as login
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from docx.shared import Mm
from io import BytesIO
from utilities import organizar_DataFrame_M_a_M, organizar_DataFrame_H_a_H, calcular_Valor_Tension_Nominal, calcular_Valor_Corriente_Nominal, filtrar_DataFrame_Por_Columnas, crear_DataFrame_Desbalance_Tension, crear_DataFrame_Desbalance_Corriente, crear_DataFrame_PQS_Potencias, crear_DataFrame_FactPotencia, crear_DataFrame_FactPotenciaGrupos, crear_DataFrame_DistTension, crear_DataFrame_Armonicos_DistTension, crear_DataFrame_DistCorriente, crear_DataFrame_Armonicos_DistCorriente, crear_DataFrame_Armonicos_CargabilidadTDD, crear_DataFrame_FactorK_Final, crear_Medidas_DataFrame_Tension, crear_Medidas_DataFrame_DesbTension, crear_Medidas_DataFrame_Corriente, crear_Medidas_DataFrame_DesbCorriente, crear_Medidas_DataFrame_PQS, crear_Medidas_DataFrame_FactorPotencia, crear_Medidas_DataFrame_FactorPotenciaGeneral, crear_Medidas_DataFrame_Distorsion_Tension, crear_Medidas_DataFrame_Armonicos_DistTension, crear_Medidas_DataFrame_Distorsion_Corriente, crear_Medidas_DataFrame_Armonicos_DistCorriente, crear_Medidas_DataFrame_FactorK, calcular_Valor_Corriente_Cortacircuito, calcular_Valor_ISC_entre_IL, calcular_Valor_Limite_TDD, calcular_Valores_Limites_Armonicos, crear_DataFrame_CargabilidadTDD_Final, crear_Medidas_DataFrame_CargabilidadTDD, crear_Medidas_DataFrame_Energias, crear_DataFrame_Energias, crear_DataFrame_Tension, crear_DataFrame_Corriente, calcular_Variacion_Tension, calcular_Valor_Cargabilidad_Disponibilidad, calcular_Observacion_Tension, calcular_Observacion_Corriente, calcular_Observacion_DesbTension, calcular_Observacion_DesbCorriente, calcular_Observacion_THDV, calcular_Observacion_Armonicos_Corriente, calcular_Observacion_TDD, graficar_Timeline_Tension, graficar_Timeline_Corriente, graficar_Timeline_DesbTension, graficar_Timeline_DesbCorriente, graficar_Timeline_PQS_ActApa, graficar_Timeline_PQS_CapInd, graficar_Timeline_FactPotencia, graficar_Timeline_Distorsion_Tension, graficar_Timeline_Distorsion_Corriente, graficar_Timeline_CargabilidadTDD, graficar_Timeline_FactorK, generar_Graficos_Barras_Energias, graficar_Timeline_Tension_Plotly, graficar_Timeline_Corriente_Plotly, graficar_Timeline_DesbTension_Plotly, graficar_Timeline_DesbCorriente_Plotly, graficar_Timeline_PQS_ActApa_Plotly, graficar_Timeline_PQS_CapInd_Plotly, graficar_Timeline_FactPotencia_Plotly, graficar_Timeline_Distorsion_Tension_Plotly, graficar_Timeline_Distorsion_Corriente_Plotly, graficar_Timeline_CargabilidadTDD_Plotly, graficar_Timeline_FactorK_Plotly, generar_Graficos_Barras_Energias_Plotly, crear_grafico, generar_Graficos_Barras_Energias2

archivo = __file__.split("/")[-1]
login.generarLogin(archivo)
if 'usuario' in st.session_state:
    st.header('Información | :orange[Página de Gráficos Dinámicos]')
    
    uploaded_file = st.file_uploader("Elige un archivo de .TXT (Minuto a Minuto)", type=["txt"])
    uploaded_file2 = st.file_uploader("Elige un archivo de .TXT (Hora a Hora)", type=["txt"])
    
    if uploaded_file and uploaded_file2:
        
        st.success("Archivos subidos correctamente.")
        
        try:
            
            #st.markdown("""
            #---
            #
            #> ## Elige la plantilla para generar el informe.
            #
            #---
            #""")

            #plantillaSeleccionada = st.selectbox("Selecciona una Plantilla:", ["Vatia", "GIGA"])
            
            #st.markdown("""
            #---
            #
            #> ## Elige si vas a visualizar o no las energías generadas.
            #
            #---
            #""")
            
            #energiaGenerada = st.selectbox("Seleccione si quiere visualizar o no la Energía Generada:", ["Sí", "No"], index=1)
            
            #st.markdown("""
            #---
            #
            #> ## Ingresa el valor de cada variable y dale click al botón para generar el informe.
            #
            #---
            #""")
            
            st.markdown("""
            ---
            
            > ## Ingresa el valor de cada variable y dale click al botón para visualizar los gráficos dinámicos.
            
            ---
            """)
            
            var1 = st.number_input("Ingrese el Valor Nominal de Tensión:", min_value=0.0, max_value=1000.0, step=0.1, format="%.1f")
            var2 = st.number_input("Ingrese el Valor de la Capacidad del Transformador:", min_value=0.0, max_value=1000000.0, step=0.1, format="%.1f")
            var3 = st.number_input("Ingrese el Valor de Referencia - Desbalance de Tensión:", min_value=0.0, max_value=2.0, step=0.1, format="%.1f")
            var4 = st.number_input("Ingrese el Valor de Referencia - Desbalance de Corriente:", min_value=0.0, max_value=20.0, step=0.1, format="%.1f")
            var5 = st.number_input("Ingrese el Valor de Límite Máximo de Distorsión Armónica de Tensión:", min_value=0.0, max_value=1000.0, step=0.1, format="%.1f")
            var6 = st.number_input("Ingrese el Valor de Impedancia de Cortocircuito (Transformador):", min_value=0.0, max_value=1000.0, step=0.1, format="%.1f")
            var7 = st.number_input("Ingrese el Valor de Referencia - PLT (Flicker):", min_value=0.0, max_value=5.0, step=0.1, format="%.1f")
            
            st.markdown("""
            ---
            
            > ## Visualización de Gráfico Interactivo para Comparación de Variables.
        
            ---
            """)
            
            df_Read_Comparacion = pd.read_csv(uploaded_file, delimiter=';', encoding="UTF-8-SIG", encoding_errors='ignore')
            #st.dataframe(df_Read.head(5))
                    
            df_Comparacion = organizar_DataFrame_M_a_M(df_Read_Comparacion)
            #st.dataframe(df.head(5))
                    
            crear_grafico(df_Comparacion.copy())
            
            if st.button("Generar Gráficos Dinámicos", type="primary"):
                
                try:

                    var_Limite_Inferior_Tension = calcular_Valor_Tension_Nominal(var1)[0]
                    var_Limite_Superior_Tension = calcular_Valor_Tension_Nominal(var1)[1]
                                
                    print(f"Limites de Tensión - Inferior ({var_Limite_Inferior_Tension}) y Superior({var_Limite_Superior_Tension})")

                    var_Corriente_Nominal_Value = calcular_Valor_Corriente_Nominal((var2 * 1000), var1)
                        
                    #df = pd.read_parquet(uploaded_file)
                    #df_Read = pd.read_csv(uploaded_file, delimiter=';', encoding="UTF-8-SIG", encoding_errors='ignore')
                    #st.dataframe(df_Read.head(5))
                        
                    #df = organizar_DataFrame_M_a_M(df_Read)
                    df = df_Comparacion.copy()
                    #st.dataframe(df.head(5))
                        
                    print("¿Quedan valores NaN en el DataFrame de Minuto a Minuto?", df.isna().any().any())
                    print(df.head(5))
                    print(df.index)  # ¿Es continuo? ¿Está vacío?
                    print(df.shape)  # ¿Tiene filas y columnas?
                        
                    df_Energias_Read = pd.read_csv(uploaded_file2, delimiter=';', encoding="UTF-8-SIG", encoding_errors='ignore')
                    #st.dataframe(df_Energias_Read.head(5))

                    df_Energias = organizar_DataFrame_H_a_H(df_Energias_Read)
                    #st.dataframe(df_Energias.head(5))
                        
                    print("¿Quedan valores NaN en el DataFrame de Hora a Hora?", df_Energias.isna().any().any())
                    print(df_Energias.head(5))
                    print(df_Energias.index)  # ¿Es continuo? ¿Está vacío?
                    print(df_Energias.shape)  # ¿Tiene filas y columnas?
                        
                        
                    st.markdown("""
                    > ## Creando DataFrame - Circuitor            
                    """)
                        
                    #if plantillaSeleccionada == "Vatia":
                            
                        #print("a")
                    var_Enlace_Plantilla = "https://github.com/gigadatagit/GIGA_Data/blob/365a61d9e72f3e175c39d5fa6cb1c189e0c70ffa/vars_Template_ETV_Circuitor_VATIA.docx?raw=true"
                        
                    #elif plantillaSeleccionada == "GIGA":
                            
                        #print("a")
                        #var_Enlace_Plantilla = "https://github.com/gigadatagit/GIGA_Data/blob/365a61d9e72f3e175c39d5fa6cb1c189e0c70ffa/vars_Template_ETV_Circuitor.docx?raw=true"
                        
                    #else:
                            
                        #print("Por favor seleccione una plantilla válida.")
                            
                    # Enlace a la Plantilla del Documento de Word que contiene toda la información del Informe
                    url = var_Enlace_Plantilla

                    # Petición para Traer la información de esa URL con la Plantilla
                    response = requests.get(url)

                    # Guardado de contenido de la Plantilla de Word en un el Almacenamiento de Memoria
                    template_data = BytesIO(response.content)

                    # Crear una instancia de DocxTemplate - Carga el contenido de la Plantilla del Documento de Word
                    doc = DocxTemplate(template_data)

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para la tensión
                    list_Columns_Grafico_Tension: list = ['Tensin L12', 'Tensin L23', 'Tensin L31']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Desbalance de Tensión
                    list_Columns_Grafico_DesbTension: list = ['Desbalance']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para la corriente
                    list_Columns_Grafico_Corriente: list = ['Corriente mx. L1', 'Corriente mx. L2', 'Corriente mx. L3', 'Corriente de neutro']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Desbalance de Tensión
                    list_Columns_Grafico_DesbCorriente: list = ['Desbalance']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el PQS - N1
                    list_Columns_Grafico_DesbCorriente_ActApa: list = ['P.Activa III', 'P.Aparente III']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el PQS - N2
                    list_Columns_Grafico_DesbCorriente_CapInd: list = ['P.Capacitiva III', 'P.Inductiva III']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Factor de Potencia
                    list_Columns_Grafico_FactorPot: list = ['F.P. III -', 'F.P. III']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Análisis de Energías
                    list_Columns_Graficos_Consolidado_Energia: list = ['E.Activa T1', 'E.Capacitiva T1', 'E.Inductiva T1', 'KVARH_CAP', 'KARH_IND']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Análisis de Distorsión de Tensión
                    list_Columns_Distorsion_Tension: list = ['V THD/d Mx. L1', 'V THD/d Mx. L2', 'V THD/d Mx. L3']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Análisis de Distorsión de Corriente
                    list_Columns_Distorsion_Corriente: list = ['A THD/d L1', 'A THD/d L2', 'A THD/d L3']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Análisis del Listado de Armónicos de Cargabilidad TDD
                    list_Columns_Armonicos_Cargabilidad_TDD: list = ['resultado_TDD_L1', 'resultado_TDD_L2', 'resultado_TDD_L3']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Análisis del Flicker
                    list_Columns_Flicker: list = ['Plt L1', 'Plt L2', 'Plt L3']

                    # Aquí tenemos una lista de las columnas que se van a graficar a través del tiempo para el Análisis del Factor K
                    list_Columns_FactorK: list = ['Factor K L1', 'Factor K L2', 'Factor K L3']



                    # Declaración de todos los DataFrames filtrando por las columnas que se van a Utilizar para generar el Documento y Realizar los Cálculos o Gráficos

                    df_Tabla_Tension = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'Tensin mn. L12', 'Tensin L12', 'Tensin mx. L12', 'Tensin mn. L23', 'Tensin L23', 'Tensin mx. L23', 'Tensin mn. L31', 'Tensin L31', 'Tensin mx. L31'], df)

                    df_Tabla_Corriente = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'Corriente mn. L1', 'Corriente L1', 'Corriente mx. L1', 'Corriente mn. L2', 'Corriente L2', 'Corriente mx. L2', 'Corriente mn. L3', 'Corriente L3', 'Corriente mx. L3', 'Corriente de neutro mn.', 'Corriente de neutro', 'Corriente de neutro mx.'], df)

                    df_Tabla_Desbalance_Tension = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'Tensin L12', 'Tensin L23', 'Tensin L31'], df)

                    df_Tabla_Desbalance_Corriente = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'Corriente L1', 'Corriente L2', 'Corriente L3'], df)

                    df_Tabla_PQS_Potencias = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'P.Activa mn. III', 'P.Activa III', 'P.Activa mx. III', 'P.Capacitiva mn. III', 'P.Capacitiva III', 'P.Capacitiva mx. III', 'P.Inductiva mn. III', 'P.Inductiva III', 'P.Inductiva mx. III', 'P.Aparente mn. III', 'P.Aparente III', 'P.Aparente mx. III'], df)

                    df_Tabla_FactorPotencia = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'F.P. Mn. III -', 'F.P. III -', 'F.P. Mx. III -', 'F.P. Mn. III', 'F.P. III', 'F.P. Mx. III'], df)

                    df_Tabla_Distorsion_Tension = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'V THD/d Mx. L1', 'V THD/d Mx. L2', 'V THD/d Mx. L3'], df)

                    df_Tabla_Armonicos_Distorsion_Tension = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'Arm. tensin 3 L1', 'Arm. tensin 5 L1', 'Arm. tensin 7 L1', 'Arm. tensin 9 L1', 'Arm. tensin 11 L1', 'Arm. tensin 13 L1', 'Arm. tensin 15 L1', 'Arm. tensin 3 L2', 'Arm. tensin 5 L2', 'Arm. tensin 7 L2', 'Arm. tensin 9 L2', 'Arm. tensin 11 L2', 'Arm. tensin 13 L2', 'Arm. tensin 15 L2', 'Arm. tensin 3 L3', 'Arm. tensin 5 L3', 'Arm. tensin 7 L3', 'Arm. tensin 9 L3', 'Arm. tensin 11 L3', 'Arm. tensin 13 L3', 'Arm. tensin 15 L3'], df)

                    df_Tabla_Distorsion_Corriente = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'A THD/d L1', 'A THD/d L2', 'A THD/d L3'], df)

                    df_Tabla_Armonicos_Distorsion_Corriente = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'Arm. corriente 3 L1', 'Arm. corriente 5 L1', 'Arm. corriente 7 L1', 'Arm. corriente 9 L1', 'Arm. corriente 11 L1', 'Arm. corriente 13 L1', 'Arm. corriente 15 L1', 'Arm. corriente 3 L2', 'Arm. corriente 5 L2', 'Arm. corriente 7 L2', 'Arm. corriente 9 L2', 'Arm. corriente 11 L2', 'Arm. corriente 13 L2', 'Arm. corriente 15 L2', 'Arm. corriente 3 L3', 'Arm. corriente 5 L3', 'Arm. corriente 7 L3', 'Arm. corriente 9 L3', 'Arm. corriente 11 L3', 'Arm. corriente 13 L3', 'Arm. corriente 15 L3'], df)

                    df_Tabla_Armonicos_Cargabilidad_TDD = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'A THD/d L1', 'A THD/d L2', 'A THD/d L3', 'Corriente mx. L1', 'Corriente mx. L2', 'Corriente mx. L3'], df)

                    #df_Tabla_Flicker = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'Plt L1', 'Plt L2', 'Plt L3'], df)

                    df_Tabla_FactorK = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'Factor K mn. L1', 'Factor K L1', 'Factor K mx. L1', 'Factor K mn. L2', 'Factor K L2', 'Factor K mx. L2', 'Factor K mn. L3', 'Factor K L3', 'Factor K mx. L3'], df)

                    df_Tabla_FactorPotencia_Grupos = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'F.P. Mn. III', 'F.P. III', 'F.P. Mx. III'], df)

                        
                        
                        
                    # En este paso se realizan los pasos adicionales como cálculos de nuevas columnas u operaciones entre columnas
                        
                        
                    var_Tabla_Tensiones = crear_DataFrame_Tension(dataFrame=df_Tabla_Tension, var_Lim_Inf_Ten=var_Limite_Inferior_Tension, val_Nom=var1, var_Lim_Sup_Ten=var_Limite_Superior_Tension)

                    st.markdown("""
                    > ## Cabecera - DataFrame de Tensión Final
                    """)

                    st.dataframe(var_Tabla_Tensiones.head(5))

                        
                    var_Tabla_Corrientes = crear_DataFrame_Corriente(dataFrame=df_Tabla_Corriente, var_Lim_Corr_Nom=var_Corriente_Nominal_Value)

                    st.markdown("""
                    > ## Cabecera - DataFrame de Corriente Final
                    """)

                    st.dataframe(var_Tabla_Corrientes.head(5))
                        

                    df_Tabla_Desb_Tension = crear_DataFrame_Desbalance_Tension(df_Tabla_Desbalance_Tension, var3)
                        
                    st.markdown("""
                    > ## Cabecera - DataFrame de Desbalance de Tensión Final
                    """)

                    st.dataframe(df_Tabla_Desb_Tension.head(5))


                    df_Tabla_Desb_Corriente = crear_DataFrame_Desbalance_Corriente(df_Tabla_Desbalance_Corriente, var4)
                        
                    st.markdown("""
                    > ## Cabecera - DataFrame de Desbalance de Corriente Final
                    """)

                    st.dataframe(df_Tabla_Desb_Corriente.head(5))
                        

                    df_Tabla_PQS_Final = crear_DataFrame_PQS_Potencias(df_Tabla_PQS_Potencias)
                        
                    st.markdown("""
                    > ## Cabecera - DataFrame de Potencias Final
                    """)

                    st.dataframe(df_Tabla_PQS_Final.head(5))
                        

                    df_Tabla_FactPotenciaFinal = crear_DataFrame_FactPotencia(df_Tabla_FactorPotencia)
                        
                    st.markdown("""
                    > ## Cabecera - DataFrame de Factor de Potencia Final
                    """)

                    st.dataframe(df_Tabla_FactPotenciaFinal.head(5))
                        

                    df_Tabla_FactorPotencia_GruposFinal = crear_DataFrame_FactPotenciaGrupos(df_Tabla_FactorPotencia_Grupos)
                        
                    #st.markdown("""
                    #> ## Cabecera - DataFrame de Factor de Potencia Grupos (Capacitivo/Inductivo) Final
                    #""")

                    #st.dataframe(df_Tabla_FactorPotencia_GruposFinal.head(5))
                        

                    df_Tabla_Distorsion_TensionFinal = crear_DataFrame_DistTension(df_Tabla_Distorsion_Tension, var5)
                        
                    st.markdown("""
                    > ## Cabecera - DataFrame de Distorsión de Tensión Final
                    """)

                    st.dataframe(df_Tabla_Distorsion_TensionFinal.head(5))
                        

                    df_Tabla_Armonicos_Distorsion_Tension_Final = crear_DataFrame_Armonicos_DistTension(df_Tabla_Armonicos_Distorsion_Tension)

                    st.markdown("""
                    > ## Cabecera - DataFrame de Armónicos de Distorsión de Tensión Final
                    """)

                    st.dataframe(df_Tabla_Armonicos_Distorsion_Tension_Final.head(5))
                        
                        
                    df_Tabla_Distorsion_CorrienteFinal = crear_DataFrame_DistCorriente(df_Tabla_Distorsion_Corriente)
                        
                    st.markdown("""
                    > ## Cabecera - DataFrame de Distorsión de Corriente Final
                    """)

                    st.dataframe(df_Tabla_Distorsion_CorrienteFinal.head(5))
                        

                    df_Tabla_Armonicos_Distorsion_Corriente_Final = crear_DataFrame_Armonicos_DistCorriente(df_Tabla_Armonicos_Distorsion_Corriente)

                    st.markdown("""
                    > ## Cabecera - DataFrame de Armónicos de Distorsión de Corriente Final
                    """)

                    st.dataframe(df_Tabla_Armonicos_Distorsion_Corriente_Final.head(5))
                        

                    df_Tabla_Armonicos_Cargabilidad_TDDFinal = crear_DataFrame_Armonicos_CargabilidadTDD(df_Tabla_Armonicos_Cargabilidad_TDD)

                    st.markdown("""
                    > ## Cabecera - DataFrame de Armónicos de Cargabilidad TDD Final
                    """)

                    st.dataframe(df_Tabla_Armonicos_Cargabilidad_TDDFinal.head(5))
                        

                    #df_Tabla_FlickerFinal = crear_DataFrame_Flicker_Final(df_Tabla_Flicker, var7)

                    df_Tabla_FactorKFinal = crear_DataFrame_FactorK_Final(df_Tabla_FactorK)
                        
                    st.markdown("""
                    > ## Cabecera - DataFrame de FactorK Final
                    """)

                    st.dataframe(df_Tabla_FactorKFinal.head(5))
                        
                        
                        
                    # En este paso se están realizando los cálculos de las tablas con Percentiles, Máximos, Promedios y Mínimos.

                    df_Tabla_Calculos_Tension = crear_Medidas_DataFrame_Tension(df_Tabla_Tension)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de Tensión
                    """)

                    st.dataframe(df_Tabla_Calculos_Tension)
                        

                    df_Tabla_Calculos_Desb_Tension = crear_Medidas_DataFrame_DesbTension(df_Tabla_Desb_Tension)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de Desbalance de Tensión
                    """)

                    st.dataframe(df_Tabla_Calculos_Desb_Tension)
                        

                    df_Tabla_Calculos_Corriente = crear_Medidas_DataFrame_Corriente(df_Tabla_Corriente)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de Corriente
                    """)

                    st.dataframe(df_Tabla_Calculos_Corriente)
                        

                    df_Tabla_Calculos_Desb_Corriente = crear_Medidas_DataFrame_DesbCorriente(df_Tabla_Desb_Corriente)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de Desbalance de Corriente
                    """)

                    st.dataframe(df_Tabla_Calculos_Desb_Corriente)
                        

                    df_Tabla_Calculos_PQS_Potencias = crear_Medidas_DataFrame_PQS(df_Tabla_PQS_Final)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de Potencias
                    """)

                    st.dataframe(df_Tabla_Calculos_PQS_Potencias)
                        

                    df_Tabla_Calculos_FactorPotencia = crear_Medidas_DataFrame_FactorPotencia(df_Tabla_FactPotenciaFinal)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de Factor de Potencia
                    """)

                    st.dataframe(df_Tabla_Calculos_FactorPotencia)
                        

                    df_Tabla_Calculos_FactorPotenciaGeneral = crear_Medidas_DataFrame_FactorPotenciaGeneral(df_Tabla_FactorPotencia_GruposFinal)

                    #st.markdown("""
                    #> ## Medidas - DataFrame de Factor de Potencia (Generado/Consumido)
                    #""")

                    #st.dataframe(df_Tabla_Calculos_FactorPotenciaGeneral)
                        

                    df_Tabla_Calculos_DistTension = crear_Medidas_DataFrame_Distorsion_Tension(df_Tabla_Distorsion_TensionFinal)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de Distorsión de Tensión
                    """)

                    st.dataframe(df_Tabla_Calculos_DistTension)
                        

                    df_Tabla_Calculos_Armonicos_DistTension = crear_Medidas_DataFrame_Armonicos_DistTension(df_Tabla_Armonicos_Distorsion_Tension_Final)

                    st.markdown("""
                    > ## Medidas - DataFrame de Armónicos de Distorsión de Tensión
                    """)

                    st.dataframe(df_Tabla_Calculos_Armonicos_DistTension)
                        

                    df_Tabla_Calculos_DistCorriente = crear_Medidas_DataFrame_Distorsion_Corriente(df_Tabla_Distorsion_CorrienteFinal)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de Distorsión de Corriente
                    """)

                    st.dataframe(df_Tabla_Calculos_DistCorriente)
                        

                    df_Tabla_Calculos_Armonicos_DistCorriente = crear_Medidas_DataFrame_Armonicos_DistCorriente(df_Tabla_Armonicos_Distorsion_Corriente_Final)

                    st.markdown("""
                    > ## Medidas - DataFrame de Armónicos de Distorsión de Corriente
                    """)

                    st.dataframe(df_Tabla_Calculos_Armonicos_DistCorriente)
                        

                    #df_Tabla_Calculos_Flicker = crear_Medidas_DataFrame_Flicker(df_Tabla_FlickerFinal)

                    df_Tabla_Calculos_FactorK = crear_Medidas_DataFrame_FactorK(df_Tabla_FactorKFinal)
                        
                    st.markdown("""
                    > ## Medidas - DataFrame de FactorK
                    """)

                    st.dataframe(df_Tabla_Calculos_FactorK)
                        
                        
                        
                    # Impresión de los resultados del Factor de Potencia de Tipo Inductivo y Tipo Capacitivo

                    print('--'*30)

                    print(f'DataFrame - Factor de Potencia {df_Tabla_FactorPotencia_Grupos}')

                    print(f'DataFrame - Factor de Potencia en Grupo Final {df_Tabla_FactorPotencia_GruposFinal}')

                    print(f'Diccionario - Medidas de Factor de Potencia {df_Tabla_Calculos_FactorPotenciaGeneral}')

                    print('--'*30)



                    # Separamos esta sección ya que es importante distinguir el uso del DataFrame que está compuesto por los datos del TDD Final y poder hacer los cálculos correspondientes

                    print('--'*30)

                    valor_Maximo_Corrientes = df[list_Columns_Grafico_Corriente[0:3]].max().max()

                    print(f"Valor Máximo de de las Corrientes: {valor_Maximo_Corrientes}")

                    valor_Corriente_Cortacircuito = calcular_Valor_Corriente_Cortacircuito(var_Corriente_Nominal_Value, var6)

                    print(f"Valor de Corriente Cortacircuito {valor_Corriente_Cortacircuito}")

                    valor_ISC_sobre_IL = calcular_Valor_ISC_entre_IL(valor_Corriente_Cortacircuito, valor_Maximo_Corrientes)

                    print(f"Valor de ISC/IL {valor_ISC_sobre_IL}")

                    valor_Limite_TDD: float = calcular_Valor_Limite_TDD(valor_ISC_sobre_IL)

                    print(f"Valor del Limite del TDD {valor_Limite_TDD}")

                    valores_Limites_Armonicos = calcular_Valores_Limites_Armonicos(valor_Limite_TDD)

                    print(f"Valores de los Límites de los Armónicos {valores_Limites_Armonicos.values()}")

                    print('--'*30)

                    df_Tabla_TDD = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'fecha_y_Hora', 'resultado_TDD_L1', 'resultado_TDD_L2', 'resultado_TDD_L3'], df_Tabla_Armonicos_Cargabilidad_TDDFinal)

                    df_Tabla_TDDFinal = crear_DataFrame_CargabilidadTDD_Final(df_Tabla_TDD, valor_Limite_TDD)

                    df_Tabla_Calculos_CargabilidadTDD = crear_Medidas_DataFrame_CargabilidadTDD(df_Tabla_TDDFinal)



                    # Separamos esta sección ya que es importante distinguir el uso del DataFrame que está compuesto por los datos del .TXT que va de Hora a Hora

                    df_Tabla_Energias = filtrar_DataFrame_Por_Columnas(['Fecha/hora', 'E.Activa T1', 'E.Capacitiva T1', 'E.Inductiva T1', 'KWH', 'KARH_IND', 'KVARH_CAP', 'F.P. III -', 'F.P. III'], crear_DataFrame_Energias(df_Energias))

                    df_Tabla_Calculos_Energias = crear_Medidas_DataFrame_Energias(df_Tabla_Energias)

                    # Convertimos la información del DataFrame que contiene las energías para luego convertirlo en un diccionario con los registros de cada una de las columnas y poder mostrarlos en una tabla de Word

                    table_Data_Energy_Info = df_Tabla_Energias.to_dict(orient="records")



                    # Separamos esta sección ya que es importante distinguir el uso del DataFrame del Factor de Potencia, para aplicarle Filtros de Medición a los Datos

                    filtro_FP_POS_CANTPOS = (df_Tabla_FactPotenciaFinal['F.P. III'] > 0)

                    filtro_FP_POS_CANTZeros = (df_Tabla_FactPotenciaFinal['F.P. III'] == abs(0))

                    filtro_FP_POS_CANTNEG = (df_Tabla_FactPotenciaFinal['F.P. III'] < 0)

                    filtro_FP_NEG_CANTPOS = (df_Tabla_FactPotenciaFinal['F.P. III -'] > 0)

                    filtro_FP_NEG_CANTZeros = (df_Tabla_FactPotenciaFinal['F.P. III -'] == abs(0))

                    filtro_FP_NEG_CANTNEG = (df_Tabla_FactPotenciaFinal['F.P. III -'] < 0)



                    # En este lugar declaramos un diccionario con los Valores negativos, ceros y positivos del Factor de Potencia

                    data_Cantidad_NEG_POS_FactorPotencia: dict = {
                        'CANT_POSITIVOS_FP_POS': len(df_Tabla_FactPotenciaFinal[filtro_FP_POS_CANTPOS]),
                        'CANT_CEROS_FP_POS': len(df_Tabla_FactPotenciaFinal[filtro_FP_POS_CANTZeros]),
                        'CANT_NEGATIVOS_FP_POS': len(df_Tabla_FactPotenciaFinal[filtro_FP_POS_CANTNEG]),
                        'CANT_POSITIVOS_FP_NEG': len(df_Tabla_FactPotenciaFinal[filtro_FP_NEG_CANTPOS]),
                        'CANT_CEROS_FP_NEG': len(df_Tabla_FactPotenciaFinal[filtro_FP_NEG_CANTZeros]),
                        'CANT_NEGATIVOS_FP_NEG': len(df_Tabla_FactPotenciaFinal[filtro_FP_NEG_CANTNEG])
                    }



                    # En este lugar declaramos diccionarios con los percentiles para utilizarlos luego en gráficos o en otras partes del código

                    data_Percentiles_Tension: dict = {
                        'PERCENTIL_TENSIN_L12': round(df_Tabla_Calculos_Tension['Tensin L12'].iloc[0], 2),
                        'PERCENTIL_TENSIN_L23': round(df_Tabla_Calculos_Tension['Tensin L23'].iloc[0], 2),
                        'PERCENTIL_TENSIN_L31': round(df_Tabla_Calculos_Tension['Tensin L31'].iloc[0], 2)
                    }

                    data_Percentiles_Corriente: dict = {
                        'PERCENTIL_CORR_MAX_L1': round(df_Tabla_Calculos_Corriente['Corriente mx. L1'].iloc[0], 2),
                        'PERCENTIL_CORR_MAX_L2': round(df_Tabla_Calculos_Corriente['Corriente mx. L2'].iloc[0], 2),
                        'PERCENTIL_CORR_MAX_L3': round(df_Tabla_Calculos_Corriente['Corriente mx. L3'].iloc[0], 2),
                        'PERCENTIL_CORR_MED_LN': round(df_Tabla_Calculos_Corriente['Corriente de neutro'].iloc[0], 2)
                    }

                    data_Percentiles_Corriente_Maximos: dict = {
                        'PERCENTIL_CORR_MAX_L1': round(df_Tabla_Calculos_Corriente['Corriente mx. L1'].iloc[0], 2),
                        'PERCENTIL_CORR_MAX_L2': round(df_Tabla_Calculos_Corriente['Corriente mx. L2'].iloc[0], 2),
                        'PERCENTIL_CORR_MAX_L3': round(df_Tabla_Calculos_Corriente['Corriente mx. L3'].iloc[0], 2)
                    }

                    data_Percentiles_DesbTension: dict = {
                        'PERCENTIL_DESBALANCE_DESBTEN': round(df_Tabla_Calculos_Desb_Tension['Desbalance'].iloc[0], 2)
                    }

                    data_Percentiles_DesbCorriente: dict = {
                        'PERCENTIL_DESBALANCE_DESBCORR': round(df_Tabla_Calculos_Desb_Corriente['Desbalance'].iloc[0], 2)
                    }

                    data_Percentiles_PQS_ActApa: dict = {
                        'PERCENTIL_PQS_ACT': round(df_Tabla_Calculos_PQS_Potencias['P.Activa III'].iloc[0], 2),
                        'PERCENTIL_PQS_APA': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente III'].iloc[0], 2)
                    }

                    data_Percentiles_PQS_CapInd: dict = {
                        'PERCENTIL_PQS_CAP': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva III'].iloc[0], 2),
                        'PERCENTIL_PQS_IND': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva III'].iloc[0], 2)
                    }

                    data_Percentiles_FactorPotencia: dict = {
                        'PERCENTIL_FACTOR_POTENCIA_NEG': round(df_Tabla_Calculos_FactorPotencia['F.P. III -'].iloc[0], 2),
                        'PERCENTIL_FACTOR_POTENCIA_POS': round(df_Tabla_Calculos_FactorPotencia['F.P. III'].iloc[0], 2)
                    }

                    data_Percentiles_Energia: dict = {
                        'PERCENTIL_ENERGIA_ACTIVA_MED': round(df_Tabla_Calculos_Energias['E.Activa T1'].iloc[0], 2),
                        'PERCENTIL_ENERGIA_CAPACITIVA_MED': round(df_Tabla_Calculos_Energias['E.Capacitiva T1'].iloc[0], 2),
                        'PERCENTIL_ENERGIA_INDUCTIVA_MED': round(df_Tabla_Calculos_Energias['E.Inductiva T1'].iloc[0], 2)
                    }

                    data_Percentiles_DistorsionTension: dict = {
                        'PERCENTIL_THDV_MAX_L1': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L1'].iloc[0],2),
                        'PERCENTIL_THDV_MAX_L2': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L2'].iloc[0],2),
                        'PERCENTIL_THDV_MAX_L3': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L3'].iloc[0],2)
                    }

                    data_Percentiles_DistorsionCorriente: dict = {
                        'PERCENTIL_THDI_MAX_L1': round(df_Tabla_Calculos_DistCorriente['A THD/d L1'].iloc[0],2),
                        'PERCENTIL_THDI_MAX_L2': round(df_Tabla_Calculos_DistCorriente['A THD/d L2'].iloc[0],2),
                        'PERCENTIL_THDI_MAX_L3': round(df_Tabla_Calculos_DistCorriente['A THD/d L3'].iloc[0],2)
                    }

                    data_Percentiles_CargabilidadTDD: dict = {
                        'PERCENTIL_TDD_L1': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L1'].iloc[0],2),
                        'PERCENTIL_TDD_L2': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L2'].iloc[0],2),
                        'PERCENTIL_TDD_L3': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L3'].iloc[0],2)
                    }

                    #"""
                    #
                    #data_Percentiles_Flicker: dict = {
                    #    'PERCENTIL_FLICKER_PLT_L1_MED': round(df_Tabla_Calculos_Flicker['Plt L1'].iloc[0],2),
                    #    'PERCENTIL_FLICKER_PLT_L2_MED': round(df_Tabla_Calculos_Flicker['Plt L2'].iloc[0],2),
                    #    'PERCENTIL_FLICKER_PLT_L3_MED': round(df_Tabla_Calculos_Flicker['Plt L3'].iloc[0],2)
                    #}
                    #
                    #"""

                    data_Percentiles_FactorK: dict = {
                        'PERCENTIL_FACTORK_L1_MED': round(df_Tabla_Calculos_FactorK['Factor K L1'].iloc[0], 2),
                        'PERCENTIL_FACTORK_L2_MED': round(df_Tabla_Calculos_FactorK['Factor K L2'].iloc[0], 2),
                        'PERCENTIL_FACTORK_L3_MED': round(df_Tabla_Calculos_FactorK['Factor K L3'].iloc[0], 2)
                    }
                        
                        
                        
                    # Creación del código que nos permite tener todos los DataFrames que estamos utilizando en su versión final, convirtiéndolos a un Excel que contiene distintas hojas
                    # En estas hojas veremos en una hoja con todas las columnas de los DataFrames y de resto, hojas individuales que contienen la información de cada uno de ellos (Minuto a Minuto)

                    # Creamos una copia de cada uno de los DataFrames Finales

                    df_Tabla_Tension_Copy = var_Tabla_Tensiones.copy()

                    df_Tabla_Desb_Tension_Copy = df_Tabla_Desb_Tension.copy()

                    df_Tabla_Corriente_Copy = var_Tabla_Corrientes.copy()

                    df_Tabla_Desb_Corriente_Copy = df_Tabla_Desb_Corriente.copy()

                    df_Tabla_PQS_Final_Copy = df_Tabla_PQS_Final.copy()

                    df_Tabla_FactPotenciaFinal_Copy = df_Tabla_FactPotenciaFinal.copy()

                    df_Tabla_Distorsion_TensionFinal_Copy = df_Tabla_Distorsion_TensionFinal.copy()

                    df_Tabla_Armonicos_Distorsion_Tension_Final_Copy = df_Tabla_Armonicos_Distorsion_Tension_Final.copy()

                    df_Tabla_Distorsion_CorrienteFinal_Copy = df_Tabla_Distorsion_CorrienteFinal.copy()

                    df_Tabla_Armonicos_Distorsion_Corriente_Final_Copy = df_Tabla_Armonicos_Distorsion_Corriente_Final.copy()

                    df_Tabla_Armonicos_Cargabilidad_TDDFinal_Copy = df_Tabla_Armonicos_Cargabilidad_TDDFinal.copy()

                    df_Tabla_TDDFinal_Copy = df_Tabla_TDDFinal.copy()

                    #df_Tabla_FlickerFinal_Copy = df_Tabla_FlickerFinal.copy()

                    df_Tabla_FactorKFinal_Copy = df_Tabla_FactorKFinal.copy()

                    df_Tabla_Energias_Copy = df_Tabla_Energias.copy()

                    # Lista de DataFrames a combinar
                    listado_DataFrames: list = [df_Tabla_Tension_Copy, df_Tabla_Desb_Tension_Copy, df_Tabla_Corriente_Copy, df_Tabla_Desb_Corriente_Copy, df_Tabla_PQS_Final_Copy, df_Tabla_FactPotenciaFinal_Copy, df_Tabla_Distorsion_TensionFinal_Copy, df_Tabla_Armonicos_Distorsion_Tension_Final_Copy, df_Tabla_Distorsion_CorrienteFinal_Copy, df_Tabla_Armonicos_Distorsion_Corriente_Final_Copy, df_Tabla_Armonicos_Cargabilidad_TDDFinal_Copy, df_Tabla_TDDFinal_Copy, df_Tabla_FactorKFinal_Copy, df_Tabla_Energias_Copy]

                    print("Generando Excel con la Información de todas las columnas analizadas.")

                    # Aquí hay una lista que almacena cada uno de los valores de la Variación para cada Percentil de las Tensiones

                    var_Lista_Variaciones = calcular_Variacion_Tension(lista_Percentiles=[df_Tabla_Calculos_Tension['Tensin mn. L12'].iloc[0], df_Tabla_Calculos_Tension['Tensin mn. L23'].iloc[0], df_Tabla_Calculos_Tension['Tensin mn. L31'].iloc[0], df_Tabla_Calculos_Tension['Tensin mx. L12'].iloc[0], df_Tabla_Calculos_Tension['Tensin mx. L23'].iloc[0], df_Tabla_Calculos_Tension['Tensin mx. L31'].iloc[0]], val_Nom=var1)

                    var_Lista_PQS_Carg_Disp = [calcular_Valor_Cargabilidad_Disponibilidad(var2, df_Tabla_Calculos_PQS_Potencias['P.Aparente mx. III'].iloc[0])[0], calcular_Valor_Cargabilidad_Disponibilidad(var2, df_Tabla_Calculos_PQS_Potencias['P.Aparente mx. III'].iloc[0])[1]]

                    print(f"Listado de Variaciones: {var_Lista_Variaciones}")



                    # Aquí vamos a determinar los resultados de cada una de las Observaciones

                    print('--'*30)

                    listado_Percentiles_Tension: list = [round(df_Tabla_Calculos_Tension['Tensin mn. L12'].iloc[0], 2), round(df_Tabla_Calculos_Tension['Tensin L12'].iloc[0], 2), round(df_Tabla_Calculos_Tension['Tensin mx. L12'].iloc[0], 2), round(df_Tabla_Calculos_Tension['Tensin mn. L23'].iloc[0], 2), round(df_Tabla_Calculos_Tension['Tensin L23'].iloc[0], 2), round(df_Tabla_Calculos_Tension['Tensin mx. L23'].iloc[0], 2), round(df_Tabla_Calculos_Tension['Tensin mn. L31'].iloc[0], 2), round(df_Tabla_Calculos_Tension['Tensin L31'].iloc[0], 2), round(df_Tabla_Calculos_Tension['Tensin mx. L31'].iloc[0], 2),]

                    listado_Limites_Tension: list = [var_Limite_Inferior_Tension, var_Limite_Superior_Tension]

                    observaciones_Tension = calcular_Observacion_Tension(listado_Percentiles_Tension, listado_Limites_Tension)

                    print(f"Observaciones de Tensión: {observaciones_Tension}")

                    diccionario_Percentiles_Corriente: dict = {
                        'CORRIENTE_L1_MIN': round(df_Tabla_Calculos_Corriente['Corriente mn. L1'].iloc[0], 2),
                        'CORRIENTE_L1_MED': round(df_Tabla_Calculos_Corriente['Corriente L1'].iloc[0], 2),
                        'CORRIENTE_L1_MAX': round(df_Tabla_Calculos_Corriente['Corriente mx. L1'].iloc[0], 2),
                        'CORRIENTE_L2_MIN': round(df_Tabla_Calculos_Corriente['Corriente mn. L2'].iloc[0], 2),
                        'CORRIENTE_L2_MED': round(df_Tabla_Calculos_Corriente['Corriente L2'].iloc[0], 2),
                        'CORRIENTE_L2_MAX': round(df_Tabla_Calculos_Corriente['Corriente mx. L2'].iloc[0], 2),
                        'CORRIENTE_L3_MIN': round(df_Tabla_Calculos_Corriente['Corriente mn. L3'].iloc[0], 2),
                        'CORRIENTE_L3_MED': round(df_Tabla_Calculos_Corriente['Corriente L3'].iloc[0], 2),
                        'CORRIENTE_L3_MAX': round(df_Tabla_Calculos_Corriente['Corriente mx. L3'].iloc[0], 2)
                    }

                    diccionario_Percentiles_CorrienteNeutra: dict = {
                        'CORRIENTE_NEUTRA_MIN': round(df_Tabla_Calculos_Corriente['Corriente de neutro mn.'].iloc[0], 2),
                        'CORRIENTE_NEUTRA_MED': round(df_Tabla_Calculos_Corriente['Corriente de neutro'].iloc[0], 2),
                        'CORRIENTE_NEUTRA_MAX': round(df_Tabla_Calculos_Corriente['Corriente de neutro mx.'].iloc[0], 2)
                    }

                    valor_Corriente_Nominal = var_Corriente_Nominal_Value

                    observaciones_Corriente = calcular_Observacion_Corriente(diccionario_Percentiles_Corriente, diccionario_Percentiles_CorrienteNeutra, valor_Corriente_Nominal)

                    print(f"Observaciones de Corriente: {observaciones_Corriente}")

                    valor_Percentil_DesbTension = round(df_Tabla_Calculos_Desb_Tension['Desbalance'].iloc[0], 2)

                    valor_Referencia_DesbTension = var3

                    observaciones_DesbTension = calcular_Observacion_DesbTension(valor_Percentil_DesbTension, valor_Referencia_DesbTension)

                    print(f"Observaciones del Desbalance de Tensión: {observaciones_DesbTension}")

                    valor_Percentil_DesbCorriente = round(df_Tabla_Calculos_Desb_Corriente['Desbalance'].iloc[0], 2)

                    valor_Referencia_DesbCorriente = var4

                    observaciones_DesbCorriente = calcular_Observacion_DesbCorriente(valor_Percentil_DesbCorriente, valor_Referencia_DesbCorriente)

                    print(f"Observaciones del Desbalance de Corriente: {observaciones_DesbCorriente}")

                    diccionario_Percentiles_THDV: dict = {
                        'THDV_DISTTENSION_L1': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L1'].iloc[0], 2),
                        'THDV_DISTTENSION_L2': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L2'].iloc[0], 2),
                        'THDV_DISTTENSION_L3': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L3'].iloc[0], 2)
                    }

                    valor_Referencia_THDV = var5

                    observaciones_THDV = calcular_Observacion_THDV(diccionario_Percentiles_THDV, valor_Referencia_THDV)

                    print(f"Observaciones del THDV: {observaciones_THDV}")

                    diccionario_Percentiles_Armonicos_3_9: dict = {
                        'ARMONICO_3_L1': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L1'].iloc[0], 2),
                        'ARMONICO_3_L2': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L2'].iloc[0], 2),
                        'ARMONICO_3_L3': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L3'].iloc[0], 2),
                        'ARMONICO_5_L1': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L1'].iloc[0], 2),
                        'ARMONICO_5_L2': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L2'].iloc[0], 2),
                        'ARMONICO_5_L3': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L3'].iloc[0], 2),
                        'ARMONICO_7_L1': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L1'].iloc[0], 2),
                        'ARMONICO_7_L2': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L2'].iloc[0], 2),
                        'ARMONICO_7_L3': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L3'].iloc[0], 2),
                        'ARMONICO_9_L1': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L1'].iloc[0], 2),
                        'ARMONICO_9_L2': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L2'].iloc[0], 2),
                        'ARMONICO_9_L3': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L3'].iloc[0], 2)
                    }

                    diccionario_Percentiles_Armonicos_11: dict = {
                        'ARMONICO_11_L1': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L1'].iloc[0], 2),
                        'ARMONICO_11_L2': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L2'].iloc[0], 2),
                        'ARMONICO_11_L3': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L3'].iloc[0], 2)
                    }

                    listado_Limites_Armonicos_Corriente: list = list(valores_Limites_Armonicos.values())[:2]

                    observaciones_ArmonicosCorriente = calcular_Observacion_Armonicos_Corriente(diccionario_Percentiles_Armonicos_3_9, diccionario_Percentiles_Armonicos_11, listado_Limites_Armonicos_Corriente)

                    print(f"Listado de Límites de los Armónicos de Corriente: {listado_Limites_Armonicos_Corriente}")

                    print(f"Observaciones de los Armónicos de Corriente: {observaciones_ArmonicosCorriente}")

                    diccionario_Percentiles_TDD: dict = {
                        'TDD_PERCENTIL_L1': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L1'].iloc[0], 2),
                        'TDD_PERCENTIL_L2': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L2'].iloc[0], 2),
                        'TDD_PERCENTIL_L3': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L3'].iloc[0], 2)
                    }

                    valor_Referencia_TDD = valor_Limite_TDD

                    observaciones_TDD = calcular_Observacion_TDD(diccionario_Percentiles_TDD, valor_Referencia_TDD)

                    print(f"Observaciones del TDD: {observaciones_TDD}")

                    print('--'*30)
                    
                    
                    print('--'*30)

                    # Buffer de la Imagen para la Línea de Tiempo de la Tensión (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_Tension = graficar_Timeline_Tension(var_Tabla_Tensiones, list_Columns_Grafico_Tension, data_Percentiles_Tension, 'fecha_y_Hora', limites=[var_Tabla_Tensiones['var_Limite_Inferior_Tension'].iloc[0], var_Tabla_Tensiones['var_Limite_Superior_Tension'].iloc[0]], titulo='REGISTROS DE TENSIÓN')

                    # Buffer de la Imagen para la Línea de Tiempo de la Corriente (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_Corriente = graficar_Timeline_Corriente(var_Tabla_Corrientes, list_Columns_Grafico_Corriente, data_Percentiles_Corriente, 'fecha_y_Hora', limite=var_Tabla_Corrientes['var_Limite_Corriente_Nominal'].iloc[0], titulo='REGISTROS DE CORRIENTE')

                    # Buffer de la Imagen para la Línea de Tiempo del Desbalance de Tensión (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_DesbTension = graficar_Timeline_DesbTension(df_Tabla_Desb_Tension, list_Columns_Grafico_DesbTension, data_Percentiles_DesbTension, 'fecha_y_Hora', limite=df_Tabla_Desb_Tension['var_Ref_Desbalance_Tension'].iloc[0], titulo='REGISTROS DESBALANCE DE TENSIÓN')

                    # Buffer de la Imagen para la Línea de Tiempo del Desbalance de Tensión (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_DesbCorriente = graficar_Timeline_DesbCorriente(df_Tabla_Desb_Corriente, list_Columns_Grafico_DesbCorriente, data_Percentiles_DesbCorriente, 'fecha_y_Hora', limite=df_Tabla_Desb_Corriente['var_Ref_Desbalance_Corriente'].iloc[0], titulo='REGISTROS DESBALANCE DE CORRIENTE')

                    # Buffer de la Imagen para la Línea de Tiempo del PQS - Activa Aparente (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_PQS_ActApa = graficar_Timeline_PQS_ActApa(df_Tabla_PQS_Final, list_Columns_Grafico_DesbCorriente_ActApa, data_Percentiles_PQS_ActApa, 'fecha_y_Hora', titulo='REGISTROS DE POTENCIA - Activa / Aparente (kW / kVA)')

                    # Buffer de la Imagen para la Línea de Tiempo del PQS - Capacitiva Inductiva (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_PQS_CapInd = graficar_Timeline_PQS_CapInd(df_Tabla_PQS_Final, list_Columns_Grafico_DesbCorriente_CapInd, data_Percentiles_PQS_CapInd, 'fecha_y_Hora', titulo='REGISTROS DE POTENCIA - Capacitiva / Inductiva (kVAR)')

                    # Buffer de la Imagen para la Línea de Tiempo del Factor de Potencia (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_FactorPotencia = graficar_Timeline_FactPotencia(df_Tabla_FactPotenciaFinal, list_Columns_Grafico_FactorPot, data_Percentiles_FactorPotencia, data_Cantidad_NEG_POS_FactorPotencia, 'fecha_y_Hora', titulo='REGISTROS DE POTENCIA - Factor de Potencia')

                    # Buffer de la Imagen para la Línea de Tiempo de la Distorsión de la Tensión (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_DistTension = graficar_Timeline_Distorsion_Tension(df_Tabla_Distorsion_TensionFinal, list_Columns_Distorsion_Tension, data_Percentiles_DistorsionTension, 'fecha_y_Hora', limite=df_Tabla_Distorsion_TensionFinal['var_Ref_Distorsion_Tension'].iloc[0], titulo='REGISTROS DISTORSIÓN ARMÓNICA DE TENSIÓN - THDV')

                    # Buffer de la Imagen para la Línea de Tiempo de la Distorsión de la Corriente (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_DistCorriente = graficar_Timeline_Distorsion_Corriente(df_Tabla_Distorsion_CorrienteFinal, list_Columns_Distorsion_Corriente, data_Percentiles_DistorsionCorriente, 'fecha_y_Hora', limite=None, titulo='REGISTROS DISTORSIÓN ARMÓNICA DE CORRIENTE - THDI')

                    # Buffer de la Imagen para la Línea de Tiempo de la Cargabilidad de TDD (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_CargabilidadTDD = graficar_Timeline_CargabilidadTDD(df_Tabla_TDDFinal, list_Columns_Armonicos_Cargabilidad_TDD, data_Percentiles_CargabilidadTDD, 'fecha_y_Hora', limite=valor_Limite_TDD, titulo='REGISTROS DISTORSIÓN TOTAL DE DEMANDA')

                    # Buffer de la Imagen para la Línea de Tiempo del Flicker (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_Flicker = graficar_Timeline_Flicker(df_Tabla_FlickerFinal, list_Columns_Flicker, data_Percentiles_Flicker, 'fecha_y_Hora', limite=var7, titulo='REGISTRO DE FLICKER')

                    # Buffer de la Imagen para la Línea de Tiempo del FactorK (Aquí se almacena el gráfico en la memoria local)
                    #img_buffer_Timeline_FactorK = graficar_Timeline_FactorK(df_Tabla_FactorKFinal, list_Columns_FactorK, data_Percentiles_FactorK, 'fecha_y_Hora', limite=None, titulo='REGISTROS DE FACTOR K')



                    # Agregar datos y el gráfico al contexto
                    #img_Timeline_Tension = InlineImage(doc, img_buffer_Timeline_Tension, Cm(18))
                    #img_Timeline_Corriente = InlineImage(doc, img_buffer_Timeline_Corriente, Cm(18))
                    #img_Timeline_DesbTension = InlineImage(doc, img_buffer_Timeline_DesbTension, Cm(18))
                    #img_Timeline_DesbCorriente = InlineImage(doc, img_buffer_Timeline_DesbCorriente, Cm(18))
                    #img_Timeline_PQS_ActInd = InlineImage(doc, img_buffer_Timeline_PQS_ActApa, Cm(18))
                    #img_Timeline_PQS_CapApa = InlineImage(doc, img_buffer_Timeline_PQS_CapInd, Cm(18))
                    #img_Timeline_FactPotencia = InlineImage(doc, img_buffer_Timeline_FactorPotencia, Cm(18))
                    #img_Timeline_DistorsionTension = InlineImage(doc, img_buffer_Timeline_DistTension, Cm(18))
                    #img_Timeline_DistorsionCorriente = InlineImage(doc, img_buffer_Timeline_DistCorriente, Cm(18))
                    #img_Timeline_CargabilidadTDD = InlineImage(doc, img_buffer_Timeline_CargabilidadTDD, Cm(18))
                    #img_Timeline_Flicker = InlineImage(doc, img_buffer_Timeline_Flicker, Cm(18))
                    #img_Timeline_FactorK = InlineImage(doc, img_buffer_Timeline_FactorK, Cm(18))
                    
                    # Contexto básico que recibe el documento de Word (Se accede a él usando el nombre de la llave del diccionario)
                    registro = {
                        'var_Lim_Inf_Tension': round(var_Limite_Inferior_Tension, 2),
                        'var_Nominal_Value': round(var1, 2),
                        'var_Lim_Sup_Tension': round(var_Limite_Superior_Tension, 2),
                        'var_Cap_Trafo': round(var2, 2),
                        'var_Corr_Nominal_Value': round(var_Corriente_Nominal_Value, 2),
                        #'imagen_Linea_Tiempo_Tension': img_Timeline_Tension,
                        #'imagen_Linea_Tiempo_Corriente': img_Timeline_Corriente,
                        #'imagen_Linea_Tiempo_DesbTension': img_Timeline_DesbTension,
                        #'imagen_Linea_Tiempo_DesbCorriente': img_Timeline_DesbCorriente,
                        #'imagen_Linea_Tiempo_PQS_ActApa': img_Timeline_PQS_ActInd,
                        #'imagen_Linea_Tiempo_PQS_CapInd': img_Timeline_PQS_CapApa,
                        #'imagen_Linea_Tiempo_FactorPotencia': img_Timeline_FactPotencia,
                        'graficos_Barras_Energias': generar_Graficos_Barras_Energias2(dataFrame=df_Tabla_Energias, variables=list_Columns_Graficos_Consolidado_Energia, percentiles=data_Percentiles_Energia, fecha_col='Fecha/hora', doc=doc),
                        #'imagen_Linea_Tiempo_DistTension': img_Timeline_DistorsionTension,
                        #'imagen_Linea_Tiempo_DistCorriente': img_Timeline_DistorsionCorriente,
                        #'imagen_Linea_Tiempo_CargTDD': img_Timeline_CargabilidadTDD,
                        #'imagen_Linea_Tiempo_Flicker': img_Timeline_Flicker,
                        #'imagen_Linea_Tiempo_FactorK': img_Timeline_FactorK,
                        'table_Data_Energy': table_Data_Energy_Info,
                        'L12_MIN_PR': round(df_Tabla_Calculos_Tension['Tensin mn. L12'].iloc[0], 2),
                        'L12_MED_PR': round(df_Tabla_Calculos_Tension['Tensin L12'].iloc[0], 2),
                        'L12_MAX_PR': round(df_Tabla_Calculos_Tension['Tensin mx. L12'].iloc[0], 2),
                        'L23_MIN_PR': round(df_Tabla_Calculos_Tension['Tensin mn. L23'].iloc[0], 2),
                        'L23_MED_PR': round(df_Tabla_Calculos_Tension['Tensin L23'].iloc[0], 2),
                        'L23_MAX_PR': round(df_Tabla_Calculos_Tension['Tensin mx. L23'].iloc[0], 2),
                        'L31_MIN_PR': round(df_Tabla_Calculos_Tension['Tensin mn. L31'].iloc[0], 2),
                        'L31_MED_PR': round(df_Tabla_Calculos_Tension['Tensin L31'].iloc[0], 2),
                        'L31_MAX_PR': round(df_Tabla_Calculos_Tension['Tensin mx. L31'].iloc[0], 2),
                        'L1_CORR_MIN_PR': round(df_Tabla_Calculos_Corriente['Corriente mn. L1'].iloc[0], 2),
                        'L1_CORR_MED_PR': round(df_Tabla_Calculos_Corriente['Corriente L1'].iloc[0], 2),
                        'L1_CORR_MAX_PR': round(df_Tabla_Calculos_Corriente['Corriente mx. L1'].iloc[0], 2),
                        'L2_CORR_MIN_PR': round(df_Tabla_Calculos_Corriente['Corriente mn. L2'].iloc[0], 2),
                        'L2_CORR_MED_PR': round(df_Tabla_Calculos_Corriente['Corriente L2'].iloc[0], 2),
                        'L2_CORR_MAX_PR': round(df_Tabla_Calculos_Corriente['Corriente mx. L2'].iloc[0], 2),
                        'L3_CORR_MIN_PR': round(df_Tabla_Calculos_Corriente['Corriente mn. L3'].iloc[0], 2),
                        'L3_CORR_MED_PR': round(df_Tabla_Calculos_Corriente['Corriente L3'].iloc[0], 2),
                        'L3_CORR_MAX_PR': round(df_Tabla_Calculos_Corriente['Corriente mx. L3'].iloc[0], 2),
                        'LN_CORR_MIN_PR': round(df_Tabla_Calculos_Corriente['Corriente de neutro mn.'].iloc[0], 2),
                        'LN_CORR_MED_PR': round(df_Tabla_Calculos_Corriente['Corriente de neutro'].iloc[0], 2),
                        'LN_CORR_MAX_PR': round(df_Tabla_Calculos_Corriente['Corriente de neutro mx.'].iloc[0], 2),
                        'L12_MIN_MX': round(df_Tabla_Calculos_Tension['Tensin mn. L12'].iloc[3], 2),
                        'L12_MED_MX': round(df_Tabla_Calculos_Tension['Tensin L12'].iloc[3], 2),
                        'L12_MAX_MX': round(df_Tabla_Calculos_Tension['Tensin mx. L12'].iloc[3], 2),
                        'L23_MIN_MX': round(df_Tabla_Calculos_Tension['Tensin mn. L23'].iloc[3], 2),
                        'L23_MED_MX': round(df_Tabla_Calculos_Tension['Tensin L23'].iloc[3], 2),
                        'L23_MAX_MX': round(df_Tabla_Calculos_Tension['Tensin mx. L23'].iloc[3], 2),
                        'L31_MIN_MX': round(df_Tabla_Calculos_Tension['Tensin mn. L31'].iloc[3], 2),
                        'L31_MED_MX': round(df_Tabla_Calculos_Tension['Tensin L31'].iloc[3], 2),
                        'L31_MAX_MX': round(df_Tabla_Calculos_Tension['Tensin mx. L31'].iloc[3], 2),
                        'L1_CORR_MIN_MX': round(df_Tabla_Calculos_Corriente['Corriente mn. L1'].iloc[3], 2),
                        'L1_CORR_MED_MX': round(df_Tabla_Calculos_Corriente['Corriente L1'].iloc[3], 2),
                        'L1_CORR_MAX_MX': round(df_Tabla_Calculos_Corriente['Corriente mx. L1'].iloc[3], 2),
                        'L2_CORR_MIN_MX': round(df_Tabla_Calculos_Corriente['Corriente mn. L2'].iloc[3], 2),
                        'L2_CORR_MED_MX': round(df_Tabla_Calculos_Corriente['Corriente L2'].iloc[3], 2),
                        'L2_CORR_MAX_MX': round(df_Tabla_Calculos_Corriente['Corriente mx. L2'].iloc[3], 2),
                        'L3_CORR_MIN_MX': round(df_Tabla_Calculos_Corriente['Corriente mn. L3'].iloc[3], 2),
                        'L3_CORR_MED_MX': round(df_Tabla_Calculos_Corriente['Corriente L3'].iloc[3], 2),
                        'L3_CORR_MAX_MX': round(df_Tabla_Calculos_Corriente['Corriente mx. L3'].iloc[3], 2),
                        'LN_CORR_MIN_MX': round(df_Tabla_Calculos_Corriente['Corriente de neutro mn.'].iloc[3], 2),
                        'LN_CORR_MED_MX': round(df_Tabla_Calculos_Corriente['Corriente de neutro'].iloc[3], 2),
                        'LN_CORR_MAX_MX': round(df_Tabla_Calculos_Corriente['Corriente de neutro mx.'].iloc[3], 2),
                        'L12_MIN_PM': round(df_Tabla_Calculos_Tension['Tensin mn. L12'].iloc[1], 2),
                        'L12_MED_PM': round(df_Tabla_Calculos_Tension['Tensin L12'].iloc[1], 2),
                        'L12_MAX_PM': round(df_Tabla_Calculos_Tension['Tensin mx. L12'].iloc[1], 2),
                        'L23_MIN_PM': round(df_Tabla_Calculos_Tension['Tensin mn. L23'].iloc[1], 2),
                        'L23_MED_PM': round(df_Tabla_Calculos_Tension['Tensin L23'].iloc[1], 2),
                        'L23_MAX_PM': round(df_Tabla_Calculos_Tension['Tensin mx. L23'].iloc[1], 2),
                        'L31_MIN_PM': round(df_Tabla_Calculos_Tension['Tensin mn. L31'].iloc[1], 2),
                        'L31_MED_PM': round(df_Tabla_Calculos_Tension['Tensin L31'].iloc[1], 2),
                        'L31_MAX_PM': round(df_Tabla_Calculos_Tension['Tensin mx. L31'].iloc[1], 2),
                        'L1_CORR_MIN_PM': round(df_Tabla_Calculos_Corriente['Corriente mn. L1'].iloc[1], 2),
                        'L1_CORR_MED_PM': round(df_Tabla_Calculos_Corriente['Corriente L1'].iloc[1], 2),
                        'L1_CORR_MAX_PM': round(df_Tabla_Calculos_Corriente['Corriente mx. L1'].iloc[1], 2),
                        'L2_CORR_MIN_PM': round(df_Tabla_Calculos_Corriente['Corriente mn. L2'].iloc[1], 2),
                        'L2_CORR_MED_PM': round(df_Tabla_Calculos_Corriente['Corriente L2'].iloc[1], 2),
                        'L2_CORR_MAX_PM': round(df_Tabla_Calculos_Corriente['Corriente mx. L2'].iloc[1], 2),
                        'L3_CORR_MIN_PM': round(df_Tabla_Calculos_Corriente['Corriente mn. L3'].iloc[1], 2),
                        'L3_CORR_MED_PM': round(df_Tabla_Calculos_Corriente['Corriente L3'].iloc[1], 2),
                        'L3_CORR_MAX_PM': round(df_Tabla_Calculos_Corriente['Corriente mx. L3'].iloc[1], 2),
                        'LN_CORR_MIN_PM': round(df_Tabla_Calculos_Corriente['Corriente de neutro mn.'].iloc[1], 2),
                        'LN_CORR_MED_PM': round(df_Tabla_Calculos_Corriente['Corriente de neutro'].iloc[1], 2),
                        'LN_CORR_MAX_PM': round(df_Tabla_Calculos_Corriente['Corriente de neutro mx.'].iloc[1], 2),
                        'L12_MIN_MN': round(df_Tabla_Calculos_Tension['Tensin mn. L12'].iloc[2], 2),
                        'L12_MED_MN': round(df_Tabla_Calculos_Tension['Tensin L12'].iloc[2], 2),
                        'L12_MAX_MN': round(df_Tabla_Calculos_Tension['Tensin mx. L12'].iloc[2], 2),
                        'L23_MIN_MN': round(df_Tabla_Calculos_Tension['Tensin mn. L23'].iloc[2], 2),
                        'L23_MED_MN': round(df_Tabla_Calculos_Tension['Tensin L23'].iloc[2], 2),
                        'L23_MAX_MN': round(df_Tabla_Calculos_Tension['Tensin mx. L23'].iloc[2], 2),
                        'L31_MIN_MN': round(df_Tabla_Calculos_Tension['Tensin mn. L31'].iloc[2], 2),
                        'L31_MED_MN': round(df_Tabla_Calculos_Tension['Tensin L31'].iloc[2], 2),
                        'L31_MAX_MN': round(df_Tabla_Calculos_Tension['Tensin mx. L31'].iloc[2], 2),
                        'L1_CORR_MIN_MN': round(df_Tabla_Calculos_Corriente['Corriente mn. L1'].iloc[2], 2),
                        'L1_CORR_MED_MN': round(df_Tabla_Calculos_Corriente['Corriente L1'].iloc[2], 2),
                        'L1_CORR_MAX_MN': round(df_Tabla_Calculos_Corriente['Corriente mx. L1'].iloc[2], 2),
                        'L2_CORR_MIN_MN': round(df_Tabla_Calculos_Corriente['Corriente mn. L2'].iloc[2], 2),
                        'L2_CORR_MED_MN': round(df_Tabla_Calculos_Corriente['Corriente L2'].iloc[2], 2),
                        'L2_CORR_MAX_MN': round(df_Tabla_Calculos_Corriente['Corriente mx. L2'].iloc[2], 2),
                        'L3_CORR_MIN_MN': round(df_Tabla_Calculos_Corriente['Corriente mn. L3'].iloc[2], 2),
                        'L3_CORR_MED_MN': round(df_Tabla_Calculos_Corriente['Corriente L3'].iloc[2], 2),
                        'L3_CORR_MAX_MN': round(df_Tabla_Calculos_Corriente['Corriente mx. L3'].iloc[2], 2),
                        'LN_CORR_MIN_MN': round(df_Tabla_Calculos_Corriente['Corriente de neutro mn.'].iloc[2], 2),
                        'LN_CORR_MED_MN': round(df_Tabla_Calculos_Corriente['Corriente de neutro'].iloc[2], 2),
                        'LN_CORR_MAX_MN': round(df_Tabla_Calculos_Corriente['Corriente de neutro mx.'].iloc[2], 2),
                        'val_Pct_Max_VL1': round(var_Lista_Variaciones[3], 2),
                        'val_Pct_Max_VL2': round(var_Lista_Variaciones[4], 2),
                        'val_Pct_Max_VL3': round(var_Lista_Variaciones[5], 2),
                        'val_Pct_Min_VL1': round(var_Lista_Variaciones[0], 2),
                        'val_Pct_Min_VL2': round(var_Lista_Variaciones[1], 2),
                        'val_Pct_Min_VL3': round(var_Lista_Variaciones[2], 2),
                        'V1_DESBTEN_MED_PR': round(df_Tabla_Calculos_Desb_Tension['Tensin L12'].iloc[0], 2),
                        'V2_DESBTEN_MED_PR': round(df_Tabla_Calculos_Desb_Tension['Tensin L23'].iloc[0], 2),
                        'V3_DESBTEN_MED_PR': round(df_Tabla_Calculos_Desb_Tension['Tensin L31'].iloc[0], 2),
                        'DESBTEN_PROMEDIO_PR': round(df_Tabla_Calculos_Desb_Tension['Promedio'].iloc[0], 2),
                        'V1_DESBTEN_DELTA_PR': round(df_Tabla_Calculos_Desb_Tension['delta_V1'].iloc[0], 2),
                        'V2_DESBTEN_DELTA_PR': round(df_Tabla_Calculos_Desb_Tension['delta_V2'].iloc[0], 2),
                        'V3_DESBTEN_DELTA_PR': round(df_Tabla_Calculos_Desb_Tension['delta_V3'].iloc[0], 2),
                        'VALUE_DESBTEN_PR': round(df_Tabla_Calculos_Desb_Tension['Desbalance'].iloc[0], 2),
                        'V1_DESBTEN_MED_MX': round(df_Tabla_Calculos_Desb_Tension['Tensin L12'].iloc[3], 2),
                        'V2_DESBTEN_MED_MX': round(df_Tabla_Calculos_Desb_Tension['Tensin L23'].iloc[3], 2),
                        'V3_DESBTEN_MED_MX': round(df_Tabla_Calculos_Desb_Tension['Tensin L31'].iloc[3], 2),
                        'DESBTEN_PROMEDIO_MX': round(df_Tabla_Calculos_Desb_Tension['Promedio'].iloc[3], 2),
                        'V1_DESBTEN_DELTA_MX': round(df_Tabla_Calculos_Desb_Tension['delta_V1'].iloc[3], 2),
                        'V2_DESBTEN_DELTA_MX': round(df_Tabla_Calculos_Desb_Tension['delta_V2'].iloc[3], 2),
                        'V3_DESBTEN_DELTA_MX': round(df_Tabla_Calculos_Desb_Tension['delta_V3'].iloc[3], 2),
                        'VALUE_DESBTEN_MX': round(df_Tabla_Calculos_Desb_Tension['Desbalance'].iloc[3], 2),
                        'V1_DESBTEN_MED_PM': round(df_Tabla_Calculos_Desb_Tension['Tensin L12'].iloc[1], 2),
                        'V2_DESBTEN_MED_PM': round(df_Tabla_Calculos_Desb_Tension['Tensin L23'].iloc[1], 2),
                        'V3_DESBTEN_MED_PM': round(df_Tabla_Calculos_Desb_Tension['Tensin L31'].iloc[1], 2),
                        'DESBTEN_PROMEDIO_PM': round(df_Tabla_Calculos_Desb_Tension['Promedio'].iloc[1], 2),
                        'V1_DESBTEN_DELTA_PM': round(df_Tabla_Calculos_Desb_Tension['delta_V1'].iloc[1], 2),
                        'V2_DESBTEN_DELTA_PM': round(df_Tabla_Calculos_Desb_Tension['delta_V2'].iloc[1], 2),
                        'V3_DESBTEN_DELTA_PM': round(df_Tabla_Calculos_Desb_Tension['delta_V3'].iloc[1], 2),
                        'VALUE_DESBTEN_PM': round(df_Tabla_Calculos_Desb_Tension['Desbalance'].iloc[1], 2),
                        'V1_DESBTEN_MED_MN': round(df_Tabla_Calculos_Desb_Tension['Tensin L12'].iloc[2], 2),
                        'V2_DESBTEN_MED_MN': round(df_Tabla_Calculos_Desb_Tension['Tensin L23'].iloc[2], 2),
                        'V3_DESBTEN_MED_MN': round(df_Tabla_Calculos_Desb_Tension['Tensin L31'].iloc[2], 2),
                        'DESBTEN_PROMEDIO_MN': round(df_Tabla_Calculos_Desb_Tension['Promedio'].iloc[2], 2),
                        'V1_DESBTEN_DELTA_MN': round(df_Tabla_Calculos_Desb_Tension['delta_V1'].iloc[2], 2),
                        'V2_DESBTEN_DELTA_MN': round(df_Tabla_Calculos_Desb_Tension['delta_V2'].iloc[2], 2),
                        'V3_DESBTEN_DELTA_MN': round(df_Tabla_Calculos_Desb_Tension['delta_V3'].iloc[2], 2),
                        'VALUE_DESBTEN_MN': round(df_Tabla_Calculos_Desb_Tension['Desbalance'].iloc[2], 2),
                        'V1_DESBCORR_MED_PR': round(df_Tabla_Calculos_Desb_Corriente['Corriente L1'].iloc[0], 2),
                        'V2_DESBCORR_MED_PR': round(df_Tabla_Calculos_Desb_Corriente['Corriente L2'].iloc[0], 2),
                        'V3_DESBCORR_MED_PR': round(df_Tabla_Calculos_Desb_Corriente['Corriente L3'].iloc[0], 2),
                        'DESBCORR_PROMEDIO_PR': round(df_Tabla_Calculos_Desb_Corriente['Promedio'].iloc[0], 2),
                        'DESBCORR_MAXMED_PR': round(df_Tabla_Calculos_Desb_Corriente['max_Corrientes_Medias'].iloc[0], 2),
                        'VALUE_DESBCORR_PR': round(df_Tabla_Calculos_Desb_Corriente['Desbalance'].iloc[0], 2),
                        'V1_DESBCORR_MED_MX': round(df_Tabla_Calculos_Desb_Corriente['Corriente L1'].iloc[3], 2),
                        'V2_DESBCORR_MED_MX': round(df_Tabla_Calculos_Desb_Corriente['Corriente L2'].iloc[3], 2),
                        'V3_DESBCORR_MED_MX': round(df_Tabla_Calculos_Desb_Corriente['Corriente L3'].iloc[3], 2),
                        'DESBCORR_PROMEDIO_MX': round(df_Tabla_Calculos_Desb_Corriente['Promedio'].iloc[3], 2),
                        'DESBCORR_MAXMED_MX': round(df_Tabla_Calculos_Desb_Corriente['max_Corrientes_Medias'].iloc[3], 2),
                        'VALUE_DESBCORR_MX': round(df_Tabla_Calculos_Desb_Corriente['Desbalance'].iloc[3], 2),
                        'V1_DESBCORR_MED_PM': round(df_Tabla_Calculos_Desb_Corriente['Corriente L1'].iloc[1], 2),
                        'V2_DESBCORR_MED_PM': round(df_Tabla_Calculos_Desb_Corriente['Corriente L2'].iloc[1], 2),
                        'V3_DESBCORR_MED_PM': round(df_Tabla_Calculos_Desb_Corriente['Corriente L3'].iloc[1], 2),
                        'DESBCORR_PROMEDIO_PM': round(df_Tabla_Calculos_Desb_Corriente['Promedio'].iloc[1], 2),
                        'DESBCORR_MAXMED_PM': round(df_Tabla_Calculos_Desb_Corriente['max_Corrientes_Medias'].iloc[1], 2),
                        'VALUE_DESBCORR_PM': round(df_Tabla_Calculos_Desb_Corriente['Desbalance'].iloc[1], 2),
                        'V1_DESBCORR_MED_MN': round(df_Tabla_Calculos_Desb_Corriente['Corriente L1'].iloc[2], 2),
                        'V2_DESBCORR_MED_MN': round(df_Tabla_Calculos_Desb_Corriente['Corriente L2'].iloc[2], 2),
                        'V3_DESBCORR_MED_MN': round(df_Tabla_Calculos_Desb_Corriente['Corriente L3'].iloc[2], 2),
                        'DESBCORR_PROMEDIO_MN': round(df_Tabla_Calculos_Desb_Corriente['Promedio'].iloc[2], 2),
                        'DESBCORR_MAXMED_MN': round(df_Tabla_Calculos_Desb_Corriente['max_Corrientes_Medias'].iloc[2], 2),
                        'VALUE_DESBCORR_MN': round(df_Tabla_Calculos_Desb_Corriente['Desbalance'].iloc[2], 2),
                        'PQS_POT_ACT_MIN_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Activa mn. III'].iloc[0], 2),
                        'PQS_POT_ACT_MED_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Activa III'].iloc[0], 2),
                        'PQS_POT_ACT_MAX_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Activa mx. III'].iloc[0], 2),
                        'PQS_POT_CAP_MIN_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva mn. III'].iloc[0], 2),
                        'PQS_POT_CAP_MED_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva III'].iloc[0], 2),
                        'PQS_POT_CAP_MAX_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva mx. III'].iloc[0], 2),
                        'PQS_POT_IND_MIN_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva mn. III'].iloc[0], 2),
                        'PQS_POT_IND_MED_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva III'].iloc[0], 2),
                        'PQS_POT_IND_MAX_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva mx. III'].iloc[0], 2),
                        'PQS_POT_APA_MIN_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente mn. III'].iloc[0], 2),
                        'PQS_POT_APA_MED_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente III'].iloc[0], 2),
                        'PQS_POT_APA_MAX_PR': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente mx. III'].iloc[0], 2),
                        'PQS_POT_ACT_MIN_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Activa mn. III'].iloc[3], 2),
                        'PQS_POT_ACT_MED_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Activa III'].iloc[3], 2),
                        'PQS_POT_ACT_MAX_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Activa mx. III'].iloc[3], 2),
                        'PQS_POT_CAP_MIN_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva mn. III'].iloc[3], 2),
                        'PQS_POT_CAP_MED_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva III'].iloc[3], 2),
                        'PQS_POT_CAP_MAX_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva mx. III'].iloc[3], 2),
                        'PQS_POT_IND_MIN_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva mn. III'].iloc[3], 2),
                        'PQS_POT_IND_MED_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva III'].iloc[3], 2),
                        'PQS_POT_IND_MAX_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva mx. III'].iloc[3], 2),
                        'PQS_POT_APA_MIN_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente mn. III'].iloc[3], 2),
                        'PQS_POT_APA_MED_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente III'].iloc[3], 2),
                        'PQS_POT_APA_MAX_MX': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente mx. III'].iloc[3], 2),
                        'PQS_POT_ACT_MIN_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Activa mn. III'].iloc[1], 2),
                        'PQS_POT_ACT_MED_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Activa III'].iloc[1], 2),
                        'PQS_POT_ACT_MAX_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Activa mx. III'].iloc[1], 2),
                        'PQS_POT_CAP_MIN_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva mn. III'].iloc[1], 2),
                        'PQS_POT_CAP_MED_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva III'].iloc[1], 2),
                        'PQS_POT_CAP_MAX_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva mx. III'].iloc[1], 2),
                        'PQS_POT_IND_MIN_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva mn. III'].iloc[1], 2),
                        'PQS_POT_IND_MED_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva III'].iloc[1], 2),
                        'PQS_POT_IND_MAX_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva mx. III'].iloc[1], 2),
                        'PQS_POT_APA_MIN_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente mn. III'].iloc[1], 2),
                        'PQS_POT_APA_MED_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente III'].iloc[1], 2),
                        'PQS_POT_APA_MAX_PM': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente mx. III'].iloc[1], 2),
                        'PQS_POT_ACT_MIN_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Activa mn. III'].iloc[2], 2),
                        'PQS_POT_ACT_MED_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Activa III'].iloc[2], 2),
                        'PQS_POT_ACT_MAX_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Activa mx. III'].iloc[2], 2),
                        'PQS_POT_CAP_MIN_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva mn. III'].iloc[2], 2),
                        'PQS_POT_CAP_MED_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva III'].iloc[2], 2),
                        'PQS_POT_CAP_MAX_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Capacitiva mx. III'].iloc[2], 2),
                        'PQS_POT_IND_MIN_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva mn. III'].iloc[2], 2),
                        'PQS_POT_IND_MED_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva III'].iloc[2], 2),
                        'PQS_POT_IND_MAX_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Inductiva mx. III'].iloc[2], 2),
                        'PQS_POT_APA_MIN_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente mn. III'].iloc[2], 2),
                        'PQS_POT_APA_MED_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente III'].iloc[2], 2),
                        'PQS_POT_APA_MAX_MN': round(df_Tabla_Calculos_PQS_Potencias['P.Aparente mx. III'].iloc[2], 2),
                        'PQS_CARGABILIDAD_MAX': round(var_Lista_PQS_Carg_Disp[0], 2),
                        'DISPONIBILIDAD_CARGA': round(var_Lista_PQS_Carg_Disp[1], 2),
                        'FP_POT_CAP_MIN_PR': round(df_Tabla_Calculos_FactorPotencia['F.P. Mn. III -'].iloc[0], 2),
                        'FP_POT_CAP_MED_PR': round(df_Tabla_Calculos_FactorPotencia['F.P. III -'].iloc[0], 2),
                        'FP_POT_CAP_MAX_PR': round(df_Tabla_Calculos_FactorPotencia['F.P. Mx. III -'].iloc[0], 2),
                        'FP_POT_IND_MIN_PR': round(df_Tabla_Calculos_FactorPotencia['F.P. Mn. III'].iloc[0], 2),
                        'FP_POT_IND_MED_PR': round(df_Tabla_Calculos_FactorPotencia['F.P. III'].iloc[0], 2),
                        'FP_POT_IND_MAX_PR': round(df_Tabla_Calculos_FactorPotencia['F.P. Mx. III'].iloc[0], 2),
                        'FP_POT_CAP_MIN_MX': round(df_Tabla_Calculos_FactorPotencia['F.P. Mn. III -'].iloc[3], 2),
                        'FP_POT_CAP_MED_MX': round(df_Tabla_Calculos_FactorPotencia['F.P. III -'].iloc[3], 2),
                        'FP_POT_CAP_MAX_MX': round(df_Tabla_Calculos_FactorPotencia['F.P. Mx. III -'].iloc[3], 2),
                        'FP_POT_IND_MIN_MX': round(df_Tabla_Calculos_FactorPotencia['F.P. Mn. III'].iloc[3], 2),
                        'FP_POT_IND_MED_MX': round(df_Tabla_Calculos_FactorPotencia['F.P. III'].iloc[3], 2),
                        'FP_POT_IND_MAX_MX': round(df_Tabla_Calculos_FactorPotencia['F.P. Mx. III'].iloc[3], 2),
                        'FP_POT_CAP_MIN_PM': round(df_Tabla_Calculos_FactorPotencia['F.P. Mn. III -'].iloc[1], 2),
                        'FP_POT_CAP_MED_PM': round(df_Tabla_Calculos_FactorPotencia['F.P. III -'].iloc[1], 2),
                        'FP_POT_CAP_MAX_PM': round(df_Tabla_Calculos_FactorPotencia['F.P. Mx. III -'].iloc[1], 2),
                        'FP_POT_IND_MIN_PM': round(df_Tabla_Calculos_FactorPotencia['F.P. Mn. III'].iloc[1], 2),
                        'FP_POT_IND_MED_PM': round(df_Tabla_Calculos_FactorPotencia['F.P. III'].iloc[1], 2),
                        'FP_POT_IND_MAX_PM': round(df_Tabla_Calculos_FactorPotencia['F.P. Mx. III'].iloc[1], 2),
                        'FP_POT_CAP_MIN_MN': round(df_Tabla_Calculos_FactorPotencia['F.P. Mn. III -'].iloc[2], 2),
                        'FP_POT_CAP_MED_MN': round(df_Tabla_Calculos_FactorPotencia['F.P. III -'].iloc[2], 2),
                        'FP_POT_CAP_MAX_MN': round(df_Tabla_Calculos_FactorPotencia['F.P. Mx. III -'].iloc[2], 2),
                        'FP_POT_IND_MIN_MN': round(df_Tabla_Calculos_FactorPotencia['F.P. Mn. III'].iloc[2], 2),
                        'FP_POT_IND_MED_MN': round(df_Tabla_Calculos_FactorPotencia['F.P. III'].iloc[2], 2),
                        'FP_POT_IND_MAX_MN': round(df_Tabla_Calculos_FactorPotencia['F.P. Mx. III'].iloc[2], 2),
                        'FACT_PO_CAP_MIN_PR': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mn. III - Cap']['F.P. Mn. III']['Percentil'], 2),
                        'FACT_PO_CAP_MED_PR': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. III - Cap']['F.P. III']['Percentil'], 2),
                        'FACT_PO_CAP_MAX_PR': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mx. III - Cap']['F.P. Mx. III']['Percentil'], 2),
                        'FACT_PO_IND_MIN_PR': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mn. III - Ind']['F.P. Mn. III']['Percentil'], 2),
                        'FACT_PO_IND_MED_PR': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. III - Ind']['F.P. III']['Percentil'], 2),
                        'FACT_PO_IND_MAX_PR': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mx. III - Ind']['F.P. Mx. III']['Percentil'], 2),
                        'FACT_PO_CAP_MIN_MX': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mn. III - Cap']['F.P. Mn. III']['Maximo'], 2),
                        'FACT_PO_CAP_MED_MX': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. III - Cap']['F.P. III']['Maximo'], 2),
                        'FACT_PO_CAP_MAX_MX': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mx. III - Cap']['F.P. Mx. III']['Maximo'], 2),
                        'FACT_PO_IND_MIN_MX': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mn. III - Ind']['F.P. Mn. III']['Maximo'], 2),
                        'FACT_PO_IND_MED_MX': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. III - Ind']['F.P. III']['Maximo'], 2),
                        'FACT_PO_IND_MAX_MX': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mx. III - Ind']['F.P. Mx. III']['Maximo'], 2),
                        'FACT_PO_CAP_MIN_PM': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mn. III - Cap']['F.P. Mn. III']['Promedio'], 2),
                        'FACT_PO_CAP_MED_PM': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. III - Cap']['F.P. III']['Promedio'], 2),
                        'FACT_PO_CAP_MAX_PM': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mx. III - Cap']['F.P. Mx. III']['Promedio'], 2),
                        'FACT_PO_IND_MIN_PM': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mn. III - Ind']['F.P. Mn. III']['Promedio'], 2),
                        'FACT_PO_IND_MED_PM': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. III - Ind']['F.P. III']['Promedio'], 2),
                        'FACT_PO_IND_MAX_PM': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mx. III - Ind']['F.P. Mx. III']['Promedio'], 2),
                        'FACT_PO_CAP_MIN_MN': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mn. III - Cap']['F.P. Mn. III']['Minimo'], 2),
                        'FACT_PO_CAP_MED_MN': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. III - Cap']['F.P. III']['Minimo'], 2),
                        'FACT_PO_CAP_MAX_MN': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mx. III - Cap']['F.P. Mx. III']['Minimo'], 2),
                        'FACT_PO_IND_MIN_MN': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mn. III - Ind']['F.P. Mn. III']['Minimo'], 2),
                        'FACT_PO_IND_MED_MN': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. III - Ind']['F.P. III']['Minimo'], 2),
                        'FACT_PO_IND_MAX_MN': round(df_Tabla_Calculos_FactorPotenciaGeneral['F.P. Mx. III - Ind']['F.P. Mx. III']['Minimo'], 2),
                        'EN_ACTIVA_MED_PR': round(df_Tabla_Calculos_Energias['E.Activa T1'].iloc[0], 2),
                        'EN_CAPACITIVA_MED_PR': round(df_Tabla_Calculos_Energias['E.Capacitiva T1'].iloc[0], 2),
                        'EN_INDUCTIVA_MED_PR': round(df_Tabla_Calculos_Energias['E.Inductiva T1'].iloc[0], 2),
                        'EN_KWH_PR': round(df_Tabla_Calculos_Energias['KWH'].iloc[0], 2),
                        'EN_KARH_IND_PR': round(df_Tabla_Calculos_Energias['KARH_IND'].iloc[0], 2),
                        'EN_KVARH_CAP_PR': round(df_Tabla_Calculos_Energias['KVARH_CAP'].iloc[0], 2),
                        'EN_FACT_POTENCIA_NEG_PR': round(df_Tabla_Calculos_Energias['F.P. III -'].iloc[0], 2),
                        'EN_FACT_POTENCIA_POS_PR': round(df_Tabla_Calculos_Energias['F.P. III'].iloc[0], 2),
                        'EN_ACTIVA_MED_MX': round(df_Tabla_Calculos_Energias['E.Activa T1'].iloc[3], 2),
                        'EN_CAPACITIVA_MED_MX': round(df_Tabla_Calculos_Energias['E.Capacitiva T1'].iloc[3], 2),
                        'EN_INDUCTIVA_MED_MX': round(df_Tabla_Calculos_Energias['E.Inductiva T1'].iloc[3], 2),
                        'EN_KWH_MX': round(df_Tabla_Calculos_Energias['KWH'].iloc[3], 2),
                        'EN_KARH_IND_MX': round(df_Tabla_Calculos_Energias['KARH_IND'].iloc[3], 2),
                        'EN_KVARH_CAP_MX': round(df_Tabla_Calculos_Energias['KVARH_CAP'].iloc[3], 2),
                        'EN_FACT_POTENCIA_NEG_MX': round(df_Tabla_Calculos_Energias['F.P. III -'].iloc[3], 2),
                        'EN_FACT_POTENCIA_POS_MX': round(df_Tabla_Calculos_Energias['F.P. III'].iloc[3], 2),
                        'EN_ACTIVA_MED_PM': round(df_Tabla_Calculos_Energias['E.Activa T1'].iloc[1], 2),
                        'EN_CAPACITIVA_MED_PM': round(df_Tabla_Calculos_Energias['E.Capacitiva T1'].iloc[1], 2),
                        'EN_INDUCTIVA_MED_PM': round(df_Tabla_Calculos_Energias['E.Inductiva T1'].iloc[1], 2),
                        'EN_KWH_PM': round(df_Tabla_Calculos_Energias['KWH'].iloc[1], 2),
                        'EN_KARH_IND_PM': round(df_Tabla_Calculos_Energias['KARH_IND'].iloc[1], 2),
                        'EN_KVARH_CAP_PM': round(df_Tabla_Calculos_Energias['KVARH_CAP'].iloc[1], 2),
                        'EN_FACT_POTENCIA_NEG_PM': round(df_Tabla_Calculos_Energias['F.P. III -'].iloc[1], 2),
                        'EN_FACT_POTENCIA_POS_PM': round(df_Tabla_Calculos_Energias['F.P. III'].iloc[1], 2),
                        'EN_ACTIVA_MED_MN': round(df_Tabla_Calculos_Energias['E.Activa T1'].iloc[2], 2),
                        'EN_CAPACITIVA_MED_MN': round(df_Tabla_Calculos_Energias['E.Capacitiva T1'].iloc[2], 2),
                        'EN_INDUCTIVA_MED_MN': round(df_Tabla_Calculos_Energias['E.Inductiva T1'].iloc[2], 2),
                        'EN_KWH_MN': round(df_Tabla_Calculos_Energias['KWH'].iloc[2], 2),
                        'EN_KARH_IND_MN': round(df_Tabla_Calculos_Energias['KARH_IND'].iloc[2], 2),
                        'EN_KVARH_CAP_MN': round(df_Tabla_Calculos_Energias['KVARH_CAP'].iloc[2], 2),
                        'EN_FACT_POTENCIA_NEG_MN': round(df_Tabla_Calculos_Energias['F.P. III -'].iloc[2], 2),
                        'EN_FACT_POTENCIA_POS_MN': round(df_Tabla_Calculos_Energias['F.P. III'].iloc[2], 2),
                        'THD_DIST_TENSION_L1_MAX_PR': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L1'].iloc[0], 2),
                        'THD_DIST_TENSION_L2_MAX_PR': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L2'].iloc[0], 2),
                        'THD_DIST_TENSION_L3_MAX_PR': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L3'].iloc[0], 2),
                        'THD_DIST_TENSION_DIST_ARM_PR': round(df_Tabla_Calculos_DistTension['var_Ref_Distorsion_Tension'].iloc[0], 2),
                        'THD_DIST_TENSION_L1_MAX_MX': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L1'].iloc[3], 2),
                        'THD_DIST_TENSION_L2_MAX_MX': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L2'].iloc[3], 2),
                        'THD_DIST_TENSION_L3_MAX_MX': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L3'].iloc[3], 2),
                        'THD_DIST_TENSION_DIST_ARM_MX': round(df_Tabla_Calculos_DistTension['var_Ref_Distorsion_Tension'].iloc[3], 2),
                        'THD_DIST_TENSION_L1_MAX_PM': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L1'].iloc[1], 2),
                        'THD_DIST_TENSION_L2_MAX_PM': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L2'].iloc[1], 2),
                        'THD_DIST_TENSION_L3_MAX_PM': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L3'].iloc[1], 2),
                        'THD_DIST_TENSION_DIST_ARM_PM': round(df_Tabla_Calculos_DistTension['var_Ref_Distorsion_Tension'].iloc[1], 2),
                        'THD_DIST_TENSION_L1_MAX_MN': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L1'].iloc[2], 2),
                        'THD_DIST_TENSION_L2_MAX_MN': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L2'].iloc[2], 2),
                        'THD_DIST_TENSION_L3_MAX_MN': round(df_Tabla_Calculos_DistTension['V THD/d Mx. L3'].iloc[2], 2),
                        'THD_DIST_TENSION_DIST_ARM_MN': round(df_Tabla_Calculos_DistTension['var_Ref_Distorsion_Tension'].iloc[2], 2),
                        'THDV_ARM_N3_L1_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L1'].iloc[0], 2),
                        'THDV_ARM_N5_L1_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L1'].iloc[0], 2),
                        'THDV_ARM_N7_L1_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L1'].iloc[0], 2),
                        'THDV_ARM_N9_L1_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L1'].iloc[0], 2),
                        'THDV_ARM_N11_L1_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L1'].iloc[0], 2),
                        'THDV_ARM_N13_L1_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L1'].iloc[0], 2),
                        'THDV_ARM_N15_L1_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L1'].iloc[0], 2),
                        'THDV_ARM_N3_L1_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L1'].iloc[3], 2),
                        'THDV_ARM_N5_L1_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L1'].iloc[3], 2),
                        'THDV_ARM_N7_L1_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L1'].iloc[3], 2),
                        'THDV_ARM_N9_L1_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L1'].iloc[3], 2),
                        'THDV_ARM_N11_L1_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L1'].iloc[3], 2),
                        'THDV_ARM_N13_L1_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L1'].iloc[3], 2),
                        'THDV_ARM_N15_L1_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L1'].iloc[3], 2),
                        'THDV_ARM_N3_L1_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L1'].iloc[1], 2),
                        'THDV_ARM_N5_L1_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L1'].iloc[1], 2),
                        'THDV_ARM_N7_L1_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L1'].iloc[1], 2),
                        'THDV_ARM_N9_L1_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L1'].iloc[1], 2),
                        'THDV_ARM_N11_L1_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L1'].iloc[1], 2),
                        'THDV_ARM_N13_L1_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L1'].iloc[1], 2),
                        'THDV_ARM_N15_L1_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L1'].iloc[1], 2),
                        'THDV_ARM_N3_L1_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L1'].iloc[2], 2),
                        'THDV_ARM_N5_L1_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L1'].iloc[2], 2),
                        'THDV_ARM_N7_L1_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L1'].iloc[2], 2),
                        'THDV_ARM_N9_L1_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L1'].iloc[2], 2),
                        'THDV_ARM_N11_L1_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L1'].iloc[2], 2),
                        'THDV_ARM_N13_L1_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L1'].iloc[2], 2),
                        'THDV_ARM_N15_L1_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L1'].iloc[2], 2),
                        'THDV_ARM_N3_L2_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L2'].iloc[0], 2),
                        'THDV_ARM_N5_L2_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L2'].iloc[0], 2),
                        'THDV_ARM_N7_L2_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L2'].iloc[0], 2),
                        'THDV_ARM_N9_L2_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L2'].iloc[0], 2),
                        'THDV_ARM_N11_L2_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L2'].iloc[0], 2),
                        'THDV_ARM_N13_L2_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L2'].iloc[0], 2),
                        'THDV_ARM_N15_L2_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L2'].iloc[0], 2),
                        'THDV_ARM_N3_L2_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L2'].iloc[3], 2),
                        'THDV_ARM_N5_L2_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L2'].iloc[3], 2),
                        'THDV_ARM_N7_L2_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L2'].iloc[3], 2),
                        'THDV_ARM_N9_L2_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L2'].iloc[3], 2),
                        'THDV_ARM_N11_L2_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L2'].iloc[3], 2),
                        'THDV_ARM_N13_L2_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L2'].iloc[3], 2),
                        'THDV_ARM_N15_L2_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L2'].iloc[3], 2),
                        'THDV_ARM_N3_L2_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L2'].iloc[1], 2),
                        'THDV_ARM_N5_L2_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L2'].iloc[1], 2),
                        'THDV_ARM_N7_L2_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L2'].iloc[1], 2),
                        'THDV_ARM_N9_L2_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L2'].iloc[1], 2),
                        'THDV_ARM_N11_L2_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L2'].iloc[1], 2),
                        'THDV_ARM_N13_L2_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L2'].iloc[1], 2),
                        'THDV_ARM_N15_L2_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L2'].iloc[1], 2),
                        'THDV_ARM_N3_L2_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L2'].iloc[2], 2),
                        'THDV_ARM_N5_L2_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L2'].iloc[2], 2),
                        'THDV_ARM_N7_L2_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L2'].iloc[2], 2),
                        'THDV_ARM_N9_L2_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L2'].iloc[2], 2),
                        'THDV_ARM_N11_L2_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L2'].iloc[2], 2),
                        'THDV_ARM_N13_L2_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L2'].iloc[2], 2),
                        'THDV_ARM_N15_L2_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L2'].iloc[2], 2),
                        'THDV_ARM_N3_L3_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L3'].iloc[0], 2),
                        'THDV_ARM_N5_L3_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L3'].iloc[0], 2),
                        'THDV_ARM_N7_L3_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L3'].iloc[0], 2),
                        'THDV_ARM_N9_L3_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L3'].iloc[0], 2),
                        'THDV_ARM_N11_L3_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L3'].iloc[0], 2),
                        'THDV_ARM_N13_L3_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L3'].iloc[0], 2),
                        'THDV_ARM_N15_L3_PR': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L3'].iloc[0], 2),
                        'THDV_ARM_N3_L3_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L3'].iloc[3], 2),
                        'THDV_ARM_N5_L3_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L3'].iloc[3], 2),
                        'THDV_ARM_N7_L3_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L3'].iloc[3], 2),
                        'THDV_ARM_N9_L3_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L3'].iloc[3], 2),
                        'THDV_ARM_N11_L3_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L3'].iloc[3], 2),
                        'THDV_ARM_N13_L3_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L3'].iloc[3], 2),
                        'THDV_ARM_N15_L3_MX': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L3'].iloc[3], 2),
                        'THDV_ARM_N3_L3_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L3'].iloc[1], 2),
                        'THDV_ARM_N5_L3_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L3'].iloc[1], 2),
                        'THDV_ARM_N7_L3_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L3'].iloc[1], 2),
                        'THDV_ARM_N9_L3_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L3'].iloc[1], 2),
                        'THDV_ARM_N11_L3_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L3'].iloc[1], 2),
                        'THDV_ARM_N13_L3_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L3'].iloc[1], 2),
                        'THDV_ARM_N15_L3_PM': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L3'].iloc[1], 2),
                        'THDV_ARM_N3_L3_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 3 L3'].iloc[2], 2),
                        'THDV_ARM_N5_L3_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 5 L3'].iloc[2], 2),
                        'THDV_ARM_N7_L3_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 7 L3'].iloc[2], 2),
                        'THDV_ARM_N9_L3_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 9 L3'].iloc[2], 2),
                        'THDV_ARM_N11_L3_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 11 L3'].iloc[2], 2),
                        'THDV_ARM_N13_L3_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 13 L3'].iloc[2], 2),
                        'THDV_ARM_N15_L3_MN': round(df_Tabla_Calculos_Armonicos_DistTension['Arm. tensin 15 L3'].iloc[2], 2),
                        'THD_DIST_CORRIENTE_L1_MED_PR': round(df_Tabla_Calculos_DistCorriente['A THD/d L1'].iloc[0], 2),
                        'THD_DIST_CORRIENTE_L2_MED_PR': round(df_Tabla_Calculos_DistCorriente['A THD/d L2'].iloc[0], 2),
                        'THD_DIST_CORRIENTE_L3_MED_PR': round(df_Tabla_Calculos_DistCorriente['A THD/d L3'].iloc[0], 2),
                        'THD_DIST_CORRIENTE_L1_MED_MX': round(df_Tabla_Calculos_DistCorriente['A THD/d L1'].iloc[3], 2),
                        'THD_DIST_CORRIENTE_L2_MED_MX': round(df_Tabla_Calculos_DistCorriente['A THD/d L2'].iloc[3], 2),
                        'THD_DIST_CORRIENTE_L3_MED_MX': round(df_Tabla_Calculos_DistCorriente['A THD/d L3'].iloc[3], 2),
                        'THD_DIST_CORRIENTE_L1_MED_PM': round(df_Tabla_Calculos_DistCorriente['A THD/d L1'].iloc[1], 2),
                        'THD_DIST_CORRIENTE_L2_MED_PM': round(df_Tabla_Calculos_DistCorriente['A THD/d L2'].iloc[1], 2),
                        'THD_DIST_CORRIENTE_L3_MED_PM': round(df_Tabla_Calculos_DistCorriente['A THD/d L3'].iloc[1], 2),
                        'THD_DIST_CORRIENTE_L1_MED_MN': round(df_Tabla_Calculos_DistCorriente['A THD/d L1'].iloc[2], 2),
                        'THD_DIST_CORRIENTE_L2_MED_MN': round(df_Tabla_Calculos_DistCorriente['A THD/d L2'].iloc[2], 2),
                        'THD_DIST_CORRIENTE_L3_MED_MN': round(df_Tabla_Calculos_DistCorriente['A THD/d L3'].iloc[2], 2),
                        'DHIT_ARM_N3_L1_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L1'].iloc[0], 2),
                        'DHIT_ARM_N5_L1_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L1'].iloc[0], 2),
                        'DHIT_ARM_N7_L1_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L1'].iloc[0], 2),
                        'DHIT_ARM_N9_L1_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L1'].iloc[0], 2),
                        'DHIT_ARM_N11_L1_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L1'].iloc[0], 2),
                        'DHIT_ARM_N13_L1_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L1'].iloc[0], 2),
                        'DHIT_ARM_N15_L1_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L1'].iloc[0], 2),
                        'DHIT_ARM_N3_L1_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L1'].iloc[3], 2),
                        'DHIT_ARM_N5_L1_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L1'].iloc[3], 2),
                        'DHIT_ARM_N7_L1_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L1'].iloc[3], 2),
                        'DHIT_ARM_N9_L1_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L1'].iloc[3], 2),
                        'DHIT_ARM_N11_L1_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L1'].iloc[3], 2),
                        'DHIT_ARM_N13_L1_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L1'].iloc[3], 2),
                        'DHIT_ARM_N15_L1_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L1'].iloc[3], 2),
                        'DHIT_ARM_N3_L1_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L1'].iloc[1], 2),
                        'DHIT_ARM_N5_L1_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L1'].iloc[1], 2),
                        'DHIT_ARM_N7_L1_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L1'].iloc[1], 2),
                        'DHIT_ARM_N9_L1_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L1'].iloc[1], 2),
                        'DHIT_ARM_N11_L1_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L1'].iloc[1], 2),
                        'DHIT_ARM_N13_L1_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L1'].iloc[1], 2),
                        'DHIT_ARM_N15_L1_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L1'].iloc[1], 2),
                        'DHIT_ARM_N3_L1_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L1'].iloc[2], 2),
                        'DHIT_ARM_N5_L1_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L1'].iloc[2], 2),
                        'DHIT_ARM_N7_L1_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L1'].iloc[2], 2),
                        'DHIT_ARM_N9_L1_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L1'].iloc[2], 2),
                        'DHIT_ARM_N11_L1_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L1'].iloc[2], 2),
                        'DHIT_ARM_N13_L1_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L1'].iloc[2], 2),
                        'DHIT_ARM_N15_L1_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L1'].iloc[2], 2),
                        'DHIT_ARM_N3_L2_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L2'].iloc[0], 2),
                        'DHIT_ARM_N5_L2_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L2'].iloc[0], 2),
                        'DHIT_ARM_N7_L2_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L2'].iloc[0], 2),
                        'DHIT_ARM_N9_L2_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L2'].iloc[0], 2),
                        'DHIT_ARM_N11_L2_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L2'].iloc[0], 2),
                        'DHIT_ARM_N13_L2_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L2'].iloc[0], 2),
                        'DHIT_ARM_N15_L2_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L2'].iloc[0], 2),
                        'DHIT_ARM_N3_L2_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L2'].iloc[3], 2),
                        'DHIT_ARM_N5_L2_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L2'].iloc[3], 2),
                        'DHIT_ARM_N7_L2_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L2'].iloc[3], 2),
                        'DHIT_ARM_N9_L2_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L2'].iloc[3], 2),
                        'DHIT_ARM_N11_L2_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L2'].iloc[3], 2),
                        'DHIT_ARM_N13_L2_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L2'].iloc[3], 2),
                        'DHIT_ARM_N15_L2_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L2'].iloc[3], 2),
                        'DHIT_ARM_N3_L2_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L2'].iloc[1], 2),
                        'DHIT_ARM_N5_L2_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L2'].iloc[1], 2),
                        'DHIT_ARM_N7_L2_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L2'].iloc[1], 2),
                        'DHIT_ARM_N9_L2_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L2'].iloc[1], 2),
                        'DHIT_ARM_N11_L2_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L2'].iloc[1], 2),
                        'DHIT_ARM_N13_L2_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L2'].iloc[1], 2),
                        'DHIT_ARM_N15_L2_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L2'].iloc[1], 2),
                        'DHIT_ARM_N3_L2_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L2'].iloc[2], 2),
                        'DHIT_ARM_N5_L2_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L2'].iloc[2], 2),
                        'DHIT_ARM_N7_L2_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L2'].iloc[2], 2),
                        'DHIT_ARM_N9_L2_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L2'].iloc[2], 2),
                        'DHIT_ARM_N11_L2_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L2'].iloc[2], 2),
                        'DHIT_ARM_N13_L2_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L2'].iloc[2], 2),
                        'DHIT_ARM_N15_L2_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L2'].iloc[2], 2),
                        'DHIT_ARM_N3_L3_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L3'].iloc[0], 2),
                        'DHIT_ARM_N5_L3_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L3'].iloc[0], 2),
                        'DHIT_ARM_N7_L3_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L3'].iloc[0], 2),
                        'DHIT_ARM_N9_L3_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L3'].iloc[0], 2),
                        'DHIT_ARM_N11_L3_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L3'].iloc[0], 2),
                        'DHIT_ARM_N13_L3_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L3'].iloc[0], 2),
                        'DHIT_ARM_N15_L3_PR': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L3'].iloc[0], 2),
                        'DHIT_ARM_N3_L3_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L3'].iloc[3], 2),
                        'DHIT_ARM_N5_L3_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L3'].iloc[3], 2),
                        'DHIT_ARM_N7_L3_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L3'].iloc[3], 2),
                        'DHIT_ARM_N9_L3_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L3'].iloc[3], 2),
                        'DHIT_ARM_N11_L3_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L3'].iloc[3], 2),
                        'DHIT_ARM_N13_L3_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L3'].iloc[3], 2),
                        'DHIT_ARM_N15_L3_MX': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L3'].iloc[3], 2),
                        'DHIT_ARM_N3_L3_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L3'].iloc[1], 2),
                        'DHIT_ARM_N5_L3_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L3'].iloc[1], 2),
                        'DHIT_ARM_N7_L3_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L3'].iloc[1], 2),
                        'DHIT_ARM_N9_L3_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L3'].iloc[1], 2),
                        'DHIT_ARM_N11_L3_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L3'].iloc[1], 2),
                        'DHIT_ARM_N13_L3_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L3'].iloc[1], 2),
                        'DHIT_ARM_N15_L3_PM': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L3'].iloc[1], 2),
                        'DHIT_ARM_N3_L3_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 3 L3'].iloc[2], 2),
                        'DHIT_ARM_N5_L3_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 5 L3'].iloc[2], 2),
                        'DHIT_ARM_N7_L3_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 7 L3'].iloc[2], 2),
                        'DHIT_ARM_N9_L3_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 9 L3'].iloc[2], 2),
                        'DHIT_ARM_N11_L3_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 11 L3'].iloc[2], 2),
                        'DHIT_ARM_N13_L3_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 13 L3'].iloc[2], 2),
                        'DHIT_ARM_N15_L3_MN': round(df_Tabla_Calculos_Armonicos_DistCorriente['Arm. corriente 15 L3'].iloc[2], 2),
                        'TDD_LINEA_1_PR': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L1'].iloc[0], 2),
                        'TDD_LINEA_2_PR': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L2'].iloc[0], 2),
                        'TDD_LINEA_3_PR': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L3'].iloc[0], 2),
                        'TDD_LINEA_1_MX': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L1'].iloc[3], 2),
                        'TDD_LINEA_2_MX': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L2'].iloc[3], 2),
                        'TDD_LINEA_3_MX': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L3'].iloc[3], 2),
                        'TDD_LINEA_1_PM': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L1'].iloc[1], 2),
                        'TDD_LINEA_2_PM': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L2'].iloc[1], 2),
                        'TDD_LINEA_3_PM': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L3'].iloc[1], 2),
                        'TDD_LINEA_1_MN': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L1'].iloc[2], 2),
                        'TDD_LINEA_2_MN': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L2'].iloc[2], 2),
                        'TDD_LINEA_3_MN': round(df_Tabla_Calculos_CargabilidadTDD['resultado_TDD_L3'].iloc[2], 2),
                        #'PLT_FLICKER_L1_MED_PR': round(df_Tabla_Calculos_Flicker['Plt L1'].iloc[0], 2),
                        #'PLT_FLICKER_L2_MED_PR': round(df_Tabla_Calculos_Flicker['Plt L2'].iloc[0], 2),
                        #'PLT_FLICKER_L3_MED_PR': round(df_Tabla_Calculos_Flicker['Plt L3'].iloc[0], 2),
                        #'PLT_FLICKER_L1_MED_MX': round(df_Tabla_Calculos_Flicker['Plt L1'].iloc[3], 2),
                        #'PLT_FLICKER_L2_MED_MX': round(df_Tabla_Calculos_Flicker['Plt L2'].iloc[3], 2),
                        #'PLT_FLICKER_L3_MED_MX': round(df_Tabla_Calculos_Flicker['Plt L3'].iloc[3], 2),
                        #'PLT_FLICKER_L1_MED_PM': round(df_Tabla_Calculos_Flicker['Plt L1'].iloc[1], 2),
                        #'PLT_FLICKER_L2_MED_PM': round(df_Tabla_Calculos_Flicker['Plt L2'].iloc[1], 2),
                        #'PLT_FLICKER_L3_MED_PM': round(df_Tabla_Calculos_Flicker['Plt L3'].iloc[1], 2),
                        #'PLT_FLICKER_L1_MED_MN': round(df_Tabla_Calculos_Flicker['Plt L1'].iloc[2], 2),
                        #'PLT_FLICKER_L2_MED_MN': round(df_Tabla_Calculos_Flicker['Plt L2'].iloc[2], 2),
                        #'PLT_FLICKER_L3_MED_MN': round(df_Tabla_Calculos_Flicker['Plt L3'].iloc[2], 2),
                        'FACTOR_K_L1_MIN_PR': round(df_Tabla_Calculos_FactorK['Factor K mn. L1'].iloc[0], 2),
                        'FACTOR_K_L2_MIN_PR': round(df_Tabla_Calculos_FactorK['Factor K mn. L2'].iloc[0], 2),
                        'FACTOR_K_L3_MIN_PR': round(df_Tabla_Calculos_FactorK['Factor K mn. L3'].iloc[0], 2),
                        'FACTOR_K_L1_MED_PR': round(df_Tabla_Calculos_FactorK['Factor K L1'].iloc[0], 2),
                        'FACTOR_K_L2_MED_PR': round(df_Tabla_Calculos_FactorK['Factor K L2'].iloc[0], 2),
                        'FACTOR_K_L3_MED_PR': round(df_Tabla_Calculos_FactorK['Factor K L3'].iloc[0], 2),
                        'FACTOR_K_L1_MAX_PR': round(df_Tabla_Calculos_FactorK['Factor K mx. L1'].iloc[0], 2),
                        'FACTOR_K_L2_MAX_PR': round(df_Tabla_Calculos_FactorK['Factor K mx. L2'].iloc[0], 2),
                        'FACTOR_K_L3_MAX_PR': round(df_Tabla_Calculos_FactorK['Factor K mx. L3'].iloc[0], 2),
                        'FACTOR_K_L1_MIN_MX': round(df_Tabla_Calculos_FactorK['Factor K mn. L1'].iloc[3], 2),
                        'FACTOR_K_L2_MIN_MX': round(df_Tabla_Calculos_FactorK['Factor K mn. L2'].iloc[3], 2),
                        'FACTOR_K_L3_MIN_MX': round(df_Tabla_Calculos_FactorK['Factor K mn. L3'].iloc[3], 2),
                        'FACTOR_K_L1_MED_MX': round(df_Tabla_Calculos_FactorK['Factor K L1'].iloc[3], 2),
                        'FACTOR_K_L2_MED_MX': round(df_Tabla_Calculos_FactorK['Factor K L2'].iloc[3], 2),
                        'FACTOR_K_L3_MED_MX': round(df_Tabla_Calculos_FactorK['Factor K L3'].iloc[3], 2),
                        'FACTOR_K_L1_MAX_MX': round(df_Tabla_Calculos_FactorK['Factor K mx. L1'].iloc[3], 2),
                        'FACTOR_K_L2_MAX_MX': round(df_Tabla_Calculos_FactorK['Factor K mx. L2'].iloc[3], 2),
                        'FACTOR_K_L3_MAX_MX': round(df_Tabla_Calculos_FactorK['Factor K mx. L3'].iloc[3], 2),
                        'FACTOR_K_L1_MIN_PM': round(df_Tabla_Calculos_FactorK['Factor K mn. L1'].iloc[1], 2),
                        'FACTOR_K_L2_MIN_PM': round(df_Tabla_Calculos_FactorK['Factor K mn. L2'].iloc[1], 2),
                        'FACTOR_K_L3_MIN_PM': round(df_Tabla_Calculos_FactorK['Factor K mn. L3'].iloc[1], 2),
                        'FACTOR_K_L1_MED_PM': round(df_Tabla_Calculos_FactorK['Factor K L1'].iloc[1], 2),
                        'FACTOR_K_L2_MED_PM': round(df_Tabla_Calculos_FactorK['Factor K L2'].iloc[1], 2),
                        'FACTOR_K_L3_MED_PM': round(df_Tabla_Calculos_FactorK['Factor K L3'].iloc[1], 2),
                        'FACTOR_K_L1_MAX_PM': round(df_Tabla_Calculos_FactorK['Factor K mx. L1'].iloc[1], 2),
                        'FACTOR_K_L2_MAX_PM': round(df_Tabla_Calculos_FactorK['Factor K mx. L2'].iloc[1], 2),
                        'FACTOR_K_L3_MAX_PM': round(df_Tabla_Calculos_FactorK['Factor K mx. L3'].iloc[1], 2),
                        'FACTOR_K_L1_MIN_MN': round(df_Tabla_Calculos_FactorK['Factor K mn. L1'].iloc[2], 2),
                        'FACTOR_K_L2_MIN_MN': round(df_Tabla_Calculos_FactorK['Factor K mn. L2'].iloc[2], 2),
                        'FACTOR_K_L3_MIN_MN': round(df_Tabla_Calculos_FactorK['Factor K mn. L3'].iloc[2], 2),
                        'FACTOR_K_L1_MED_MN': round(df_Tabla_Calculos_FactorK['Factor K L1'].iloc[2], 2),
                        'FACTOR_K_L2_MED_MN': round(df_Tabla_Calculos_FactorK['Factor K L2'].iloc[2], 2),
                        'FACTOR_K_L3_MED_MN': round(df_Tabla_Calculos_FactorK['Factor K L3'].iloc[2], 2),
                        'FACTOR_K_L1_MAX_MN': round(df_Tabla_Calculos_FactorK['Factor K mx. L1'].iloc[2], 2),
                        'FACTOR_K_L2_MAX_MN': round(df_Tabla_Calculos_FactorK['Factor K mx. L2'].iloc[2], 2),
                        'FACTOR_K_L3_MAX_MN': round(df_Tabla_Calculos_FactorK['Factor K mx. L3'].iloc[2], 2),
                        'var_valor_Maximo_Corrientes_Max': round(valor_Maximo_Corrientes, 2),
                        'var_valor_Corriente_Cortacircuito': round(valor_Corriente_Cortacircuito, 2),
                        'var_valor_ISC_sobre_IL': round(valor_ISC_sobre_IL, 2),
                        'var_valor_Limite_TDD': round(valor_Limite_TDD, 2),
                        'LIMITE_ARMONICO_0_10': round(valores_Limites_Armonicos['ARM_0_10'], 2),
                        'LIMITE_ARMONICO_11_16': round(valores_Limites_Armonicos['ARM_11_16'], 2),
                        'LIMITE_ARMONICO_17_22': round(valores_Limites_Armonicos['ARM_17_22'], 2),
                        'LIMITE_ARMONICO_23_34': round(valores_Limites_Armonicos['ARM_23_34'], 2),
                        'LIMITE_ARMONICO_35': round(valores_Limites_Armonicos['ARM_35'], 2),
                        'OBSERVACION_TENSION_NUM_1': observaciones_Tension['cumple_Condicion'],
                        'OBSERVACION_TENSION_NUM_2': observaciones_Tension['cumple_Condicion_2'],
                        'OBSERVACION_TENSION_NUM_3': observaciones_Tension['cumple_Condicion_3'],
                        'OBSERVACION_CORRIENTE_NUM_1': list(observaciones_Corriente['val_Maximo_Corriente'].keys())[0],
                        'OBSERVACION_CORRIENTE_NUM_2': list(observaciones_Corriente['val_Maximo_Corriente'].values())[0],
                        'OBSERVACION_CORRIENTE_NUM_3': observaciones_Corriente['resultado_Comparacion_Corriente'],
                        'OBSERVACION_CORRIENTE_NUM_4': list(observaciones_Corriente['val_Maximo_CorrienteNeutra'].values())[0],
                        'OBSERVACION_DESBTENSION_NUM_1': f"{observaciones_DesbTension[0]}, {observaciones_DesbTension[1]}",
                        'OBSERVACION_DESBTENSION_NUM_2': f"{observaciones_DesbTension[2]}",
                        'OBSERVACION_DESBCORRIENTE_NUM_1': f"{observaciones_DesbCorriente[0]}, {observaciones_DesbCorriente[1]}",
                        'OBSERVACION_DESBCORRIENTE_NUM_2': f"{observaciones_DesbCorriente[2]}",
                        'OBSERVACION_THDV_NUM_1': f"{observaciones_THDV}",
                        'OBSERVACION_THDV_NUM_2': f"{valor_Referencia_THDV}",
                        'OBSERVACION_ARMCORRIENTE_NUM_1': f"{observaciones_ArmonicosCorriente['resultado_Arm_3_9']}",
                        'OBSERVACION_ARMCORRIENTE_NUM_2': f"{observaciones_ArmonicosCorriente['resultado_Arm_11']}",
                        'OBSERVACION_TDD_NUM_1': f"{observaciones_TDD[0]}",
                        'OBSERVACION_TDD_NUM_2': f"{observaciones_TDD[1]}"
                    }
                    
                    
                    graficar_Timeline_Tension_Plotly(dataFrame=var_Tabla_Tensiones, variables=list_Columns_Grafico_Tension, fecha_col='fecha_y_Hora', limites=[var_Tabla_Tensiones['var_Limite_Inferior_Tension'].iloc[0], var_Tabla_Tensiones['var_Limite_Superior_Tension'].iloc[0]], titulo='REGISTROS DE TENSIÓN')

                    graficar_Timeline_Corriente_Plotly(dataFrame=var_Tabla_Corrientes, variables=list_Columns_Grafico_Corriente, fecha_col='fecha_y_Hora', limite=var_Tabla_Corrientes['var_Limite_Corriente_Nominal'].iloc[0], titulo='REGISTROS DE CORRIENTE')

                    graficar_Timeline_DesbTension_Plotly(dataFrame=df_Tabla_Desb_Tension, variables=list_Columns_Grafico_DesbTension, fecha_col='fecha_y_Hora', limite=df_Tabla_Desb_Tension['var_Ref_Desbalance_Tension'].iloc[0], titulo='REGISTROS DESBALANCE DE TENSIÓN')
                        
                    graficar_Timeline_DesbCorriente_Plotly(dataFrame=df_Tabla_Desb_Corriente, variables=list_Columns_Grafico_DesbCorriente, fecha_col='fecha_y_Hora', limite=df_Tabla_Desb_Corriente['var_Ref_Desbalance_Corriente'].iloc[0], titulo='REGISTROS DESBALANCE DE CORRIENTE')

                    graficar_Timeline_PQS_ActApa_Plotly(dataFrame=df_Tabla_PQS_Final, variables=list_Columns_Grafico_DesbCorriente_ActApa, fecha_col='fecha_y_Hora', titulo='REGISTROS DE POTENCIA - Activa / Aparente (kW / kVA)')

                    graficar_Timeline_PQS_CapInd_Plotly(dataFrame=df_Tabla_PQS_Final, variables=list_Columns_Grafico_DesbCorriente_CapInd, fecha_col='fecha_y_Hora', titulo='REGISTROS DE POTENCIA - Capacitiva / Inductiva (kVAR)')

                    graficar_Timeline_FactPotencia_Plotly(dataFrame=df_Tabla_FactPotenciaFinal, variables=list_Columns_Grafico_FactorPot, medidas_dataFrame=data_Cantidad_NEG_POS_FactorPotencia, fecha_col='fecha_y_Hora', titulo='REGISTROS DE POTENCIA - Factor de Potencia')

                    graficar_Timeline_Distorsion_Tension_Plotly(dataFrame=df_Tabla_Distorsion_TensionFinal, variables=list_Columns_Distorsion_Tension, fecha_col='fecha_y_Hora', limite=df_Tabla_Distorsion_TensionFinal['var_Ref_Distorsion_Tension'].iloc[0], titulo='REGISTROS DISTORSIÓN ARMÓNICA DE TENSIÓN - THDV')

                    graficar_Timeline_Distorsion_Corriente_Plotly(dataFrame=df_Tabla_Distorsion_CorrienteFinal, variables=list_Columns_Distorsion_Corriente, fecha_col='fecha_y_Hora', limite=None, titulo='REGISTROS DISTORSIÓN ARMÓNICA DE CORRIENTE - THDI')
                        
                    graficar_Timeline_CargabilidadTDD_Plotly(dataFrame=df_Tabla_TDDFinal, variables=list_Columns_Armonicos_Cargabilidad_TDD, fecha_col='fecha_y_Hora', limite=valor_Limite_TDD, titulo='REGISTROS DISTORSIÓN TOTAL DE DEMANDA')

                    graficar_Timeline_FactorK_Plotly(dataFrame=df_Tabla_FactorKFinal, variables=list_Columns_FactorK, fecha_col='fecha_y_Hora', limite=None, titulo='REGISTROS DE FACTOR K')
                        
                    #generar_Graficos_Barras_Energias_Plotly(dataFrame=df_Tabla_Energias, variables=list_Columns_Graficos_Consolidado_Energia, fecha_col='Fecha/Hora')
                
                    generar_Graficos_Barras_Energias_Plotly(dataFrame=df_Tabla_Energias.copy(), variables=list_Columns_Graficos_Consolidado_Energia, fecha_col='Fecha/hora')


                    st.success("Gráficos generados correctamente.")
                
                except Exception as e:
                    
                    print(e)
            
        except Exception as e:
            st.error(f"Error al cargar los archivo .txt o procesar los datos: {e}")
            
    else:
        st.write("Por favor, sube los archivos .txt para comenzar.")
