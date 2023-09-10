from typing import Text
import streamlit as st
import io
from datetime import date
import time
import base64
import pandas as pd
import numpy as np
from datetime import date, datetime
import re
import unicodedata
import xlsxwriter
import openpyxl
from conexia_b2b import b2b


#### codigo

def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True


if check_password():

    buffer = io.BytesIO()

    #df = pd.read_excel('/home/andrerodrigues/Documentos/8_dashboard_python/2_streamlit/conexia/atual.xlsx')
    df = pd.read_excel('atual.xlsx')
    
    
    df = df.drop(columns=['Tenant','School','OrderDate','ComboCode','CNPJ', 'Ean do produto', 'StartDate','EndDate'])
    df = df.rename(columns={'SchoolName':'Escola','OrderNumber':'Nº Pedido','LicenseName':'Nome da Licença','Grade':'Segmento',
                           'Student':'Licenças'})
    df['Nº Pedido'] = df['Nº Pedido'].astype(str)
    qtd_licencas = df['Licenças'].sum().astype(int)

    qtd_escola = len(df['Escola'].unique())


    #####################################################################################

 
    #configurações do streamlit

    st.set_page_config(page_title="Script de soluções",page_icon="⭐",layout="wide",initial_sidebar_state="expanded")

    ##################################################################################
    ## funcoes
    def maiuscula(data):
        data.columns = data.columns.str.upper()
        for columns in data.columns:
            data[columns] = data[columns].str.upper()
        return data

    def minuscula(data):
        data.columns = data.columns.str.lower()
        for columns in data.columns:
            data[columns] = data[columns].str.lower()
        return data
    
    def download_link(df, file_name, file_label):
        csv = df.to_excel(index=False)
        b64 = base64.b64encode(csv.encode()).decode()  # Converte para base64
        href = f'<a href="data:file/csv;base64,{b64}" download="{file_name}">{file_label}</a>'
        return href
        
    
    # sidebar
    st.sidebar.image('https://sso.lex.education/assets/images/new-lex-logo.png', width=100)
    st.sidebar.title('Script de solução - Simulador')


    page = ['CONEXIA B2B','CONEXIA B2C','SEB','PREMIUM/UNIQUE']
    choice = st.sidebar.selectbox('Selecione:',page)


    with open('consolidado.xlsm', "rb") as template_file:
        template_byte = template_file.read()

        st.sidebar.download_button(label="Download arquivo template",
                            data=template_byte,
                            file_name="template_simulador.xlsm",
                            mime='application/octet-stream')
        

    

    ##########B2B################


    if choice == 'CONEXIA B2B':
        b2b()


    if choice == "B2C":
        st.info('Módulo B2C em construção :)')



    if choice == "B2B2C":
        st.info('Módulo B2C em construção :)')
                    


        #st.write(df_geral)
                    
    
