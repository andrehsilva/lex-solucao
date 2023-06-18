from typing import Text
import streamlit as st
import pandas as pd
import io
import datetime
from datetime import datetime as dt
import plotly.express as px
from datetime import date
import re
import unicodedata
import time
import base64

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
    #df_lex = pd.read_csv('/home/andrerodrigues/Documentos/8_dashboard_python/2_streamlit/conexia/lex_escolas_produtos.csv', sep=';')
    df_lex = pd.read_csv('lex_escolas_produtos.csv', sep=';')
    
    df_lex = df_lex.loc[df_lex['Profile']=='Aluno']
    df_lex = df_lex.drop(columns=['Tenant'])
    df_lex = df_lex.rename(columns={'School':'Escola','Profile':'Perfil','Qtde':'Licenças', 'Product':'Produto'})
    df_lex['Licenças'] = df_lex['Licenças'].astype(int)
    df_lex = df_lex.sort_values(by=['Licenças'], ascending=False)
    #configurações do streamlit

    st.set_page_config(page_title="Script de licenças",page_icon="🧊",layout="wide",initial_sidebar_state="expanded")

    # funcoes
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
    st.sidebar.title('Script de licenças - Simulador')


    page = ['B2B','B2C']
    choice = st.sidebar.selectbox('Selecione:',page)

    if choice == 'B2B':

        today = date.today().strftime('%d-%m-%Y')
        cliente_tipo = 'B2B'
        
        st.info("Simulador B2B")

        marca = st.radio("Selecione a marca:",('AZ', 'High Five', 'My Lyfe'), horizontal=True)

        if marca == 'AZ':
            st.write('Selecionado: AZ')
            sheetname = 'itens_b2b_az' 
        elif marca == 'High Five':
            st.write('Selecionado: High Five')
            sheetname = 'itens_b2b_hf'
        elif marca == 'My Lyfe':
            st.write('Selecionado: My Life')
            sheetname = 'itens_b2b_my'
        
        # CNPJ da escola
        cliente = st.text_input('Digite o CNPJ da escola:')


        # Carrega o arquivo
        file = st.file_uploader("Selecione um arquivo Excel", type=["xlsx"])

        if file is not None:
           # Lê o arquivo Excel
            simul = pd.read_excel(file)

           # Exibe o DataFrame
            #st.dataframe(simul)
            simulador = simul.copy()

            df_cliente = simulador.loc[simulador['CNPJ ESCOLA'] == cliente]
            df_cliente = df_cliente.fillna(0)
            df_cliente = df_cliente.mask(df_cliente == 'x', 1)
            df_cliente = df_cliente.mask(df_cliente == 'X', 1)
            df_cliente['SEGMENTO'] = df_cliente['SEGMENTO'].str.replace('ANOS INICIAIS','FUNDAMENTAL ANOS INICIAIS')
            df_cliente['SEGMENTO'] = df_cliente['SEGMENTO'].str.replace('ANOS FINAIS','FUNDAMENTAL ANOS FINAIS')
            df_cliente['SEGMENTO'] = df_cliente['SEGMENTO'].str.replace('ENSINO MÉDIO ','ENSINO MÉDIO')
        
            name_escola = df_cliente['ESCOLA'].unique()
            name_escola_completo = name_escola[0]
            name_escola_completo = name_escola_completo.encode('ascii', errors='ignore').decode('utf-8')
            name_escola = name_escola_completo.split()[0]
            #name_escola
            listname = ['EIRELI','ESCOLA','COLÉGIO','COLEGIO','CRECHE','INSTITUTO','EDUCANDARIO','COMUNIDADE','SOCIEDADE','CENTRO','EDUCACIONAL','EDUCACAO','EDUCAÇÃO','ASSOCIACAO','ASSOCIAÇÃO','INFANTIL','ENSINO','FUNDAMENTAL','MEDIO','MÉDIO','LTDA','-','/',' ','.']
            for n in listname:
                name_escola_completo = name_escola_completo.replace(n,'')
            #name_escola_completo

            df_cliente = df_cliente[['CUSTOMER GROUP - ESCOLA','SQUAD','ESCOLA','CNPJ ESCOLA','SÉRIE','SEGMENTO','TOTAL ALUNOS 2023',
                                    'Materiais Impressos','Plataforma AZ','H5 Bilingual Education - Language Book + CLIL  e PBL',
                                    'International Journey + App H5','Aulas Ao Vivo - ZOOM','Módulo de Comunicação',
                                    'Liga das Corujinhas Games Educativos','Educacross Games Matemática','Educacross Games Lingua Portuguesa',
                                    'Educacross High Five','Cantalelê','My Life','UBBU','High Five Plus','4 Avaliações Nacionais','1 Simulado ENEM',
                                    '5 Simulados Enem','1 Simulado Regional','Itinerários Formativos Micro cursos (2 IF)','Mundo Leitor','ELT Aluno',
                                    'Alfabetização ','Learning','TOTAL PREÇO TABELA 2023','DESCONTO POR VOLUME',
                                    'CUPOM EXTRA DE DESCONTO','PREÇO ACORDADO ESCOLA 2023 (ANUAL)','% DESCONTO POR SÉRIE 2023',
                                    'TOTAL SEM DESCONTO','TOTAL COM DESCONTO SEM ELT','PREÇO  ANUAL FINAL SEM ELT','PREÇO FINAL COM  ELT (SE APLICÁVEL)']]

            df_cliente=df_cliente.assign(Extra="")

            st.dataframe(df_cliente)

            @st.cache_data
            
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv().encode('utf-8')

            csv = convert_df(df_cliente)

            st.download_button(
                label="Download arquivo",
                data=csv,
                file_name=f'{cliente}.csv',
                mime='text/csv',
            )

    


    if choice == "B2C":
        
        escolas = df['Escola'].unique()
        options = st.multiselect('Selecionados',escolas)


        if not options:
            st.info('Selecione uma ou mais escolas')
            
        else:
            for escola in escolas:
                if escola in options:
                    #df_escola = df.loc[(df['Escola'] == escola) & (df['Segmento'] == segmentos)]
                    df_escola = df.loc[(df['Escola'] == escola)]
                    df_escola = df_escola.sort_values(by=['Licenças'], ascending=False)
                    df_escola = df_escola.reset_index(drop=True)

                    #df_escola = df_escola.groupby(['Nome da Licença','Segmento']).sum().reset_index()

                    
                    st.subheader(escola)
                    

                    fig = px.bar(df_escola, x='Nome da Licença', y='Licenças', color='Segmento')
                    st.plotly_chart(fig)

                    col1, col2 = st.columns(2)

                    qtd_licencas_escolas = df_escola['Licenças'].sum().astype(int)
                    col1.info((f'licenças ativas: {qtd_licencas_escolas :,} '.format(qtd_licencas_escolas)).upper())

                    prod = len(df_escola['Nome da Licença'].unique())
                    col2.success((f'Produtos adquiridos: {prod}').upper())

                    #st.bar_chart(df_escola)
                    
                    

                    buffer2 = io.BytesIO()
                    st.write(df_escola)
                    with pd.ExcelWriter(buffer2, engine='xlsxwriter') as writer:
                        df_escola.to_excel(writer, sheet_name="Licenças Ativas", index=False)
                        writer.save()

                        st.download_button(
                            label="📥 Download",
                            data=buffer2,
                            file_name=escola+".xlsx",
                            mime="application/vnd.ms-excel"
                        )

    

    
                    


        #st.write(df_geral)
                    
    