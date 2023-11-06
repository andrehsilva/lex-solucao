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
import os
from converter import excel_csv, csv_excel


##########################################################################################################################################################
##########################################################################################################################################################

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


##########################################################################################################################################################
##########################################################################################################################################################

if check_password():
    buffer = io.BytesIO()

    #configurações do streamlit

    st.set_page_config(page_title="Script de soluções",page_icon="⭐",layout="wide",initial_sidebar_state="expanded")

    
    ## sidebar
    st.sidebar.image('https://sso.lex.education/assets/images/new-lex-logo.png', width=100)
    st.sidebar.title('Script de solução - Simulador')

    #  'SEB','PREMIUM/UNIQUE',
    
    page = ['CONEXIA B2B','CONEXIA B2C','SEB','PREMIUM-UNIQUE','EXCEL PARA CSV','CSV PARA EXCEL','PEDIDO PROGRAMADO']
    choice = st.sidebar.selectbox('Selecione:',page)


##########################################################################################################################################################
##########################################################################################################################################################


    if choice == 'CONEXIA B2B':
        
        marca = 'AZ' ## ou AZ SESC B2B ou AZ/SESC
        sheetname = 'itens'
        planilha = 'itens.xlsx'
        today = date.today().strftime('%d-%m-%Y')
        cliente_tipo = 'B2B'

        st.info("Simulador - CONEXIA B2B")
        agree = st.checkbox('Marque para usar o cálculo do script - (Não recomendado!)')
        #  29.271.264/0001-61
        cliente = st.text_input('Digite o CNPJ da escola:')
        # Carrega o arquivo
        file = st.file_uploader("Selecione um arquivo Excel", type=["xlsm"])
        
        
        if file is not None:
            simul0 = pd.read_excel(file, sheet_name='cálculos Anual')
            simul0=simul0.assign(Bimestre="ANUAL")
            simul0.replace(0, np.nan, inplace = True)
            
            simul1 = pd.read_excel(file, sheet_name='cálculos 1º Bim')
            simul1=simul1.assign(Bimestre="1º BIMESTRE")
            simul1.replace(0, np.nan, inplace = True)

            simul2 = pd.read_excel(file, sheet_name='cálculos 2º Bim')
            simul2=simul2.assign(Bimestre="2º BIMESTRE")
            simul2.replace(0, np.nan, inplace = True)

            simul3 = pd.read_excel(file, sheet_name='cálculos 3º Bim')
            simul3=simul3.assign(Bimestre="3º BIMESTRE")
            simul3.replace(0, np.nan, inplace = True)

            simul4 = pd.read_excel(file, sheet_name='cálculos 4º Bim')
            simul4=simul4.assign(Bimestre="4º BIMESTRE")
            simul4.replace(0, np.nan, inplace = True)

            #alterar regra conforme leitura das planilhas
            simul = pd.concat([simul0,simul1,simul2,simul3,simul4])
            
            #simul = pd.concat([simul1,simul2,simul3,simul4])
            simul = simul[simul['Quantidade de alunos']>0]
            #simul
            if agree:
                desconto = pd.read_excel(file, sheet_name='Formulário Anual 2024')
                desconto = desconto.iloc[:, :6] 
                desconto = desconto[['FORMULÁRIO DE AQUISIÇÃO DE MATERIAL DIDÁTICO','Unnamed: 5']]
                desconto = desconto.rename(columns={'FORMULÁRIO DE AQUISIÇÃO DE MATERIAL DIDÁTICO':'Série','Unnamed: 5':'% Desconto Extra%'})
            
                indice = [25,26,27,28,29,47,48,49,50,51,67,68,69,70,84,85,98,112,113]
                desconto = desconto.iloc[indice]

                del(simul['% Desconto Extra'])
                del(simul['% Desconto Total'])
                simul = simul.drop_duplicates()
                
                simul = pd.merge(simul, desconto, on=['Série'], how='inner')
                simul['% Desconto Volume'] = simul['% Desconto Volume'].apply(lambda x: x[:-1])
                simul['% Desconto Volume'] = simul['% Desconto Volume'].astype('float64')/100
                simul['% Desconto Total'] = simul['% Desconto Extra%'] + simul['% Desconto Volume']
                simul = simul.rename(columns={'% Desconto Extra%':'% Desconto Extra'})
            
            simul = simul.rename(columns={'Construindo a Alfabetização':'Alfabetização','Itinerários Formativos Micro cursos     (2 IF)':'Itinerários','H5 - (3 Horas) Language Book + CLIL e PBL ':'H5 - 3 Horas','H5 - (2 horas)\nInternational Journey + \nApp H5':'H5 - 2 horas Journey','H5 Plus\n (3 horas extras)':'H5 Plus','My Life\n(Base)':'My Life - Base','My Life\n(2024)':'My Life - 2024','Binoculo By Tell Me\n(Base)':'Binoculo - Base','Educacross Ed. Infantil\n(Base)':'Educacross Infantil - Base','Educacross\n(Base)':'Educacross - Base','Educacross AZ\n(Base)':'Educacross AZ - Base','Educacross H5\n(Base)':'Educacross H5 - Base','Ubbu\n(Base)':'Ubbu - Base','Binoculo By Tell Me\n(2024)':'Binoculo - 2024','Educacross Ed. Infantil\n(2024)':'Educacross Infantil - 2024','Educacross\n(2024)':'Educacross - 2024','Educacross AZ\n(2024)':'Educacross AZ - 2024','Educacross H5\n(2024)':'Educacross H5 - 2024','Ubbu\n(2024)':'Ubbu - 2024','Árvore\n(1 Módulo)':'Árvore 1 Módulo','Árvore\n(2 Módulos)':'Árvore 2 Módulos','Árvore\n(3 Módulos)':'Árvore 3 Módulos','total aluno/ano\nsem desconto':'total aluno sem desconto','total aluno/ano\ncom desconto sem complementar':'total aluno com desconto sem complementar','total aluno/ano\ncom desconto + Complementares':'total aluno com desconto com Complementares',})
            
            simul = simul[['Série','Segmento','Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês','% Desconto Volume','Quantidade de alunos','Razão Social','CNPJ','Squad','Tipo','Observação','Grupo de cliente','Bimestre','% Desconto Extra','% Desconto Total']]
            #simul.to_excel('simul.xlsx')
            
            
            simulador = simul.copy()
            df_cliente = simulador.loc[simulador['CNPJ'].str.strip() == cliente]
            df_cliente = df_cliente.fillna(0)

            df_cliente['Plataforma AZ'] = df_cliente['Plataforma AZ'].where(df_cliente['Plataforma AZ'] == 0, 1)
            df_cliente['Materiais Impressos AZ'] = df_cliente['Materiais Impressos AZ'].where(df_cliente['Materiais Impressos AZ'] == 0, 1)
            df_cliente['Alfabetização'] = df_cliente['Alfabetização'].where(df_cliente['Alfabetização'] == 0, 1)
            df_cliente['Cantalelê'] = df_cliente['Cantalelê'].where(df_cliente['Cantalelê'] == 0, 1)
            df_cliente['Mundo Leitor'] = df_cliente['Mundo Leitor'].where(df_cliente['Mundo Leitor'] == 0, 1)
            df_cliente['4 Avaliações Nacionais'] = df_cliente['4 Avaliações Nacionais'].where(df_cliente['4 Avaliações Nacionais'] == 0, 1)
            df_cliente['1 Simulado ENEM'] = df_cliente['1 Simulado ENEM'].where(df_cliente['1 Simulado ENEM'] == 0, 1)
            df_cliente['5 Simulados ENEM'] = df_cliente['5 Simulados ENEM'].where(df_cliente['5 Simulados ENEM'] == 0, 1)
            df_cliente['1 Simulado Regional'] = df_cliente['1 Simulado Regional'].where(df_cliente['1 Simulado Regional'] == 0, 1)
            df_cliente['Itinerários'] = df_cliente['Itinerários'].where(df_cliente['Itinerários'] == 0, 1)
            df_cliente['H5 - 3 Horas'] = df_cliente['H5 - 3 Horas'].where(df_cliente['H5 - 3 Horas'] == 0, 1)
            df_cliente['H5 - 2 horas Journey'] = df_cliente['H5 - 2 horas Journey'].where(df_cliente['H5 - 2 horas Journey'] == 0, 1)
            df_cliente['H5 Plus'] = df_cliente['H5 Plus'].where(df_cliente['H5 Plus'] == 0, 1)
            df_cliente['My Life - Base'] = df_cliente['My Life - Base'].where(df_cliente['My Life - Base'] == 0, 1)
            df_cliente['My Life - 2024'] = df_cliente['My Life - 2024'].where(df_cliente['My Life - 2024'] == 0, 1)
            df_cliente['Binoculo - Base'] = df_cliente['Binoculo - Base'].where(df_cliente['Binoculo - Base'] == 0, 1)
            df_cliente['Educacross Infantil - Base'] = df_cliente['Educacross Infantil - Base'].where(df_cliente['Educacross Infantil - Base'] == 0, 1)
            df_cliente['Educacross - Base'] = df_cliente['Educacross - Base'].where(df_cliente['Educacross - Base'] == 0, 1)
            df_cliente['Educacross AZ - Base'] = df_cliente['Educacross AZ - Base'].where(df_cliente['Educacross AZ - Base'] == 0, 1)
            df_cliente['Educacross H5 - Base'] = df_cliente['Educacross H5 - Base'].where(df_cliente['Educacross H5 - Base'] == 0, 1)
            df_cliente['Ubbu - Base'] = df_cliente['Ubbu - Base'].where(df_cliente['Ubbu - Base'] == 0, 1)
            df_cliente['Binoculo - 2024'] = df_cliente['Binoculo - 2024'].where(df_cliente['Binoculo - 2024'] == 0, 1)
            df_cliente['Educacross Infantil - 2024'] = df_cliente['Educacross Infantil - 2024'].where(df_cliente['Educacross Infantil - 2024'] == 0, 1)
            df_cliente['Educacross - 2024'] = df_cliente['Educacross - 2024'].where(df_cliente['Educacross - 2024'] == 0, 1)
            df_cliente['Educacross AZ - 2024'] = df_cliente['Educacross AZ - 2024'].where(df_cliente['Educacross AZ - 2024'] == 0, 1)
            df_cliente['Educacross H5 - 2024'] = df_cliente['Educacross H5 - 2024'].where(df_cliente['Educacross H5 - 2024'] == 0, 1)
            df_cliente['Ubbu - 2024'] = df_cliente['Ubbu - 2024'].where(df_cliente['Ubbu - 2024'] == 0, 1)
            df_cliente['Árvore 1 Módulo'] = df_cliente['Árvore 1 Módulo'].where(df_cliente['Árvore 1 Módulo'] == 0, 1)
            df_cliente['Árvore 2 Módulos'] = df_cliente['Árvore 2 Módulos'].where(df_cliente['Árvore 2 Módulos'] == 0, 1)
            df_cliente['Árvore 3 Módulos'] = df_cliente['Árvore 3 Módulos'].where(df_cliente['Árvore 3 Módulos'] == 0, 1)
            df_cliente['School Guardian'] = df_cliente['School Guardian'].where(df_cliente['School Guardian'] == 0, 1)
            df_cliente['Tindin'] = df_cliente['Tindin'].where(df_cliente['Tindin'] == 0, 1)
            df_cliente['Scholastic Earlybird and Bookflix'] = df_cliente['Scholastic Earlybird and Bookflix'].where(df_cliente['Scholastic Earlybird and Bookflix'] == 0, 1)
            df_cliente['Scholastic Literacy Pro'] = df_cliente['Scholastic Literacy Pro'].where(df_cliente['Scholastic Literacy Pro'] == 0, 1)
            df_cliente['Livro de Inglês'] = df_cliente['Livro de Inglês'].where(df_cliente['Livro de Inglês'] == 0,1)

            
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ed. Infantil','INFANTIL')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Iniciais','FUNDAMENTAL ANOS INICIAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Finais','FUNDAMENTAL ANOS FINAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ensino Médio','ENSINO MÉDIO')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('PV','PRÉ VESTIBULAR')
            df_cliente=df_cliente.assign(Extra="")
            
            ###regra do AZ e Plataforma
            df_cliente.loc[(df_cliente['Plataforma AZ'] == 1) & (df_cliente['Materiais Impressos AZ'] == 1), ['Plataforma AZ']] = 0
            
            ####regra do h5
            df_cliente.loc[(df_cliente['H5 Plus'] == 1) & (df_cliente['H5 - 2 horas Journey'] == 1), ['H5 - 2 horas Journey','H5 - 3 Horas']] = 0
            df_cliente.loc[(df_cliente['H5 - 2 horas Journey'] == 1) & (df_cliente['H5 - 3 Horas'] == 1), ['H5 - 3 Horas']] = 0
            

             ####
            df_client = df_cliente.copy()
            lista = ['Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês']
            
            #df_client.to_excel('cliente.xlsx')
            #df_client

            for item in lista:
                df_client.loc[df_client[item] == 1.0, item] = item
            COLUNAS = ['Série', 'Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra']
            p = pd.DataFrame(columns=COLUNAS)
            
            for i in lista:
                data = df_client[df_client[i] == i].groupby(['Série', 'Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Squad','Tipo','Bimestre',i])['Extra'].count().reset_index()
                data = data.rename(columns={i: 'Produto'})
                p = pd.concat([p,data])
            p = p.sort_values(by=['Série'])
            p = p.reset_index()
            p = p.drop(columns=['index'])
            p = p.drop_duplicates()
            
            
            itens = pd.read_excel(planilha, sheet_name=sheetname)
            itens = itens[['MARCA',2024,'2024+','Produto','DESCRIÇÃO MAGENTO (B2C e B2B)','BIMESTRE','SEGMENTO','SÉRIE','PÚBLICO','TIPO DE FATURAMENTO']]
            itens = itens.rename(columns={'MARCA':'Marca','DESCRIÇÃO MAGENTO (B2C e B2B)':'Descrição Magento','BIMESTRE':'Bimestre','SEGMENTO':'Segmento','SÉRIE':'Série','PÚBLICO':'Público','TIPO DE FATURAMENTO':'Faturamento'})
            itens = itens[(itens['Marca'] == marca) | (itens['Marca'] == 'CONEXIA') | (itens['Marca'] == 'MUNDO LEITOR') | (itens['Marca'] == 'MY LIFE')| (itens['Marca'] == 'HIGH FIVE')]
            
            pdt = pd.merge(p, itens, on=['Série','Bimestre','Segmento','Produto'], how='inner')
            
            cod_serial = pd.read_excel(planilha, sheet_name='cod_serial')
            
            pdt = pd.merge(pdt, cod_serial, on=['Série','Bimestre','Segmento','Público'], how='inner')

            pdt['Ano'] = '2024'
            pdt['SKU'] = pdt['Ano'] + pdt['Serial']
            pdt = pdt[['Série','Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra','Produto','Marca',2024,'2024+','Descrição Magento','Público','Faturamento','Serial','Categoria','Ano','SKU']]
            

            h = re.compile(r'[../\-]')
            pdt['CNPJ_off'] = [h.sub('', x) for x in pdt['CNPJ']]
            pdt['CNPJ_off'] = [x.lstrip('0') for x in pdt['CNPJ_off']]
            pdt['CNPJ_off'] = pdt['CNPJ_off'].astype(float)
            
            cod_nome = pd.read_excel(planilha, sheet_name='nome')
            cod_nome['CNPJ_off'] = cod_nome['CNPJ_off'].astype(float)
            pdt = pd.merge(pdt, cod_nome, on=['CNPJ_off'], how='inner')
            ####################################################################################################
            ######NOVAS REGRAS POR SÉRIE#####################################################
            
            serie = pdt['Série'].unique()
            pdt_final = []
            for i in serie:
                pdt_serie = pdt.loc[pdt['Série'] == i]

                ######Regras

                #pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'CONEXIA') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]
                #pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'MY LIFE') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]
                

                if (pdt_serie['Marca'].str.contains('AZ').any()):
                    pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'CONEXIA') & (pdt_serie['Bimestre'].str.contains('ANUAL')))]
                    pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'AZ') & (pdt_serie['Bimestre'].str.contains('ANUAL')))]
                    #pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'MY LIFE') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]
                    pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'MY LIFE') & (pdt_serie['Bimestre'].str.contains('ANUAL')))]
                    pdt_serie['Marca'] = pdt_serie['Marca'].str.replace('MY LIFE','AZ')
                    pdt_serie['Marca'] = pdt_serie['Marca'].str.replace('CONEXIA','AZ')
                    pdt_serie['Marca'] = pdt_serie['Marca'].str.replace('MUNDO LEITOR','AZ')

                else:
                    pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'CONEXIA') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]
                    #pdt_serie['Marca'] = pdt_serie['Marca'].str.replace('CONEXIA','HIGH FIVE')
                    #pdt_serie
                 
                if (pdt_serie['Marca'].str.contains('MY LIFE').any()):
                        pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'MY LIFE') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]

                if (pdt_serie['Marca'].str.contains('HIGH FIVE').any()):
                    pdt_serie.loc[(pdt_serie['Bimestre'] == 'ANUAL') & (pdt_serie['Marca'] == 'CONEXIA'), ['Marca']] = 'HIGH FIVE'
                    if (pdt_serie['Marca'].str.contains('MY LIFE').any()):
                        pdt_serie.loc[(pdt_serie['Bimestre'] == 'ANUAL') & (pdt_serie['Marca'] == 'CONEXIA'), ['Marca']] = 'HIGH FIVE'
                        pdt_serie.loc[(pdt_serie['Bimestre'] == 'ANUAL') & (pdt_serie['Marca'] == 'MY LIFE'), ['Marca']] = 'HIGH FIVE'

                if (pdt_serie['Marca'].str.contains('MY LIFE').any()):
                        pdt_serie.loc[(pdt_serie['Bimestre'] == 'ANUAL') & (pdt_serie['Marca'] == 'CONEXIA'), ['Marca']] = 'MY LIFE'
                        pdt_serie.loc[(pdt_serie['Bimestre'] == 'ANUAL') & (pdt_serie['Marca'] == 'MY LIFE'), ['Marca']] = 'MY LIFE'
                
                pdt_final.append(pdt_serie)
                pdt_full = pd.concat(pdt_final)

            #pdt_full = pdt_full[~((pdt_full['Marca'] == 'AZ') & (pdt_full['Bimestre'].str.contains('ANUAL')))]
            pdt = pdt_full.copy()
            

            ######End Regra   
        
            pdt['Nome'] = 'SOLUÇÃO ' + pdt['Marca']  + ' - ' + pdt['Escola'] + ' - ' + pdt['Segmento'] + ' - ' + pdt['Série'] + ' - ' + pdt['Bimestre']
            pdt['SKU'] = pdt['Escola'] + '2024' + pdt['Marca'] + pdt['Serial']
            pdt['SKU'] = pdt['SKU'].str.replace(' ','')
            pdt = pdt.drop_duplicates()
            
            
            operacoes = pdt[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome',2024,'2024+','Descrição Magento','Quantidade de alunos','% Desconto Volume','% Desconto Extra','% Desconto Total','Customer Group','Squad']]
            operacoes = operacoes.rename(columns = {2024:'Cód Itens'} )
            solucao = operacoes.copy()
            operacao = operacoes.copy()
            operacao = operacao[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome','Cód Itens','Descrição Magento','Quantidade de alunos','% Desconto Volume','% Desconto Extra','% Desconto Total','Customer Group','Squad']]
            #operacao.to_excel('operacao.xlsx')
            operacao = operacao.sort_values(by=['Série','Bimestre'])

            
                

            solucao = solucao.groupby(['Escola','CNPJ','Série','Bimestre','Marca','Segmento','Ano','Público','Serial','SKU','Nome','Customer Group','Squad'])['2024+'].sum().reset_index()
            solucao['visibilidade'] = 'N'
            solucao['faturamento_produto'] = 'MATERIAL'
            solucao['cliente_produto'] = cliente_tipo
            solucao['ativar_restricao'] = 'S'
            #solucao.to_csv('teste_solução.csv')
            #solucao
            categoria = pd.read_excel(planilha, sheet_name='categoriab2b')
            solucao = pd.merge(solucao,categoria, on=['Série'], how='inner')
            #solucao
            solucao['Categorias'] = solucao['Marca'] + '/' + solucao['Categorias']
            solucao = solucao.sort_values(by=['Bimestre','Série'], ascending=True)
            solucao = solucao.rename(columns={'Público':'grupo_de_atributo','Marca':'marca_produto', 'Nome':'nome', 'SKU':'sku', 'Ano':'ano_produto', 'Série':'serie_produto', 'Bimestre':'utilizacao_produto', 'Categorias':'categorias', '2024+':'items', 'Customer Group':'grupos_permissao'})
            solucao['items'] = solucao['items'].apply(lambda x: x[:-1])
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            
            solucao['nome'] = solucao['nome'].str.replace('INFANTIL','EI')
            solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
            solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
            solucao['nome'] = solucao['nome'].str.replace('ENSINO MÉDIO','EM')
            operacao['Nome'] = operacao['Nome'].str.replace('INFANTIL','EI')
            operacao['Nome'] = operacao['Nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
            operacao['Nome'] = operacao['Nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
            operacao['Nome'] = operacao['Nome'].str.replace('ENSINO MÉDIO','EM')
            
            
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('°','º')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 1','1 ANO')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 2','2 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 3','3 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 4','4 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 5','5 ANOS')

            
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('série','SÉRIE')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('ano','ANO')
            solucao['nome'] = solucao['nome'].str.replace('°','º')
            solucao['publico_produto'] = 'ALUNO'
            

            solucao.loc[(solucao['nome'].str.contains('BIMESTRE')) , ['periodo_produto']] = 'BIMESTRAL'
            solucao.loc[(solucao['nome'].str.contains('ANUAL')) , ['periodo_produto']] = 'ANUAL'
            solucao.loc[(solucao['nome'].str.contains('SEMESTRAL')) , ['periodo_produto']] = 'SEMESTRAL'
            solucao.loc[(solucao['serie_produto'].str.contains('Semi')) , ['periodo_produto']] = 'SEMESTRAL'

            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','periodo_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            df_brinde = operacao[['CNPJ','SKU','Série','Bimestre','Descrição Magento','Cód Itens','Customer Group']]
            df_brinde_input = pd.read_excel(planilha, sheet_name='brinde')
            df_brinde = pd.merge(df_brinde,df_brinde_input, on=['Cód Itens'], how='inner')
            df_brinde_final = df_brinde.copy()
            df_brinde_final = df_brinde_final[['Série_x','Nome da Regra','Customer Group','SKU_x','SKU_y']]
            df_brinde_final['Status'] = 'ATIVO'
            df_brinde_infantil = df_brinde_final.loc[df_brinde_final['Série_x'].str.contains('Grupo')]
            df_brinde_infantil['Qtd Incremento'] = 11
            df_brinde_demais = df_brinde_final.loc[~df_brinde_final['Série_x'].str.contains('Grupo')]
            df_brinde_demais['Qtd Incremento'] = 20
            df_brinde_final = pd.concat([df_brinde_infantil,df_brinde_demais])
            df_brinde_final['Qtd Condicao'] = 1
            df_brinde_final = df_brinde_final.rename(columns={'Customer Group':'Grupo do Cliente','SKU_x':'Sku Condicao','SKU_y':'Sku Brinde'})
            df_brinde_final = df_brinde_final[['Nome da Regra','Status','Grupo do Cliente','Sku Condicao','Qtd Condicao','Sku Brinde','Qtd Incremento']]
            df_brinde_final = df_brinde_final.rename(columns= {'Nome da Regra':'nome_da_regra','Status':'status','Grupo do Cliente':'grupo_do_cliente',
                                                               'Sku Condicao':'sku_condicao','Qtd Condicao':'qtd_condicao','Sku Brinde':'sku_brinde','Qtd Incremento':'qtd_incremento'})
            df_brinde_final = df_brinde_final.sort_values(by=['grupo_do_cliente','nome_da_regra'])
            df_brinde_final['id'] = ''
            df_brinde_final = df_brinde_final[['id','nome_da_regra','status','grupo_do_cliente','sku_condicao','qtd_condicao','sku_brinde','qtd_incremento']]
            ######## Exibir na tela para conferência #####
            escola = operacao['Escola'].unique()[0]
            df_brinde_h5 = df_brinde_final.loc[df_brinde_final['nome_da_regra'].str.contains('H5')]
            #df_brinde_h5
            df_brinde_final2 = df_brinde_final.copy()
            

            #### subir nas demais orreções
            solucao['nome'] = solucao['nome'].str.replace('Grupo 1','1 ANO')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 2','2 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 3','3 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 4','4 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 5','5 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('série','SÉRIE')
            solucao['nome'] = solucao['nome'].str.replace('ano','ANO')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 1','1 ANO')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 2','2 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 3','3 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 4','4 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 5','5 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('série','SÉRIE')
            operacao['Nome'] = operacao['Nome'].str.replace('ano','ANO')
            operacao['Série'] = operacao['Série'].str.replace('°','º')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 1','1 ANO')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 2','2 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 3','3 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 4','4 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 5','5 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('ano','ANO')
            operacao['Série'] = operacao['Série'].str.replace('série','SÉRIE')
            operacao['Segmento'] = operacao['Segmento'].str.replace('FUNDAMENTAL ANOS INICIAIS','FUNDAMENTAL I')
            operacao['Segmento'] = operacao['Segmento'].str.replace('FUNDAMENTAL ANOS FINAIS','FUNDAMENTAL II')
            operacao['Nome'] = operacao['Nome'].str.replace('Extensivo','EXTENSIVO')
            operacao['Nome'] = operacao['Nome'].str.replace('Semi','SEMI')
            operacao['Série'] = operacao['Série'].str.replace('Extensivo','PRE VESTIBULAR')
            operacao['Série'] = operacao['Série'].str.replace('Semi','SEMI EXTENSIVO II')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Extensivo','PRE VESTIBULAR')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Semi','SEMI EXTENSIVO II')
            solucao['nome'] = solucao['nome'].str.replace('Extensivo','EXTENSIVO')
            solucao['nome'] = solucao['nome'].str.replace('Semi','SEMI')

            solucao = solucao.rename(columns={'utilizacao_produto':'utilizacao_produto2','periodo_produto':'periodo_produto2'})
            solucao = solucao.rename(columns={'utilizacao_produto2':'periodo_produto','periodo_produto2':'utilizacao_produto'})
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','periodo_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            
            solucao.loc[solucao['periodo_produto'].str.contains('ANUAL'), ['periodo_produto']] = '1º BIMESTRE'

            ope3bim = operacao.loc[operacao['Bimestre'] == '3º BIMESTRE']
            #ope3bim
            sol3bim = solucao.loc[solucao['periodo_produto'] == '3º BIMESTRE']
            #sol3bim
            brinde3bim = df_brinde_final.loc[df_brinde_final['nome_da_regra'].str.contains('3º BIMESTRE')]
            #brinde3bim

            
            #operacao
            st.divider()

            with st.spinner('Aguarde...'):
                time.sleep(3)

            st.success('Concluído com sucesso!', icon="✅")
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv(index=False).encode('UTF-8')
            
            
            col1, col2, col3 = st.columns(3)
            with col1:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    operacao.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                st.download_button(
                        label="Download do cadastro (XLSX)",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-operacao.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    ope3bim.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                st.download_button(
                        label="Download do cadastro 3º Bimestre (XLSX)",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-3bim.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col2:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    solucao.to_excel(writer, index=False, sheet_name='Sheet1')
                # Configurar os parâmetros para o botão de download
                st.download_button(
                    label="Download Solução (XLSX)",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-solucao.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                solucao = convert_df(solucao)
                st.download_button(
                label="Download Solução (CSV) ",
                    data=solucao,
                    file_name=f'{today}-{escola}-solucao_import.csv',
                    mime='text/csv'
                )
                #solucao 3bim
                sol3bim = convert_df(sol3bim)
                st.download_button(
                label="Download Solução 3º Bimestre (CSV)",
                    data=sol3bim,
                    file_name=f'{today}-{escola}-solucao_import_3bim.csv',
                    mime='text/csv'
                )
                    
            with col3:
                df_brinde_final = convert_df(df_brinde_final)
                st.download_button(
                label="Download do brinde (CSV)",
                    data=df_brinde_final,
                    file_name=f'{today}-{escola}-brinde_import.csv',
                    mime='text/csv'
                )
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_brinde_final2.to_excel(writer, index=False, sheet_name='Sheet1')
                # Configurar os parâmetros para o botão de download
                st.download_button(
                    label="Download do brinde (XLSX)",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-brinde.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                df_brinde_h5 = convert_df(df_brinde_h5)
                st.download_button(
                label="Download do brinde H5 (CSV)",
                    data=df_brinde_h5,
                    file_name=f'{today}-{escola}-brinde_h5_import.csv',
                    mime='text/csv'
                )
                #3BIMESTRE BRINDE
                brinde3bim = convert_df(brinde3bim)
                st.download_button(
                label="Download Brinde 3º Bimestre (CSV)",
                    data=brinde3bim,
                    file_name=f'{today}-{escola}-brinde_import_3bimes.csv',
                    mime='text/csv'
                )
               

            ###### DEBUG COM FILTRO
            st.divider()
            st.write("Cliente:", escola)
            st.divider()
            st.write('Resultado:')
            filter = operacao[['Escola','Marca','Segmento','Série','Bimestre','Nome','Descrição Magento','Quantidade de alunos','Customer Group']]
            selected = st.selectbox('Selecione a série:', ['',*filter['Série'].unique()])
            if selected:
                selected_serie = filter[filter['Série'] == selected]
                selected_serie
            else:
                filter
            ##################


##########################################################################################################################################################
##########################################################################################################################################################


    if choice == 'CONEXIA B2C':
        marca = 'AZ B2C'
        planilha = 'itens.xlsx'
        today = date.today().strftime('%d-%m-%Y')
        cliente_tipo = 'B2C'

        st.warning("Simulador - CONEXIA B2C")
        #agree = st.checkbox('Marque para usar o cálculo do script')
        #  29.271.264/0001-61
        cliente = st.text_input('Digite o CNPJ da escola:')
        # Carrega o arquivo
        file = st.file_uploader("Selecione um arquivo Excel", type=["xlsm"])
        
        if file is not None:
            simul0 = pd.read_excel(file, sheet_name='cálculos B2C')
            simul0=simul0.assign(Bimestre="1º BIMESTRE")
            simul0.replace(0, np.nan, inplace = True)
            

            simul2 = pd.read_excel(file, sheet_name='cálculos 2º Bim')
            simul2=simul2.assign(Bimestre="2º BIMESTRE")
            simul2.replace(0, np.nan, inplace = True)

            simul3 = pd.read_excel(file, sheet_name='cálculos 3º Bim')
            simul3=simul3.assign(Bimestre="3º BIMESTRE")
            simul3.replace(0, np.nan, inplace = True)

            simul4 = pd.read_excel(file, sheet_name='cálculos 4º Bim')
            simul4=simul4.assign(Bimestre="4º BIMESTRE")
            simul4.replace(0, np.nan, inplace = True)

            #alterar regra conforme leitura das planilhas
            simul = pd.concat([simul0,simul2,simul3,simul4])
            
            simul = simul[simul['Quantidade de alunos']>0]
            simul = simul.rename(columns={'% Desconto Extra%':'% Desconto Extra'})
            
            simul = simul.rename(columns={'Construindo a Alfabetização':'Alfabetização','Itinerários Formativos Micro cursos     (2 IF)':'Itinerários',
                                          'H5 - (3 Horas) Language Book + CLIL e PBL ':'H5 - 3 Horas','H5 - (2 horas)\nInternational Journey + \nApp H5':'H5 - 2 horas Journey',
                                          'H5 Plus\n (3 horas extras)':'H5 Plus','My Life\n(Base)':'My Life - Base','My Life\n(2024)':'My Life - 2024',
                                          'Binoculo By Tell Me\n(Base)':'Binoculo - Base','Educacross Ed. Infantil\n(Base)':'Educacross Infantil - Base',
                                          'Educacross\n(Base)':'Educacross - Base','Educacross AZ\n(Base)':'Educacross AZ - Base','Educacross H5\n(Base)':'Educacross H5 - Base',
                                          'Ubbu\n(Base)':'Ubbu - Base','Binoculo By Tell Me\n(2024)':'Binoculo - 2024','Educacross Ed. Infantil\n(2024)':'Educacross Infantil - 2024',
                                          'Educacross\n(2024)':'Educacross - 2024','Educacross AZ\n(2024)':'Educacross AZ - 2024','Educacross H5\n(2024)':'Educacross H5 - 2024',
                                          'Ubbu\n(2024)':'Ubbu - 2024','Árvore\n(1 Módulo)':'Árvore 1 Módulo','Árvore\n(2 Módulos)':'Árvore 2 Módulos','Árvore\n(3 Módulos)':'Árvore 3 Módulos',
                                          'total aluno/ano\nsem desconto':'total aluno sem desconto','total aluno/ano\ncom desconto sem complementar':'total aluno com desconto sem complementar',
                                          'total aluno/ano\ncom desconto + Complementares':'total aluno com desconto com Complementares',})
            
            simul = simul[['Série','Segmento','Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês','% Desconto Volume','Quantidade de alunos','Razão Social','CNPJ','Squad','Tipo','Grupo de cliente','Bimestre','% Desconto Total','Valor de venda (B2C)']]
            
            simulador = simul.copy()
            df_cliente = simulador.loc[simulador['CNPJ'].str.strip() == cliente]
            df_cliente = df_cliente.fillna(0)
            df_cliente['Plataforma AZ'] = df_cliente['Plataforma AZ'].where(df_cliente['Plataforma AZ'] == 0, 1)
            df_cliente['Materiais Impressos AZ'] = df_cliente['Materiais Impressos AZ'].where(df_cliente['Materiais Impressos AZ'] == 0, 1)
            df_cliente['Alfabetização'] = df_cliente['Alfabetização'].where(df_cliente['Alfabetização'] == 0, 1)
            df_cliente['Cantalelê'] = df_cliente['Cantalelê'].where(df_cliente['Cantalelê'] == 0, 1)
            df_cliente['Mundo Leitor'] = df_cliente['Mundo Leitor'].where(df_cliente['Mundo Leitor'] == 0, 1)
            df_cliente['4 Avaliações Nacionais'] = df_cliente['4 Avaliações Nacionais'].where(df_cliente['4 Avaliações Nacionais'] == 0, 1)
            df_cliente['1 Simulado ENEM'] = df_cliente['1 Simulado ENEM'].where(df_cliente['1 Simulado ENEM'] == 0, 1)
            df_cliente['5 Simulados ENEM'] = df_cliente['5 Simulados ENEM'].where(df_cliente['5 Simulados ENEM'] == 0, 1)
            df_cliente['1 Simulado Regional'] = df_cliente['1 Simulado Regional'].where(df_cliente['1 Simulado Regional'] == 0, 1)
            df_cliente['Itinerários'] = df_cliente['Itinerários'].where(df_cliente['Itinerários'] == 0, 1)
            df_cliente['H5 - 3 Horas'] = df_cliente['H5 - 3 Horas'].where(df_cliente['H5 - 3 Horas'] == 0, 1)
            df_cliente['H5 - 2 horas Journey'] = df_cliente['H5 - 2 horas Journey'].where(df_cliente['H5 - 2 horas Journey'] == 0, 1)
            df_cliente['H5 Plus'] = df_cliente['H5 Plus'].where(df_cliente['H5 Plus'] == 0, 1)
            df_cliente['My Life - Base'] = df_cliente['My Life - Base'].where(df_cliente['My Life - Base'] == 0, 1)
            df_cliente['My Life - 2024'] = df_cliente['My Life - 2024'].where(df_cliente['My Life - 2024'] == 0, 1)
            df_cliente['Binoculo - Base'] = df_cliente['Binoculo - Base'].where(df_cliente['Binoculo - Base'] == 0, 1)
            df_cliente['Educacross Infantil - Base'] = df_cliente['Educacross Infantil - Base'].where(df_cliente['Educacross Infantil - Base'] == 0, 1)
            df_cliente['Educacross - Base'] = df_cliente['Educacross - Base'].where(df_cliente['Educacross - Base'] == 0, 1)
            df_cliente['Educacross AZ - Base'] = df_cliente['Educacross AZ - Base'].where(df_cliente['Educacross AZ - Base'] == 0, 1)
            df_cliente['Educacross H5 - Base'] = df_cliente['Educacross H5 - Base'].where(df_cliente['Educacross H5 - Base'] == 0, 1)
            df_cliente['Ubbu - Base'] = df_cliente['Ubbu - Base'].where(df_cliente['Ubbu - Base'] == 0, 1)
            df_cliente['Binoculo - 2024'] = df_cliente['Binoculo - 2024'].where(df_cliente['Binoculo - 2024'] == 0, 1)
            df_cliente['Educacross Infantil - 2024'] = df_cliente['Educacross Infantil - 2024'].where(df_cliente['Educacross Infantil - 2024'] == 0, 1)
            df_cliente['Educacross - 2024'] = df_cliente['Educacross - 2024'].where(df_cliente['Educacross - 2024'] == 0, 1)
            df_cliente['Educacross AZ - 2024'] = df_cliente['Educacross AZ - 2024'].where(df_cliente['Educacross AZ - 2024'] == 0, 1)
            df_cliente['Educacross H5 - 2024'] = df_cliente['Educacross H5 - 2024'].where(df_cliente['Educacross H5 - 2024'] == 0, 1)
            df_cliente['Ubbu - 2024'] = df_cliente['Ubbu - 2024'].where(df_cliente['Ubbu - 2024'] == 0, 1)
            df_cliente['Árvore 1 Módulo'] = df_cliente['Árvore 1 Módulo'].where(df_cliente['Árvore 1 Módulo'] == 0, 1)
            df_cliente['Árvore 2 Módulos'] = df_cliente['Árvore 2 Módulos'].where(df_cliente['Árvore 2 Módulos'] == 0, 1)
            df_cliente['Árvore 3 Módulos'] = df_cliente['Árvore 3 Módulos'].where(df_cliente['Árvore 3 Módulos'] == 0, 1)
            df_cliente['School Guardian'] = df_cliente['School Guardian'].where(df_cliente['School Guardian'] == 0, 1)
            df_cliente['Tindin'] = df_cliente['Tindin'].where(df_cliente['Tindin'] == 0, 1)
            df_cliente['Scholastic Earlybird and Bookflix'] = df_cliente['Scholastic Earlybird and Bookflix'].where(df_cliente['Scholastic Earlybird and Bookflix'] == 0, 1)
            df_cliente['Scholastic Literacy Pro'] = df_cliente['Scholastic Literacy Pro'].where(df_cliente['Scholastic Literacy Pro'] == 0, 1)
            df_cliente['Livro de Inglês'] = df_cliente['Livro de Inglês'].where(df_cliente['Livro de Inglês'] == 0,1)
            
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ed. Infantil','INFANTIL')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Iniciais','FUNDAMENTAL ANOS INICIAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Finais','FUNDAMENTAL ANOS FINAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ensino Médio','ENSINO MÉDIO')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('PV','PRÉ VESTIBULAR')
            df_cliente=df_cliente.assign(Extra="")
           
            ###regra do AZ e Plataforma
            df_cliente.loc[(df_cliente['Plataforma AZ'] == 1) & (df_cliente['Materiais Impressos AZ'] == 1), ['Plataforma AZ']] = 0
            
            ####
            ####regra do h5
            df_cliente.loc[(df_cliente['H5 Plus'] == 1) & (df_cliente['H5 - 2 horas Journey'] == 1), ['H5 - 2 horas Journey','H5 - 3 Horas']] = 0
            df_cliente.loc[(df_cliente['H5 - 2 horas Journey'] == 1) & (df_cliente['H5 - 3 Horas'] == 1), ['H5 - 3 Horas']] = 0



            df_client = df_cliente.copy()
            lista = ['Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM',
                     '5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024',
                     'Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024',
                     'Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês']
            
            #df_client.to_excel('cliente.xlsx')
         

            for item in lista:
                df_client.loc[df_client[item] == 1.0, item] = item
            COLUNAS = ['Série', 'Segmento','% Desconto Total','Valor de venda (B2C)','Quantidade de alunos','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra']
            p = pd.DataFrame(columns=COLUNAS)
            
            for i in lista:
                data = df_client[df_client[i] == i].groupby(['Série', 'Segmento','% Desconto Total','Valor de venda (B2C)','Quantidade de alunos','Razão Social','CNPJ','Squad','Tipo','Bimestre',i])['Extra'].count().reset_index()
                data = data.rename(columns={i: 'Produto'})
                p = pd.concat([p,data])
            p = p.sort_values(by=['Série'])
            p = p.reset_index()
            p = p.drop(columns=['index'])
            
            
            itens = pd.read_excel(planilha, sheet_name='itens_b2c')
            itens = itens[['MARCA',2024,'2024+','Produto','DESCRIÇÃO MAGENTO (B2C e B2B)','BIMESTRE','SEGMENTO','SÉRIE','PÚBLICO','TIPO DE FATURAMENTO']]
            itens = itens.rename(columns={'MARCA':'Marca','DESCRIÇÃO MAGENTO (B2C e B2B)':'Descrição Magento','BIMESTRE':'Bimestre','SEGMENTO':'Segmento','SÉRIE':'Série','PÚBLICO':'Público','TIPO DE FATURAMENTO':'Faturamento'})
            itens = itens[(itens['Marca'] == marca) | (itens['Marca'] == 'CONEXIA') | (itens['Marca'] == 'MUNDO LEITOR') | (itens['Marca'] == 'MY LIFE')| (itens['Marca'] == 'HIGH FIVE')]
        
          

            pdt = pd.merge(p, itens, on=['Série','Bimestre','Segmento','Produto'], how='inner')
        
        
            cod_serial = pd.read_excel(planilha, sheet_name='cod_serial')
            pdt = pd.merge(pdt, cod_serial, on=['Série','Bimestre','Segmento','Público'], how='inner')


            pdt['Ano'] = '2024'
            pdt['SKU'] = pdt['Ano'] + pdt['Serial']
            pdt = pdt[['Série','Segmento','% Desconto Total','Valor de venda (B2C)','Quantidade de alunos','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra','Produto','Marca',2024,'2024+','Descrição Magento','Público','Faturamento','Serial','Categoria','Ano','SKU']]
            

            h = re.compile(r'[../\-]')
            pdt['CNPJ_off'] = [h.sub('', x) for x in pdt['CNPJ']]
            pdt['CNPJ_off'] = [x.lstrip('0') for x in pdt['CNPJ_off']]
            pdt['CNPJ_off'] = pdt['CNPJ_off'].astype(float)
            
            cod_nome = pd.read_excel(planilha, sheet_name='nome_b2c')
            cod_nome['CNPJ_off'] = cod_nome['CNPJ_off'].astype(float)
            pdt = pd.merge(pdt, cod_nome, on=['CNPJ_off'], how='inner')
            
            ####################################################################################################
            ######NOVAS REGRAS DA SOLUÇÃO #####################################################
            pdt['Marca'] = pdt['Marca'].str.replace('MY LIFE','SOLUÇÃO')
            pdt['Marca'] = pdt['Marca'].str.replace('CONEXIA','SOLUÇÃO')
            pdt['Marca'] = pdt['Marca'].str.replace('HIGH FIVE','SOLUÇÃO')
            pdt['Marca'] = pdt['Marca'].str.replace('AZ B2C','SOLUÇÃO')

            ######NOVAS REGRAS POR SÉRIE#####################################################
            
         


            pdt['Nome'] = pdt['Marca']  + ' - ' + pdt['Escola'] + ' - ' + pdt['Segmento'] + ' - ' + pdt['Série'] + ' - ' + pdt['Bimestre']
            pdt['Marca'] = pdt['Marca'].str.replace('AZ B2C','AZ')
            pdt['SKU'] = pdt['Escola'] + "2024" + pdt['Serial']
            pdt['SKU'] = pdt['SKU'].str.replace(' ','')
            
            
            operacoes = pdt[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome',2024,'2024+','Descrição Magento','Quantidade de alunos','Valor de venda (B2C)','% Desconto Total','Customer Group','Squad']]
            operacoes = operacoes.rename(columns = {2024:'Cód Itens'} )
            solucao = operacoes.copy()
            operacao = operacoes.copy()
            operacao = operacao[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome','Cód Itens','Descrição Magento','Quantidade de alunos','% Desconto Total','Valor de venda (B2C)','Customer Group','Squad']]
            #operacao.to_excel('operacao.xlsx')
            operacao = operacao.sort_values(by=['Série','Bimestre'])
  

            solucao = solucao.groupby(['Escola','CNPJ','Série','Bimestre','Marca','Segmento','Ano','Público','Serial','SKU','Nome','Customer Group','Squad'])['2024+'].sum().reset_index()
            solucao['visibilidade'] = 'N'
            solucao['faturamento_produto'] = 'MATERIAL'
            solucao['cliente_produto'] = cliente_tipo
            solucao['ativar_restricao'] = 'S'
            #solucao.to_csv('teste_solução.csv')

            categoria = pd.read_excel(planilha, sheet_name='categoriab2c')
            solucao = pd.merge(solucao,categoria, on=['Série'], how='inner')
            solucao['Categorias'] = solucao['Categorias']
            solucao = solucao.sort_values(by=['Bimestre','Série'], ascending=True)
            solucao = solucao.rename(columns={'Público':'grupo_de_atributo','Marca':'marca_produto', 'Nome':'nome', 'SKU':'sku', 'Ano':'ano_produto', 'Série':'serie_produto', 'Bimestre':'utilizacao_produto', 'Categorias':'categorias', '2024+':'items', 'Customer Group':'grupos_permissao'})
            solucao['items'] = solucao['items'].apply(lambda x: x[:-1])
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            
            solucao['nome'] = solucao['nome'].str.replace('INFANTIL','EI')
            solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
            solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
            solucao['nome'] = solucao['nome'].str.replace('ENSINO MÉDIO','EM')
            operacao['Nome'] = operacao['Nome'].str.replace('INFANTIL','EI')
            operacao['Nome'] = operacao['Nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
            operacao['Nome'] = operacao['Nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
            operacao['Nome'] = operacao['Nome'].str.replace('ENSINO MÉDIO','EM')
            
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('°','º')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 1','1 ANO')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 2','2 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 3','3 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 4','4 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 5','5 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('série','SÉRIE')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('ano','ANO')
            solucao['nome'] = solucao['nome'].str.replace('°','º')
            solucao['publico_produto'] = 'ALUNO'
            
            solucao.loc[(solucao['nome'].str.contains('BIMESTRE')) , ['periodo_produto']] = 'BIMESTRAL'
            solucao.loc[(solucao['nome'].str.contains('ANUAL')) , ['periodo_produto']] = 'ANUAL'
            solucao.loc[(solucao['nome'].str.contains('SEMESTRAL')) , ['periodo_produto']] = 'SEMESTRAL'
            solucao.loc[(solucao['serie_produto'].str.contains('Semi')) , ['periodo_produto']] = 'SEMESTRAL'

            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','periodo_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            #solucao

            #### subir nas demais orreções
            solucao['nome'] = solucao['nome'].str.replace('Grupo 1','1 ANO')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 2','2 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 3','3 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 4','4 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 5','5 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('série','SÉRIE')
            solucao['nome'] = solucao['nome'].str.replace('ano','ANO')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 1','1 ANO')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 2','2 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 3','3 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 4','4 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 5','5 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('série','SÉRIE')
            operacao['Nome'] = operacao['Nome'].str.replace('ano','ANO')
            operacao['Série'] = operacao['Série'].str.replace('°','º')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 1','1 ANO')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 2','2 ANOS') 
            operacao['Série'] = operacao['Série'].str.replace('Grupo 3','3 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 4','4 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 5','5 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('ano','ANO')
            operacao['Série'] = operacao['Série'].str.replace('série','SÉRIE')
            operacao['Segmento'] = operacao['Segmento'].str.replace('FUNDAMENTAL ANOS INICIAIS','FUNDAMENTAL I')
            operacao['Segmento'] = operacao['Segmento'].str.replace('FUNDAMENTAL ANOS FINAIS','FUNDAMENTAL II')
            operacao['Nome'] = operacao['Nome'].str.replace('Extensivo','EXTENSIVO')
            operacao['Nome'] = operacao['Nome'].str.replace('Semi','SEMI')
            operacao['Série'] = operacao['Série'].str.replace('Extensivo','PRE VESTIBULAR')
            operacao['Série'] = operacao['Série'].str.replace('Semi','SEMI EXTENSIVO II')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Extensivo','PRE VESTIBULAR')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Semi','SEMI EXTENSIVO II')
            solucao['nome'] = solucao['nome'].str.replace('Extensivo','EXTENSIVO')
            solucao['nome'] = solucao['nome'].str.replace('Semi','SEMI')
            solucao['sku'] = solucao['sku'].str.replace('SOLUÇÃO','SOLUCAO')
            operacao['SKU'] = operacao['SKU'].str.replace('SOLUÇÃO','SOLUCAO')
            #operacao

            solucao = solucao.rename(columns={'utilizacao_produto':'utilizacao_produto2','periodo_produto':'periodo_produto2'})
            solucao = solucao.rename(columns={'utilizacao_produto2':'periodo_produto','periodo_produto2':'utilizacao_produto'})
            #solucao = solucao.rename(columns={'utilizacao_produto2':'periodo_produto','periodo_produto2':'utilizacao_produto'})
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','periodo_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            
            solucao['nome'] = solucao['nome'].str.replace('1º BIMESTRE','ANUAL')



            ######## Exibir na tela para conferência #####
            escola = operacao['Escola'].unique()[0]
        
            #operacao

            st.divider()

            with st.spinner('Aguarde...'):
                time.sleep(3)

            st.success('Concluído com sucesso!', icon="✅")
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv(index=False).encode('UTF-8')
            
            
            
            
            col1, col2, col3 = st.columns(3)
            with col1:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    operacao.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                st.download_button(
                        label="Download do cadastro (XLSX)",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-operacao.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col2:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        solucao.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                    st.download_button(
                        label="Download Solução (XLSX)",
                        data=output.getvalue(),
                        file_name=f'{today}-{escola}-solucao.xlsx',
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                    solucao = convert_df(solucao)
                    st.download_button(
                    label="Download Solução (CSV)",
                        data=solucao,
                        file_name=f'{today}-{escola}-solucao_import.csv',
                        mime='text/csv'
                    )
            #with col3:
                    
                    
                    
            
            ###### DEBUG COM FILTRO
            st.divider()
            st.write("Cliente:", escola)
            st.divider()
            st.write('Resultado:')
            filter = operacao[['Escola','Marca','Segmento','Série','SKU','Bimestre','Nome','Descrição Magento','Quantidade de alunos','Customer Group']]
            selected = st.selectbox('Selecione a série:', ['',*filter['Série'].unique()])
            if selected:
                selected_serie = filter[filter['Série'] == selected]
                selected_serie
            else:
                filter


##########################################################################################################################################################
##########################################################################################################################################################


    if choice == 'SEB':
            
            marca = 'AZ' ## ou AZ SESC B2B ou AZ/SESC
            sheetname = 'itens_performance'
            planilha = 'itens.xlsx'
            today = date.today().strftime('%d-%m-%Y')
            cliente_tipo = 'B2B'

            st.success("Simulador - SEB")
            #agree = st.checkbox('Marque para usar o cálculo do script - (Não recomendado!)')
            agree = ''
            #  29.271.264/0001-61
            cliente = st.text_input('Digite o CNPJ da escola:')
            # Carrega o arquivo
            file = st.file_uploader("Selecione um arquivo Excel", type=["xlsm"])
            
            
            if file is not None:
                simul0 = pd.read_excel(file, sheet_name='cálculos Anual')
                simul0=simul0.assign(Bimestre="ANUAL")
                simul0.replace(0, np.nan, inplace = True)
                
                simul1 = pd.read_excel(file, sheet_name='cálculos 1º Bim')
                simul1=simul1.assign(Bimestre="1º BIMESTRE")
                simul1.replace(0, np.nan, inplace = True)

                simul2 = pd.read_excel(file, sheet_name='cálculos 2º Bim')
                simul2=simul2.assign(Bimestre="2º BIMESTRE")
                simul2.replace(0, np.nan, inplace = True)

                simul3 = pd.read_excel(file, sheet_name='cálculos 3º Bim')
                simul3=simul3.assign(Bimestre="3º BIMESTRE")
                simul3.replace(0, np.nan, inplace = True)

                simul4 = pd.read_excel(file, sheet_name='cálculos 4º Bim')
                simul4=simul4.assign(Bimestre="4º BIMESTRE")
                simul4.replace(0, np.nan, inplace = True)

                #alterar regra conforme leitura das planilhas
                simul = pd.concat([simul0,simul1,simul2,simul3,simul4])
                
                #simul = pd.concat([simul1,simul2,simul3,simul4])
                simul = simul[simul['Quantidade de alunos']>0]
                #simul
                if agree:
                    desconto = pd.read_excel(file, sheet_name='Formulário Anual 2024')
                    desconto = desconto.iloc[:, :6] 
                    desconto = desconto[['FORMULÁRIO DE AQUISIÇÃO DE MATERIAL DIDÁTICO','Unnamed: 5']]
                    desconto = desconto.rename(columns={'FORMULÁRIO DE AQUISIÇÃO DE MATERIAL DIDÁTICO':'Série','Unnamed: 5':'% Desconto Extra%'})
                
                    indice = [25,26,27,28,29,47,48,49,50,51,67,68,69,70,84,85,98,112,113]
                    desconto = desconto.iloc[indice]

                    del(simul['% Desconto Extra'])
                    del(simul['% Desconto Total'])
                    simul = simul.drop_duplicates()
                    
                    simul = pd.merge(simul, desconto, on=['Série'], how='inner')
                    simul['% Desconto Volume'] = simul['% Desconto Volume'].apply(lambda x: x[:-1])
                    simul['% Desconto Volume'] = simul['% Desconto Volume'].astype('float64')/100
                    simul['% Desconto Total'] = simul['% Desconto Extra%'] + simul['% Desconto Volume']
                    simul = simul.rename(columns={'% Desconto Extra%':'% Desconto Extra'})
                
                simul = simul.rename(columns={'Construindo a Alfabetização':'Alfabetização','Itinerários Formativos Micro cursos     (2 IF)':'Itinerários','H5 - (3 Horas) Language Book + CLIL e PBL ':'H5 - 3 Horas','H5 - (2 horas)\nInternational Journey + \nApp H5':'H5 - 2 horas Journey','H5 Plus\n (3 horas extras)':'H5 Plus','My Life\n(Base)':'My Life - Base','My Life\n(2024)':'My Life - 2024','Binoculo By Tell Me\n(Base)':'Binoculo - Base','Educacross Ed. Infantil\n(Base)':'Educacross Infantil - Base','Educacross\n(Base)':'Educacross - Base','Educacross AZ\n(Base)':'Educacross AZ - Base','Educacross H5\n(Base)':'Educacross H5 - Base','Ubbu\n(Base)':'Ubbu - Base','Binoculo By Tell Me\n(2024)':'Binoculo - 2024','Educacross Ed. Infantil\n(2024)':'Educacross Infantil - 2024','Educacross\n(2024)':'Educacross - 2024','Educacross AZ\n(2024)':'Educacross AZ - 2024','Educacross H5\n(2024)':'Educacross H5 - 2024','Ubbu\n(2024)':'Ubbu - 2024','Árvore\n(1 Módulo)':'Árvore 1 Módulo','Árvore\n(2 Módulos)':'Árvore 2 Módulos','Árvore\n(3 Módulos)':'Árvore 3 Módulos','total aluno/ano\nsem desconto':'total aluno sem desconto','total aluno/ano\ncom desconto sem complementar':'total aluno com desconto sem complementar','total aluno/ano\ncom desconto + Complementares':'total aluno com desconto com Complementares',})
                
                simul = simul[['Série','Segmento','Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês','% Desconto Volume','Quantidade de alunos','Razão Social','CNPJ','Squad','Tipo','Observação','Grupo de cliente','Bimestre','% Desconto Extra','% Desconto Total']]
                #simul.to_excel('simul.xlsx')
                
                
                simulador = simul.copy()
                df_cliente = simulador.loc[simulador['CNPJ'].str.strip() == cliente]
                df_cliente = df_cliente.fillna(0)

                df_cliente['Plataforma AZ'] = df_cliente['Plataforma AZ'].where(df_cliente['Plataforma AZ'] == 0, 1)
                df_cliente['Materiais Impressos AZ'] = df_cliente['Materiais Impressos AZ'].where(df_cliente['Materiais Impressos AZ'] == 0, 1)
                df_cliente['Alfabetização'] = df_cliente['Alfabetização'].where(df_cliente['Alfabetização'] == 0, 1)
                df_cliente['Cantalelê'] = df_cliente['Cantalelê'].where(df_cliente['Cantalelê'] == 0, 1)
                df_cliente['Mundo Leitor'] = df_cliente['Mundo Leitor'].where(df_cliente['Mundo Leitor'] == 0, 1)
                df_cliente['4 Avaliações Nacionais'] = df_cliente['4 Avaliações Nacionais'].where(df_cliente['4 Avaliações Nacionais'] == 0, 1)
                df_cliente['1 Simulado ENEM'] = df_cliente['1 Simulado ENEM'].where(df_cliente['1 Simulado ENEM'] == 0, 1)
                df_cliente['5 Simulados ENEM'] = df_cliente['5 Simulados ENEM'].where(df_cliente['5 Simulados ENEM'] == 0, 1)
                df_cliente['1 Simulado Regional'] = df_cliente['1 Simulado Regional'].where(df_cliente['1 Simulado Regional'] == 0, 1)
                df_cliente['Itinerários'] = df_cliente['Itinerários'].where(df_cliente['Itinerários'] == 0, 1)
                df_cliente['H5 - 3 Horas'] = df_cliente['H5 - 3 Horas'].where(df_cliente['H5 - 3 Horas'] == 0, 1)
                df_cliente['H5 - 2 horas Journey'] = df_cliente['H5 - 2 horas Journey'].where(df_cliente['H5 - 2 horas Journey'] == 0, 1)
                df_cliente['H5 Plus'] = df_cliente['H5 Plus'].where(df_cliente['H5 Plus'] == 0, 1)
                df_cliente['My Life - Base'] = df_cliente['My Life - Base'].where(df_cliente['My Life - Base'] == 0, 1)
                df_cliente['My Life - 2024'] = df_cliente['My Life - 2024'].where(df_cliente['My Life - 2024'] == 0, 1)
                df_cliente['Binoculo - Base'] = df_cliente['Binoculo - Base'].where(df_cliente['Binoculo - Base'] == 0, 1)
                df_cliente['Educacross Infantil - Base'] = df_cliente['Educacross Infantil - Base'].where(df_cliente['Educacross Infantil - Base'] == 0, 1)
                df_cliente['Educacross - Base'] = df_cliente['Educacross - Base'].where(df_cliente['Educacross - Base'] == 0, 1)
                df_cliente['Educacross AZ - Base'] = df_cliente['Educacross AZ - Base'].where(df_cliente['Educacross AZ - Base'] == 0, 1)
                df_cliente['Educacross H5 - Base'] = df_cliente['Educacross H5 - Base'].where(df_cliente['Educacross H5 - Base'] == 0, 1)
                df_cliente['Ubbu - Base'] = df_cliente['Ubbu - Base'].where(df_cliente['Ubbu - Base'] == 0, 1)
                df_cliente['Binoculo - 2024'] = df_cliente['Binoculo - 2024'].where(df_cliente['Binoculo - 2024'] == 0, 1)
                df_cliente['Educacross Infantil - 2024'] = df_cliente['Educacross Infantil - 2024'].where(df_cliente['Educacross Infantil - 2024'] == 0, 1)
                df_cliente['Educacross - 2024'] = df_cliente['Educacross - 2024'].where(df_cliente['Educacross - 2024'] == 0, 1)
                df_cliente['Educacross AZ - 2024'] = df_cliente['Educacross AZ - 2024'].where(df_cliente['Educacross AZ - 2024'] == 0, 1)
                df_cliente['Educacross H5 - 2024'] = df_cliente['Educacross H5 - 2024'].where(df_cliente['Educacross H5 - 2024'] == 0, 1)
                df_cliente['Ubbu - 2024'] = df_cliente['Ubbu - 2024'].where(df_cliente['Ubbu - 2024'] == 0, 1)
                df_cliente['Árvore 1 Módulo'] = df_cliente['Árvore 1 Módulo'].where(df_cliente['Árvore 1 Módulo'] == 0, 1)
                df_cliente['Árvore 2 Módulos'] = df_cliente['Árvore 2 Módulos'].where(df_cliente['Árvore 2 Módulos'] == 0, 1)
                df_cliente['Árvore 3 Módulos'] = df_cliente['Árvore 3 Módulos'].where(df_cliente['Árvore 3 Módulos'] == 0, 1)
                df_cliente['School Guardian'] = df_cliente['School Guardian'].where(df_cliente['School Guardian'] == 0, 1)
                df_cliente['Tindin'] = df_cliente['Tindin'].where(df_cliente['Tindin'] == 0, 1)
                df_cliente['Scholastic Earlybird and Bookflix'] = df_cliente['Scholastic Earlybird and Bookflix'].where(df_cliente['Scholastic Earlybird and Bookflix'] == 0, 1)
                df_cliente['Scholastic Literacy Pro'] = df_cliente['Scholastic Literacy Pro'].where(df_cliente['Scholastic Literacy Pro'] == 0, 1)
                df_cliente['Livro de Inglês'] = df_cliente['Livro de Inglês'].where(df_cliente['Livro de Inglês'] == 0,1)

                
                df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ed. Infantil','INFANTIL')
                df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Iniciais','FUNDAMENTAL ANOS INICIAIS')
                df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Finais','FUNDAMENTAL ANOS FINAIS')
                df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ensino Médio','ENSINO MÉDIO')
                df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('PV','PRÉ VESTIBULAR')
                df_cliente=df_cliente.assign(Extra="")
                
                ###regra do AZ e Plataforma
                df_cliente.loc[(df_cliente['Plataforma AZ'] == 1) & (df_cliente['Materiais Impressos AZ'] == 1), ['Plataforma AZ']] = 0
                

                ####regra do h5
                df_cliente.loc[(df_cliente['H5 Plus'] == 1) & (df_cliente['H5 - 2 horas Journey'] == 1), ['H5 - 2 horas Journey','H5 - 3 Horas']] = 0
                df_cliente.loc[(df_cliente['H5 - 2 horas Journey'] == 1) & (df_cliente['H5 - 3 Horas'] == 1), ['H5 - 3 Horas']] = 0
            

                ####
                df_client = df_cliente.copy()
                lista = ['Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês']
                
                #df_client.to_excel('cliente.xlsx')
                #df_client

                for item in lista:
                    df_client.loc[df_client[item] == 1.0, item] = item
                COLUNAS = ['Série', 'Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra']
                p = pd.DataFrame(columns=COLUNAS)
                
                for i in lista:
                    data = df_client[df_client[i] == i].groupby(['Série', 'Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Squad','Tipo','Bimestre',i])['Extra'].count().reset_index()
                    data = data.rename(columns={i: 'Produto'})
                    p = pd.concat([p,data])
                p = p.sort_values(by=['Série'])
                p = p.reset_index()
                p = p.drop(columns=['index'])
                p = p.drop_duplicates()
                
                
                
                itens = pd.read_excel(planilha, sheet_name=sheetname)
                itens = itens[['MARCA',2024,'2024+','Produto','DESCRIÇÃO MAGENTO (B2C e B2B)','BIMESTRE','SEGMENTO','SÉRIE','PÚBLICO','TIPO DE FATURAMENTO']]
                itens = itens.rename(columns={'MARCA':'Marca','DESCRIÇÃO MAGENTO (B2C e B2B)':'Descrição Magento','BIMESTRE':'Bimestre','SEGMENTO':'Segmento','SÉRIE':'Série','PÚBLICO':'Público','TIPO DE FATURAMENTO':'Faturamento'})
                itens = itens[(itens['Marca'] == marca) | (itens['Marca'] == 'CONEXIA') | (itens['Marca'] == 'MUNDO LEITOR') | (itens['Marca'] == 'MY LIFE')| (itens['Marca'] == 'HIGH FIVE')]
                
                pdt = pd.merge(p, itens, on=['Série','Bimestre','Segmento','Produto'], how='inner')
                
                cod_serial = pd.read_excel(planilha, sheet_name='cod_serial')
                
                pdt = pd.merge(pdt, cod_serial, on=['Série','Bimestre','Segmento','Público'], how='inner')

                pdt['Ano'] = '2024'
                pdt['SKU'] = pdt['Ano'] + pdt['Serial']
                pdt = pdt[['Série','Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra','Produto','Marca',2024,'2024+','Descrição Magento','Público','Faturamento','Serial','Categoria','Ano','SKU']]
                

                h = re.compile(r'[../\-]')
                pdt['CNPJ_off'] = [h.sub('', x) for x in pdt['CNPJ']]
                pdt['CNPJ_off'] = [x.lstrip('0') for x in pdt['CNPJ_off']]
                pdt['CNPJ_off'] = pdt['CNPJ_off'].astype(float)
                
                cod_nome = pd.read_excel(planilha, sheet_name='nome')
                cod_nome['CNPJ_off'] = cod_nome['CNPJ_off'].astype(float)
                pdt = pd.merge(pdt, cod_nome, on=['CNPJ_off'], how='inner')
                
                
                ####################################################################################################
                ######NOVAS REGRAS POR SÉRIE#####################################################
                
                serie = pdt['Série'].unique()
                pdt_final = []
                pdt = pdt[~((pdt['Marca'] == 'CONEXIA') & (pdt['Bimestre'].str.contains('BIMESTRE')))]
                for i in serie:
                    pdt_serie = pdt.loc[pdt['Série'] == i]
                    pdt_serie['Marca'] = pdt_serie['Marca'].str.replace('MUNDO LEITOR','AZ')    

                    #if (pdt_serie['Marca'].str.contains('CONEXIA').any()):
                      #      pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'CONEXIA') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]   
                      # 
                    if (pdt_serie['Marca'].str.contains('HIGH FIVE').any()):
                            pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'HIGH FIVE') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]              
                    
                    if (pdt_serie['Marca'].str.contains('MY LIFE').any()):
                            pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'MY LIFE') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]
                    
                    pdt_final.append(pdt_serie)
                    pdt_full = pd.concat(pdt_final)

                #pdt_full = pdt_full[~((pdt_full['Marca'] == 'AZ') & (pdt_full['Bimestre'].str.contains('ANUAL')))]
                pdt = pdt_full.copy()
                pdt.loc[pdt['Marca'] == 'MY LIFE', ['Marca']] = 'DIGITAL'
                pdt.loc[pdt['Marca'] == 'CONEXIA', ['Marca']] = 'DIGITAL'

                pdt.loc[(pdt['Marca'] == 'AZ')&(pdt['Bimestre'].str.contains('ANUAL')), ['Bimestre']] = '1º BIMESTRE'

                ######End Regra   
            
                pdt['Nome'] = 'SOLUÇÃO ' + pdt['Marca']  + ' - ' + pdt['Escola'] + ' - ' + pdt['Segmento'] + ' - ' + pdt['Série'] + ' - ' + pdt['Bimestre']
                pdt['SKU'] = pdt['Escola'] + '2024' + pdt['Marca'] + pdt['Serial']
                pdt['SKU'] = pdt['SKU'].str.replace(' ','')
                pdt = pdt.drop_duplicates()
                
                
                operacoes = pdt[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome',2024,'2024+','Descrição Magento','Quantidade de alunos','% Desconto Volume','% Desconto Extra','% Desconto Total','Customer Group','Squad']]
                operacoes = operacoes.rename(columns = {2024:'Cód Itens'} )
                solucao = operacoes.copy()
                operacao = operacoes.copy()
                operacao = operacao[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome','Cód Itens','Descrição Magento','Quantidade de alunos','% Desconto Volume','% Desconto Extra','% Desconto Total','Customer Group','Squad']]
                #operacao.to_excel('operacao.xlsx')
                operacao = operacao.sort_values(by=['Série','Bimestre'])

                
                    

                solucao = solucao.groupby(['Escola','CNPJ','Série','Bimestre','Marca','Segmento','Ano','Público','Serial','SKU','Nome','Customer Group','Squad'])['2024+'].sum().reset_index()
                solucao['visibilidade'] = 'N'
                solucao['faturamento_produto'] = 'MATERIAL'
                solucao['cliente_produto'] = cliente_tipo
                solucao['ativar_restricao'] = 'S'
                #solucao.to_csv('teste_solução.csv')
                #solucao
                categoria = pd.read_excel(planilha, sheet_name='categoriab2b')
                solucao = pd.merge(solucao,categoria, on=['Série'], how='inner')
                #solucao
                solucao['Categorias'] = solucao['Marca'] + '/' + solucao['Categorias']
                solucao = solucao.sort_values(by=['Bimestre','Série'], ascending=True)
                solucao = solucao.rename(columns={'Público':'grupo_de_atributo','Marca':'marca_produto', 'Nome':'nome', 'SKU':'sku', 'Ano':'ano_produto', 'Série':'serie_produto', 'Bimestre':'utilizacao_produto', 'Categorias':'categorias', '2024+':'items', 'Customer Group':'grupos_permissao'})
                solucao['items'] = solucao['items'].apply(lambda x: x[:-1])
                solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
                
                solucao['nome'] = solucao['nome'].str.replace('INFANTIL','EI')
                solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
                solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
                solucao['nome'] = solucao['nome'].str.replace('ENSINO MÉDIO','EM')
                operacao['Nome'] = operacao['Nome'].str.replace('INFANTIL','EI')
                operacao['Nome'] = operacao['Nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
                operacao['Nome'] = operacao['Nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
                operacao['Nome'] = operacao['Nome'].str.replace('ENSINO MÉDIO','EM')
                
                
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('°','º')
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 1','1 ANO')
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 2','2 ANOS')
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 3','3 ANOS')
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 4','4 ANOS')
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 5','5 ANOS')

                
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('série','SÉRIE')
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('ano','ANO')
                solucao['nome'] = solucao['nome'].str.replace('°','º')
                solucao['publico_produto'] = 'ALUNO'
                

                solucao.loc[(solucao['nome'].str.contains('BIMESTRE')) , ['periodo_produto']] = 'BIMESTRAL'
                solucao.loc[(solucao['nome'].str.contains('ANUAL')) , ['periodo_produto']] = 'ANUAL'
                solucao.loc[(solucao['nome'].str.contains('SEMESTRAL')) , ['periodo_produto']] = 'SEMESTRAL'
                solucao.loc[(solucao['serie_produto'].str.contains('Semi')) , ['periodo_produto']] = 'SEMESTRAL'

                solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','periodo_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
                df_brinde = operacao[['CNPJ','SKU','Série','Bimestre','Descrição Magento','Cód Itens','Customer Group']]
                df_brinde_input = pd.read_excel(planilha, sheet_name='brinde')
                df_brinde = pd.merge(df_brinde,df_brinde_input, on=['Cód Itens'], how='inner')
                df_brinde_final = df_brinde.copy()
                df_brinde_final = df_brinde_final[['Série_x','Nome da Regra','Customer Group','SKU_x','SKU_y']]
                df_brinde_final['Status'] = 'ATIVO'
                df_brinde_infantil = df_brinde_final.loc[df_brinde_final['Série_x'].str.contains('Grupo')]
                df_brinde_infantil['Qtd Incremento'] = 11
                df_brinde_demais = df_brinde_final.loc[~df_brinde_final['Série_x'].str.contains('Grupo')]
                df_brinde_demais['Qtd Incremento'] = 20
                df_brinde_final = pd.concat([df_brinde_infantil,df_brinde_demais])
                df_brinde_final['Qtd Condicao'] = 1
                df_brinde_final = df_brinde_final.rename(columns={'Customer Group':'Grupo do Cliente','SKU_x':'Sku Condicao','SKU_y':'Sku Brinde'})
                df_brinde_final = df_brinde_final[['Nome da Regra','Status','Grupo do Cliente','Sku Condicao','Qtd Condicao','Sku Brinde','Qtd Incremento']]
                df_brinde_final = df_brinde_final.rename(columns= {'Nome da Regra':'nome_da_regra','Status':'status','Grupo do Cliente':'grupo_do_cliente',
                                                                'Sku Condicao':'sku_condicao','Qtd Condicao':'qtd_condicao','Sku Brinde':'sku_brinde','Qtd Incremento':'qtd_incremento'})
                df_brinde_final = df_brinde_final.sort_values(by=['grupo_do_cliente','nome_da_regra'])
                df_brinde_final['id'] = ''
                df_brinde_final = df_brinde_final[['id','nome_da_regra','status','grupo_do_cliente','sku_condicao','qtd_condicao','sku_brinde','qtd_incremento']]
                ######## Exibir na tela para conferência #####
                escola = operacao['Escola'].unique()[0]
                df_brinde_h5 = df_brinde_final.loc[df_brinde_final['nome_da_regra'].str.contains('H5')]
                #df_brinde_h5
                df_brinde_final2 = df_brinde_final.copy()
                

                #### subir nas demais orreções
                solucao['nome'] = solucao['nome'].str.replace('Grupo 1','1 ANO')
                solucao['nome'] = solucao['nome'].str.replace('Grupo 2','2 ANOS')
                solucao['nome'] = solucao['nome'].str.replace('Grupo 3','3 ANOS')
                solucao['nome'] = solucao['nome'].str.replace('Grupo 4','4 ANOS')
                solucao['nome'] = solucao['nome'].str.replace('Grupo 5','5 ANOS')
                solucao['nome'] = solucao['nome'].str.replace('série','SÉRIE')
                solucao['nome'] = solucao['nome'].str.replace('ano','ANO')
                operacao['Nome'] = operacao['Nome'].str.replace('Grupo 1','1 ANO')
                operacao['Nome'] = operacao['Nome'].str.replace('Grupo 2','2 ANOS')
                operacao['Nome'] = operacao['Nome'].str.replace('Grupo 3','3 ANOS')
                operacao['Nome'] = operacao['Nome'].str.replace('Grupo 4','4 ANOS')
                operacao['Nome'] = operacao['Nome'].str.replace('Grupo 5','5 ANOS')
                operacao['Nome'] = operacao['Nome'].str.replace('série','SÉRIE')
                operacao['Nome'] = operacao['Nome'].str.replace('ano','ANO')
                operacao['Série'] = operacao['Série'].str.replace('°','º')
                operacao['Série'] = operacao['Série'].str.replace('Grupo 1','1 ANO')
                operacao['Série'] = operacao['Série'].str.replace('Grupo 2','2 ANOS')
                operacao['Série'] = operacao['Série'].str.replace('Grupo 3','3 ANOS')
                operacao['Série'] = operacao['Série'].str.replace('Grupo 4','4 ANOS')
                operacao['Série'] = operacao['Série'].str.replace('Grupo 5','5 ANOS')
                operacao['Série'] = operacao['Série'].str.replace('ano','ANO')
                operacao['Série'] = operacao['Série'].str.replace('série','SÉRIE')
                operacao['Segmento'] = operacao['Segmento'].str.replace('FUNDAMENTAL ANOS INICIAIS','FUNDAMENTAL I')
                operacao['Segmento'] = operacao['Segmento'].str.replace('FUNDAMENTAL ANOS FINAIS','FUNDAMENTAL II')
                operacao['Nome'] = operacao['Nome'].str.replace('Extensivo','EXTENSIVO')
                operacao['Nome'] = operacao['Nome'].str.replace('Semi','SEMI')
                operacao['Série'] = operacao['Série'].str.replace('Extensivo','PRE VESTIBULAR')
                operacao['Série'] = operacao['Série'].str.replace('Semi','SEMI EXTENSIVO II')
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('Extensivo','PRE VESTIBULAR')
                solucao['serie_produto'] = solucao['serie_produto'].str.replace('Semi','SEMI EXTENSIVO II')
                solucao['nome'] = solucao['nome'].str.replace('Extensivo','EXTENSIVO')
                solucao['nome'] = solucao['nome'].str.replace('Semi','SEMI')

                solucao = solucao.rename(columns={'utilizacao_produto':'utilizacao_produto2','periodo_produto':'periodo_produto2'})
                solucao = solucao.rename(columns={'utilizacao_produto2':'periodo_produto','periodo_produto2':'utilizacao_produto'})
                solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','periodo_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
                #solucao

                ope3bim = operacao.loc[operacao['Bimestre'] == '3º BIMESTRE']
                #ope3bim
                sol3bim = solucao.loc[solucao['periodo_produto'] == '3º BIMESTRE']
                #sol3bim
                brinde3bim = df_brinde_final.loc[df_brinde_final['nome_da_regra'].str.contains('3º BIMESTRE')]
                #brinde3bim


                
                #operacao
                st.divider()

                with st.spinner('Aguarde...'):
                    time.sleep(3)

                st.success('Concluído com sucesso!', icon="✅")
                def convert_df(df):
                    # IMPORTANT: Cache the conversion to prevent computation on every rerun
                    return df.to_csv(index=False).encode('UTF-8')
                
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        operacao.to_excel(writer, index=False, sheet_name='Sheet1')
                        # Configurar os parâmetros para o botão de download
                    st.download_button(
                            label="Download do cadastro (XLSX)",
                        data=output.getvalue(),
                        file_name=f'{today}-{escola}-operacao.xlsx',
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    #output = io.BytesIO()
                    #with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    #    ope3bim.to_excel(writer, index=False, sheet_name='Sheet1')
                    #    # Configurar os parâmetros para o botão de download
                    #st.download_button(
                    #        label="Download do cadastro 3º Bimestre (XLSX)",
                    #    data=output.getvalue(),
                    #    file_name=f'{today}-{escola}-3bim.xlsx',
                    #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    #)
                    
                with col2:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        solucao.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                    st.download_button(
                        label="Download Solução (XLSX)",
                        data=output.getvalue(),
                        file_name=f'{today}-{escola}-solucao.xlsx',
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    solucao = convert_df(solucao)
                    st.download_button(
                    label="Download Solução (CSV) ",
                        data=solucao,
                        file_name=f'{today}-{escola}-solucao_import.csv',
                        mime='text/csv'
                    )
                    #sol3bim = convert_df(sol3bim)
                    #st.download_button(
                    #label="Download Solução 3º Bimestre (CSV)",
                    #    data=sol3bim,
                    #    file_name=f'{today}-{escola}-solucao_import_3bim.csv',
                    #    mime='text/csv'
                    #)
                        
                with col3:
                    df_brinde_final = convert_df(df_brinde_final)
                    st.download_button(
                    label="Download do brinde (CSV)",
                        data=df_brinde_final,
                        file_name=f'{today}-{escola}-brinde_import.csv',
                        mime='text/csv'
                    )
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_brinde_final2.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                    st.download_button(
                        label="Download do brinde (XLSX)",
                        data=output.getvalue(),
                        file_name=f'{today}-{escola}-brinde.xlsx',
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    #df_brinde_h5 = convert_df(df_brinde_h5)
                    #st.download_button(
                    #label="Download do brinde H5 (CSV)",
                    #    data=df_brinde_h5,
                    #    file_name=f'{today}-{escola}-brinde_h5_import.csv',
                    #    mime='text/csv'
                    #)
                    #brinde3bim = convert_df(brinde3bim)
                    #st.download_button(
                    #label="Download Brinde 3º Bimestre (CSV)",
                    #    data=brinde3bim,
                    #    file_name=f'{today}-{escola}-brinde_import_3bimes.csv',
                    #    mime='text/csv'
                    #)
                

                ###### DEBUG COM FILTRO
                st.divider()
                st.write("Cliente:", escola)
                st.divider()
                st.write('Resultado:')
                filter = operacao[['Escola','Marca','Segmento','Série','Bimestre','Nome','Descrição Magento','Quantidade de alunos','Customer Group']]
                selected = st.selectbox('Selecione a série:', ['',*filter['Série'].unique()])
                if selected:
                    selected_serie = filter[filter['Série'] == selected]
                    selected_serie
                else:
                    filter
                ##################


##########################################################################################################################################################
##########################################################################################################################################################

    if choice == 'PREMIUM-UNIQUE':
        marca = 'AZ' ## ou AZ SESC B2B ou AZ/SESC
        sheetname = 'itens_performance'
        planilha = 'itens.xlsx'
        today = date.today().strftime('%d-%m-%Y')
        cliente_tipo = 'B2B'

        st.success("Simulador - SEB")
        cliente = st.text_input('Digite o CNPJ da escola:')
        file = st.file_uploader("Selecione um arquivo Excel", type=["xlsm"])

        if file is not None:
            simul = pd.read_excel(file, sheet_name='cálculos Anual')
            simul=simul.assign(Bimestre="ANUAL")
            simul.replace(0, np.nan, inplace = True)

            simul = simul[simul['Quantidade de alunos']>0]

            simul = simul.rename(columns={'Construindo a Alfabetização':'Alfabetização','Itinerários Formativos Micro cursos     (2 IF)':'Itinerários','H5 - (3 Horas) Language Book + CLIL e PBL ':'H5 - 3 Horas','H5 - (2 horas)\nInternational Journey + \nApp H5':'H5 - 2 horas Journey','H5 Plus\n (3 horas extras)':'H5 Plus','My Life\n(Base)':'My Life - Base','My Life\n(2024)':'My Life - 2024','Binoculo By Tell Me\n(Base)':'Binoculo - Base','Educacross Ed. Infantil\n(Base)':'Educacross Infantil - Base','Educacross\n(Base)':'Educacross - Base','Educacross AZ\n(Base)':'Educacross AZ - Base','Educacross H5\n(Base)':'Educacross H5 - Base','Ubbu\n(Base)':'Ubbu - Base','Binoculo By Tell Me\n(2024)':'Binoculo - 2024','Educacross Ed. Infantil\n(2024)':'Educacross Infantil - 2024','Educacross\n(2024)':'Educacross - 2024','Educacross AZ\n(2024)':'Educacross AZ - 2024','Educacross H5\n(2024)':'Educacross H5 - 2024','Ubbu\n(2024)':'Ubbu - 2024','Árvore\n(1 Módulo)':'Árvore 1 Módulo','Árvore\n(2 Módulos)':'Árvore 2 Módulos','Árvore\n(3 Módulos)':'Árvore 3 Módulos','total aluno/ano\nsem desconto':'total aluno sem desconto','total aluno/ano\ncom desconto sem complementar':'total aluno com desconto sem complementar','total aluno/ano\ncom desconto + Complementares':'total aluno com desconto com Complementares',})
                
            simul = simul[['Série','Segmento','Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Unique','% Desconto Volume','Quantidade de alunos','Razão Social','CNPJ','Squad','Tipo','Observação','Grupo de cliente','Bimestre','% Desconto Extra','% Desconto Total']]
            simulador = simul.copy()
            df_cliente = simulador.loc[simulador['CNPJ'].str.strip() == cliente]
            df_cliente = df_cliente.fillna(0)

            df_cliente['Plataforma AZ'] = df_cliente['Plataforma AZ'].where(df_cliente['Plataforma AZ'] == 0, 1)
            df_cliente['Materiais Impressos AZ'] = df_cliente['Materiais Impressos AZ'].where(df_cliente['Materiais Impressos AZ'] == 0, 1)
            df_cliente['Alfabetização'] = df_cliente['Alfabetização'].where(df_cliente['Alfabetização'] == 0, 1)
            df_cliente['Cantalelê'] = df_cliente['Cantalelê'].where(df_cliente['Cantalelê'] == 0, 1)
            df_cliente['Mundo Leitor'] = df_cliente['Mundo Leitor'].where(df_cliente['Mundo Leitor'] == 0, 1)
            df_cliente['4 Avaliações Nacionais'] = df_cliente['4 Avaliações Nacionais'].where(df_cliente['4 Avaliações Nacionais'] == 0, 1)
            df_cliente['1 Simulado ENEM'] = df_cliente['1 Simulado ENEM'].where(df_cliente['1 Simulado ENEM'] == 0, 1)
            df_cliente['5 Simulados ENEM'] = df_cliente['5 Simulados ENEM'].where(df_cliente['5 Simulados ENEM'] == 0, 1)
            df_cliente['1 Simulado Regional'] = df_cliente['1 Simulado Regional'].where(df_cliente['1 Simulado Regional'] == 0, 1)
            df_cliente['Itinerários'] = df_cliente['Itinerários'].where(df_cliente['Itinerários'] == 0, 1)
            df_cliente['H5 - 3 Horas'] = df_cliente['H5 - 3 Horas'].where(df_cliente['H5 - 3 Horas'] == 0, 1)
            df_cliente['H5 - 2 horas Journey'] = df_cliente['H5 - 2 horas Journey'].where(df_cliente['H5 - 2 horas Journey'] == 0, 1)
            df_cliente['H5 Plus'] = df_cliente['H5 Plus'].where(df_cliente['H5 Plus'] == 0, 1)
            df_cliente['My Life - Base'] = df_cliente['My Life - Base'].where(df_cliente['My Life - Base'] == 0, 1)
            df_cliente['My Life - 2024'] = df_cliente['My Life - 2024'].where(df_cliente['My Life - 2024'] == 0, 1)
            df_cliente['Binoculo - Base'] = df_cliente['Binoculo - Base'].where(df_cliente['Binoculo - Base'] == 0, 1)
            df_cliente['Educacross Infantil - Base'] = df_cliente['Educacross Infantil - Base'].where(df_cliente['Educacross Infantil - Base'] == 0, 1)
            df_cliente['Educacross - Base'] = df_cliente['Educacross - Base'].where(df_cliente['Educacross - Base'] == 0, 1)
            df_cliente['Educacross AZ - Base'] = df_cliente['Educacross AZ - Base'].where(df_cliente['Educacross AZ - Base'] == 0, 1)
            df_cliente['Educacross H5 - Base'] = df_cliente['Educacross H5 - Base'].where(df_cliente['Educacross H5 - Base'] == 0, 1)
            df_cliente['Ubbu - Base'] = df_cliente['Ubbu - Base'].where(df_cliente['Ubbu - Base'] == 0, 1)
            df_cliente['Binoculo - 2024'] = df_cliente['Binoculo - 2024'].where(df_cliente['Binoculo - 2024'] == 0, 1)
            df_cliente['Educacross Infantil - 2024'] = df_cliente['Educacross Infantil - 2024'].where(df_cliente['Educacross Infantil - 2024'] == 0, 1)
            df_cliente['Educacross - 2024'] = df_cliente['Educacross - 2024'].where(df_cliente['Educacross - 2024'] == 0, 1)
            df_cliente['Educacross AZ - 2024'] = df_cliente['Educacross AZ - 2024'].where(df_cliente['Educacross AZ - 2024'] == 0, 1)
            df_cliente['Educacross H5 - 2024'] = df_cliente['Educacross H5 - 2024'].where(df_cliente['Educacross H5 - 2024'] == 0, 1)
            df_cliente['Ubbu - 2024'] = df_cliente['Ubbu - 2024'].where(df_cliente['Ubbu - 2024'] == 0, 1)
            df_cliente['Árvore 1 Módulo'] = df_cliente['Árvore 1 Módulo'].where(df_cliente['Árvore 1 Módulo'] == 0, 1)
            df_cliente['Árvore 2 Módulos'] = df_cliente['Árvore 2 Módulos'].where(df_cliente['Árvore 2 Módulos'] == 0, 1)
            df_cliente['Árvore 3 Módulos'] = df_cliente['Árvore 3 Módulos'].where(df_cliente['Árvore 3 Módulos'] == 0, 1)
            df_cliente['School Guardian'] = df_cliente['School Guardian'].where(df_cliente['School Guardian'] == 0, 1)
            df_cliente['Tindin'] = df_cliente['Tindin'].where(df_cliente['Tindin'] == 0, 1)
            df_cliente['Scholastic Earlybird and Bookflix'] = df_cliente['Scholastic Earlybird and Bookflix'].where(df_cliente['Scholastic Earlybird and Bookflix'] == 0, 1)
            df_cliente['Scholastic Literacy Pro'] = df_cliente['Scholastic Literacy Pro'].where(df_cliente['Scholastic Literacy Pro'] == 0, 1)
            df_cliente['Livro de Inglês'] = df_cliente['Unique'].where(df_cliente['Unique'] == 0,1)

            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ed. Infantil','INFANTIL')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Iniciais','FUNDAMENTAL ANOS INICIAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Finais','FUNDAMENTAL ANOS FINAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ensino Médio','ENSINO MÉDIO')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('PV','PRÉ VESTIBULAR')
            df_cliente=df_cliente.assign(Extra="")
            
            ###regra do AZ e Plataforma
            df_cliente.loc[(df_cliente['Plataforma AZ'] == 1) & (df_cliente['Materiais Impressos AZ'] == 1), ['Plataforma AZ']] = 0
            
            ####regra do h5
            df_cliente.loc[(df_cliente['H5 Plus'] == 1) & (df_cliente['H5 - 2 horas Journey'] == 1), ['H5 - 2 horas Journey','H5 - 3 Horas']] = 0
            df_cliente.loc[(df_cliente['H5 - 2 horas Journey'] == 1) & (df_cliente['H5 - 3 Horas'] == 1), ['H5 - 3 Horas']] = 0
            
            df_client = df_cliente.copy()
            lista = ['Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês']
                
            #df_client.to_excel('cliente.xlsx')
            for item in lista:
                df_client.loc[df_client[item] == 1.0, item] = item
            COLUNAS = ['Série', 'Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra']
            p = pd.DataFrame(columns=COLUNAS)
            
            for i in lista:
                data = df_client[df_client[i] == i].groupby(['Série', 'Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Squad','Tipo','Bimestre',i])['Extra'].count().reset_index()
                data = data.rename(columns={i: 'Produto'})
                p = pd.concat([p,data])
            p = p.sort_values(by=['Série'])
            p = p.reset_index()
            p = p.drop(columns=['index'])
            p = p.drop_duplicates()
            
            itens = pd.read_excel(planilha, sheet_name=sheetname)
            itens = itens[['MARCA',2024,'2024+','Produto','DESCRIÇÃO MAGENTO (B2C e B2B)','BIMESTRE','SEGMENTO','SÉRIE','PÚBLICO','TIPO DE FATURAMENTO']]
            itens = itens.rename(columns={'MARCA':'Marca','DESCRIÇÃO MAGENTO (B2C e B2B)':'Descrição Magento','BIMESTRE':'Bimestre','SEGMENTO':'Segmento','SÉRIE':'Série','PÚBLICO':'Público','TIPO DE FATURAMENTO':'Faturamento'})
            itens = itens[(itens['Marca'] == marca) | (itens['Marca'] == 'CONEXIA') | (itens['Marca'] == 'MUNDO LEITOR') | (itens['Marca'] == 'MY LIFE')| (itens['Marca'] == 'HIGH FIVE')]
            
            pdt = pd.merge(p, itens, on=['Série','Bimestre','Segmento','Produto'], how='inner')
                
            cod_serial = pd.read_excel(planilha, sheet_name='cod_serial')
            
            pdt = pd.merge(pdt, cod_serial, on=['Série','Bimestre','Segmento','Público'], how='inner')
            pdt['Ano'] = '2024'
            pdt['SKU'] = pdt['Ano'] + pdt['Serial']
            pdt = pdt[['Série','Segmento','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra','Produto','Marca',2024,'2024+','Descrição Magento','Público','Faturamento','Serial','Categoria','Ano','SKU']]
            
            h = re.compile(r'[../\-]')
            pdt['CNPJ_off'] = [h.sub('', x) for x in pdt['CNPJ']]
            pdt['CNPJ_off'] = [x.lstrip('0') for x in pdt['CNPJ_off']]
            pdt['CNPJ_off'] = pdt['CNPJ_off'].astype(float)
            
            cod_nome = pd.read_excel(planilha, sheet_name='nome')
            cod_nome['CNPJ_off'] = cod_nome['CNPJ_off'].astype(float)
            pdt = pd.merge(pdt, cod_nome, on=['CNPJ_off'], how='inner')
            ####################################################################################################
            ######NOVAS REGRAS POR SÉRIE#####################################################
            
            serie = pdt['Série'].unique()
            pdt_final = []
            pdt = pdt[~((pdt['Marca'] == 'CONEXIA') & (pdt['Bimestre'].str.contains('BIMESTRE')))]
            for i in serie:
                pdt_serie = pdt.loc[pdt['Série'] == i]
                pdt_serie['Marca'] = pdt_serie['Marca'].str.replace('MUNDO LEITOR','AZ')                     
                
                if (pdt_serie['Marca'].str.contains('MY LIFE').any()):
                        pdt_serie = pdt_serie[~((pdt_serie['Marca'] == 'MY LIFE') & (pdt_serie['Bimestre'].str.contains('BIMESTRE')))]
                
                pdt_final.append(pdt_serie)
                pdt_full = pd.concat(pdt_final)
            #pdt_full = pdt_full[~((pdt_full['Marca'] == 'AZ') & (pdt_full['Bimestre'].str.contains('ANUAL')))]
            pdt = pdt_full.copy()
            pdt.loc[pdt['Marca'] == 'MY LIFE', ['Marca']] = 'CONEXIA'
            pdt.loc[pdt['Marca'] == 'HIGH FIVE', ['Marca']] = 'CONEXIA'
            ######End Regra   
        
            pdt['Nome'] = 'SOLUÇÃO ' + pdt['Marca']  + ' - ' + pdt['Escola'] + ' - ' + pdt['Segmento'] + ' - ' + pdt['Série'] + ' - ' + pdt['Bimestre']
            pdt['SKU'] = pdt['Escola'] + '2024' + pdt['Marca'] + pdt['Serial']
            pdt['SKU'] = pdt['SKU'].str.replace(' ','')
            pdt = pdt.drop_duplicates()
            
            operacoes = pdt[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome',2024,'2024+','Descrição Magento','Quantidade de alunos','% Desconto Volume','% Desconto Extra','% Desconto Total','Customer Group','Squad']]
            operacoes = operacoes.rename(columns = {2024:'Cód Itens'} )
            solucao = operacoes.copy()
            operacao = operacoes.copy()
            operacao = operacao[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome','Cód Itens','Descrição Magento','Quantidade de alunos','% Desconto Volume','% Desconto Extra','% Desconto Total','Customer Group','Squad']]
            #operacao.to_excel('operacao.xlsx')
            operacao = operacao.sort_values(by=['Série','Bimestre'])
            
                
            solucao = solucao.groupby(['Escola','CNPJ','Série','Bimestre','Marca','Segmento','Ano','Público','Serial','SKU','Nome','Customer Group','Squad'])['2024+'].sum().reset_index()
            solucao['visibilidade'] = 'N'
            solucao['faturamento_produto'] = 'MATERIAL'
            solucao['cliente_produto'] = cliente_tipo
            solucao['ativar_restricao'] = 'S'
            #solucao.to_csv('teste_solução.csv')
            #solucao
            categoria = pd.read_excel(planilha, sheet_name='categoriab2b')
            solucao = pd.merge(solucao,categoria, on=['Série'], how='inner')
            #solucao
            solucao['Categorias'] = solucao['Marca'] + '/' + solucao['Categorias']
            solucao = solucao.sort_values(by=['Bimestre','Série'], ascending=True)
            solucao = solucao.rename(columns={'Público':'grupo_de_atributo','Marca':'marca_produto', 'Nome':'nome', 'SKU':'sku', 'Ano':'ano_produto', 'Série':'serie_produto', 'Bimestre':'utilizacao_produto', 'Categorias':'categorias', '2024+':'items', 'Customer Group':'grupos_permissao'})
            solucao['items'] = solucao['items'].apply(lambda x: x[:-1])
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            
            solucao['nome'] = solucao['nome'].str.replace('INFANTIL','EI')
            solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
            solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
            solucao['nome'] = solucao['nome'].str.replace('ENSINO MÉDIO','EM')
            operacao['Nome'] = operacao['Nome'].str.replace('INFANTIL','EI')
            operacao['Nome'] = operacao['Nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
            operacao['Nome'] = operacao['Nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
            operacao['Nome'] = operacao['Nome'].str.replace('ENSINO MÉDIO','EM')
            
            
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('°','º')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 1','1 ANO')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 2','2 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 3','3 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 4','4 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 5','5 ANOS')
            
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('série','SÉRIE')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('ano','ANO')
            solucao['nome'] = solucao['nome'].str.replace('°','º')
            solucao['publico_produto'] = 'ALUNO'
            
            solucao.loc[(solucao['nome'].str.contains('BIMESTRE')) , ['periodo_produto']] = 'BIMESTRAL'
            solucao.loc[(solucao['nome'].str.contains('ANUAL')) , ['periodo_produto']] = 'ANUAL'
            solucao.loc[(solucao['nome'].str.contains('SEMESTRAL')) , ['periodo_produto']] = 'SEMESTRAL'
            solucao.loc[(solucao['serie_produto'].str.contains('Semi')) , ['periodo_produto']] = 'SEMESTRAL'
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','periodo_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            df_brinde = operacao[['CNPJ','SKU','Série','Bimestre','Descrição Magento','Cód Itens','Customer Group']]
            df_brinde_input = pd.read_excel(planilha, sheet_name='brinde')
            df_brinde = pd.merge(df_brinde,df_brinde_input, on=['Cód Itens'], how='inner')
            df_brinde_final = df_brinde.copy()
            df_brinde_final = df_brinde_final[['Série_x','Nome da Regra','Customer Group','SKU_x','SKU_y']]
            df_brinde_final['Status'] = 'ATIVO'
            df_brinde_infantil = df_brinde_final.loc[df_brinde_final['Série_x'].str.contains('Grupo')]
            df_brinde_infantil['Qtd Incremento'] = 11
            df_brinde_demais = df_brinde_final.loc[~df_brinde_final['Série_x'].str.contains('Grupo')]
            df_brinde_demais['Qtd Incremento'] = 20
            df_brinde_final = pd.concat([df_brinde_infantil,df_brinde_demais])
            df_brinde_final['Qtd Condicao'] = 1
            df_brinde_final = df_brinde_final.rename(columns={'Customer Group':'Grupo do Cliente','SKU_x':'Sku Condicao','SKU_y':'Sku Brinde'})
            df_brinde_final = df_brinde_final[['Nome da Regra','Status','Grupo do Cliente','Sku Condicao','Qtd Condicao','Sku Brinde','Qtd Incremento']]
            df_brinde_final = df_brinde_final.rename(columns= {'Nome da Regra':'nome_da_regra','Status':'status','Grupo do Cliente':'grupo_do_cliente','Sku Condicao':'sku_condicao','Qtd Condicao':'qtd_condicao','Sku Brinde':'sku_brinde','Qtd Incremento':'qtd_incremento'})
            df_brinde_final = df_brinde_final.sort_values(by=['grupo_do_cliente','nome_da_regra'])
            df_brinde_final['id'] = ''
            df_brinde_final = df_brinde_final[['id','nome_da_regra','status','grupo_do_cliente','sku_condicao','qtd_condicao','sku_brinde','qtd_incremento']]
            ######## Exibir na tela para conferência #####
            escola = operacao['Escola'].unique()[0]
            df_brinde_h5 = df_brinde_final.loc[df_brinde_final['nome_da_regra'].str.contains('H5')]
            #df_brinde_h5
            df_brinde_final2 = df_brinde_final.copy()

            #### subir nas demais orreções
            solucao['nome'] = solucao['nome'].str.replace('Grupo 1','1 ANO')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 2','2 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 3','3 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 4','4 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('Grupo 5','5 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('série','SÉRIE')
            solucao['nome'] = solucao['nome'].str.replace('ano','ANO')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 1','1 ANO')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 2','2 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 3','3 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 4','4 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('Grupo 5','5 ANOS')
            operacao['Nome'] = operacao['Nome'].str.replace('série','SÉRIE')
            operacao['Nome'] = operacao['Nome'].str.replace('ano','ANO')
            operacao['Série'] = operacao['Série'].str.replace('°','º')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 1','1 ANO')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 2','2 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 3','3 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 4','4 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('Grupo 5','5 ANOS')
            operacao['Série'] = operacao['Série'].str.replace('ano','ANO')
            operacao['Série'] = operacao['Série'].str.replace('série','SÉRIE')
            operacao['Segmento'] = operacao['Segmento'].str.replace('FUNDAMENTAL ANOS INICIAIS','FUNDAMENTAL I')
            operacao['Segmento'] = operacao['Segmento'].str.replace('FUNDAMENTAL ANOS FINAIS','FUNDAMENTAL II')
            operacao['Nome'] = operacao['Nome'].str.replace('Extensivo','EXTENSIVO')
            operacao['Nome'] = operacao['Nome'].str.replace('Semi','SEMI')
            operacao['Série'] = operacao['Série'].str.replace('Extensivo','PRE VESTIBULAR')
            operacao['Série'] = operacao['Série'].str.replace('Semi','SEMI EXTENSIVO II')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Extensivo','PRE VESTIBULAR')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Semi','SEMI EXTENSIVO II')
            solucao['nome'] = solucao['nome'].str.replace('Extensivo','EXTENSIVO')
            solucao['nome'] = solucao['nome'].str.replace('Semi','SEMI')

            solucao = solucao.rename(columns={'utilizacao_produto':'utilizacao_produto2','periodo_produto':'periodo_produto2'})
            solucao = solucao.rename(columns={'utilizacao_produto2':'periodo_produto','periodo_produto2':'utilizacao_produto'})
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','periodo_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]
            #solucao

            ope3bim = operacao.loc[operacao['Bimestre'] == '3º BIMESTRE']
            #ope3bim
            sol3bim = solucao.loc[solucao['periodo_produto'] == '3º BIMESTRE']
            #sol3bim
            brinde3bim = df_brinde_final.loc[df_brinde_final['nome_da_regra'].str.contains('3º BIMESTRE')]
            #brinde3bim

            #operacao
            st.divider()
            with st.spinner('Aguarde...'):
                time.sleep(3)
            st.success('Concluído com sucesso!', icon="✅")
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv(index=False).encode('UTF-8')
            
            
            col1, col2, col3 = st.columns(3)
            with col1:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    operacao.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                st.download_button(
                        label="Download do cadastro (XLSX)",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-operacao.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                #output = io.BytesIO()
                #with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                #    ope3bim.to_excel(writer, index=False, sheet_name='Sheet1')
                #    # Configurar os parâmetros para o botão de download
                #st.download_button(
                #        label="Download do cadastro 3º Bimestre (XLSX)",
                #    data=output.getvalue(),
                #    file_name=f'{today}-{escola}-3bim.xlsx',
                #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                #)
                
            with col2:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    solucao.to_excel(writer, index=False, sheet_name='Sheet1')
                # Configurar os parâmetros para o botão de download
                st.download_button(
                    label="Download Solução (XLSX)",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-solucao.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                solucao = convert_df(solucao)
                st.download_button(
                label="Download Solução (CSV) ",
                    data=solucao,
                    file_name=f'{today}-{escola}-solucao_import.csv',
                    mime='text/csv'
                )
                #sol3bim = convert_df(sol3bim)
                #st.download_button(
                #label="Download Solução 3º Bimestre (CSV)",
                #    data=sol3bim,
                #    file_name=f'{today}-{escola}-solucao_import_3bim.csv',
                #    mime='text/csv'
                #)
                    
            with col3:
                df_brinde_final = convert_df(df_brinde_final)
                st.download_button(
                label="Download do brinde (CSV)",
                    data=df_brinde_final,
                    file_name=f'{today}-{escola}-brinde_import.csv',
                    mime='text/csv'
                )
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_brinde_final2.to_excel(writer, index=False, sheet_name='Sheet1')
                # Configurar os parâmetros para o botão de download
                st.download_button(
                    label="Download do brinde (XLSX)",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-brinde.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                #df_brinde_h5 = convert_df(df_brinde_h5)
                #st.download_button(
                #label="Download do brinde H5 (CSV)",
                #    data=df_brinde_h5,
                #    file_name=f'{today}-{escola}-brinde_h5_import.csv',
                #    mime='text/csv'
                #)
                #brinde3bim = convert_df(brinde3bim)
                #st.download_button(
                #label="Download Brinde 3º Bimestre (CSV)",
                #    data=brinde3bim,
                #    file_name=f'{today}-{escola}-brinde_import_3bimes.csv',
                #    mime='text/csv'
                #)
            
            ###### DEBUG COM FILTRO
            st.divider()
            st.write("Cliente:", escola)
            st.divider()
            st.write('Resultado:')
            filter = operacao[['Escola','Marca','Segmento','Série','Bimestre','Nome','Descrição Magento','Quantidade de alunos','Customer Group']]
            selected = st.selectbox('Selecione a série:', ['',*filter['Série'].unique()])
            if selected:
                selected_serie = filter[filter['Série'] == selected]
                selected_serie
            else:
                filter
                ##################


            
################°#########################################################################################################################################
################°##########################################################################################################################################

    if choice == 'CSV PARA EXCEL':
        csv_excel()

##########################################################################################################################################################
##########################################################################################################################################################

    if choice == 'EXCEL PARA CSV':
        excel_csv()

##########################################################################################################################################################
##########################################################################################################################################################

    if choice == 'PEDIDO PROGRAMADO':
        today = date.today().strftime('%d-%m-%Y')
        st.info("Pedido Programado")

        file = st.file_uploader("Selecione um arquivo Excel para gerar o pedido programado", type=["xlsx"])
        if file:
            df = pd.read_excel(file)

            df['SKU_SCHOOL'] = df['SKU']
            df['SKU'] =  df['SKU'] +','

            df['DESCRIÇÃO'] = df['DESCRIÇÃO MAGENTO']
            df.sort_values('DESCRIÇÃO MAGENTO', ascending=False)

            df['DESCRIÇÃO'] = df['DESCRIÇÃO'].str.replace(' - 1º BIMESTRE','')
            df['DESCRIÇÃO'] = df['DESCRIÇÃO'].str.replace(' - 2º BIMESTRE','')
            df['DESCRIÇÃO'] = df['DESCRIÇÃO'].str.replace(' - 3º BIMESTRE','')
            df['DESCRIÇÃO'] = df['DESCRIÇÃO'].str.replace(' - 4º BIMESTRE','')
            df['DESCRIÇÃO'] = df['DESCRIÇÃO'].str.replace(' - ANUAL','')

            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF1AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF2AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF3AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF4AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF5AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND11AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND11AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND11AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND11AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND12AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND12AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND12AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND12AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND13AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND13AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND13AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND13AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND14AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND14AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND14AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND14AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND15AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND15AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND15AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND15AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND26AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND26AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND26AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND26AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND27AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND27AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND27AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND27AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND28AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND28AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND28AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND28AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND29AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND29AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND29AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND29AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM1AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM1AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM1AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM1AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM2AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM2AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM2AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM2AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM3AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM3AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM3AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM3AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('SEMI4ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('SEMI4AL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('SEMI4AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('SEMI4AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('SEMI4AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF1AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF2AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF3AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF4AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF5AL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF1AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF2AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF3AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF4AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF5AL3B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF1AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF2AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF3AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF4AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF5AL4B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF1ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF2ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF3ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF4ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('INF5ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND11ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND12ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND13ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND14ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND15ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND26ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND27ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND28ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('FUND29ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM1ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM2ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EM3ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EXTI4ALA','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EXTPAL1B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EXTPAL2B','')
            df['SKU_SCHOOL'] = df['SKU_SCHOOL'].str.replace('EXTPAL3B','')


            df = df.groupby(['SÉRIE','SKU_SCHOOL','DESCRIÇÃO'])['SKU'].sum().reset_index()
            df['SKU'] = df['SKU'].apply(lambda x: x[:-1])
            df['enabled'] = 1
            df = df.rename(columns={'SKU':'sku','DESCRIÇÃO':'title'})
            df = df[['title','enabled','sku']]
            df['title'] = df['title'] + ' - ANUAL'

            df

            with st.spinner('Aguarde...'):
                time.sleep(3)
            st.success('Concluído com sucesso!', icon="✅")

            st.divider()
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv(index=False).encode('UTF-8')


            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            # Configurar os parâmetros para o botão de download
            st.download_button(
                label="Pedido Programado (XLSX)",
                data=output.getvalue(),
                file_name=f'{today}-pedidoprogramado.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            df = convert_df(df)
            st.download_button(
            label="Download Pedido Programado (CSV) ",
                data=df,
                file_name=f'{today}-pedidoprogramado.csv',
                mime='text/csv'
            )

##########################################################################################################################################################
##########################################################################################################################################################