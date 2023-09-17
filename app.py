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


    #####################################################################################

 
    #configurações do streamlit

    st.set_page_config(page_title="Script de soluções",page_icon="⭐",layout="wide",initial_sidebar_state="expanded")

    ##################################################################################

    
    # sidebar
    st.sidebar.image('https://sso.lex.education/assets/images/new-lex-logo.png', width=100)
    st.sidebar.title('Script de solução - Simulador')


    page = ['CONEXIA B2B','CONEXIA B2C','SEB','PREMIUM/UNIQUE']
    choice = st.sidebar.selectbox('Selecione:',page)
        
    

    ##########B2B################


    if choice == 'CONEXIA B2B':
 
        marca = 'AZ' ## ou AZ SESC B2B ou AZ/SESC
        sheetname = 'itens'
        planilha = 'itens.xlsx'
        today = date.today().strftime('%d-%m-%Y')
        cliente_tipo = 'B2B'

        st.info("Simulador - CONEXIA B2B")

        #  29.271.264/0001-61
        cliente = st.text_input('Digite o CNPJ da escola:')
        # Carrega o arquivo
        file = st.file_uploader("Selecione um arquivo Excel", type=["xlsm"])
        
        if file is not None:
           # Lê o arquivo Excel
            #simul0 = pd.read_excel(file, sheet_name='cálculos Anual')
            #simul0=simul0.assign(Bimestre="ANUAL")
            simul1 = pd.read_excel(file, sheet_name='cálculos 1º Bim')
            simul1=simul1.assign(Bimestre="1º BIMESTRE")
            simul2 = pd.read_excel(file, sheet_name='cálculos 2º Bim')
            simul2=simul2.assign(Bimestre="2º BIMESTRE")
            simul3 = pd.read_excel(file, sheet_name='cálculos 3º Bim')
            simul3=simul3.assign(Bimestre="3º BIMESTRE")
            simul4 = pd.read_excel(file, sheet_name='cálculos 4º Bim')
            simul4=simul4.assign(Bimestre="4º BIMESTRE")
            
            #alterar regra conforme leitura das planilhas
            #simul = pd.concat([simul0,simul1,simul2,simul3,simul4])
            simul = pd.concat([simul1,simul2,simul3,simul4])

            simul = simul.rename(columns={'Construindo a Alfabetização':'Alfabetização','Itinerários Formativos Micro cursos     (2 IF)':'Itinerários','H5 - (3 Horas) Language Book + CLIL e PBL ':'H5 - 3 Horas','H5 - (2 horas)\nInternational Journey + \nApp H5':'H5 - 2 horas Journey','H5 Plus\n (3 horas extras)':'H5 Plus','My Life\n(Base)':'My Life - Base','My Life\n(2024)':'My Life - 2024','Binoculo By Tell Me\n(Base)':'Binoculo - Base','Educacross Ed. Infantil\n(Base)':'Educacross Infantil - Base','Educacross\n(Base)':'Educacross - Base','Educacross AZ\n(Base)':'Educacross AZ - Base','Educacross H5\n(Base)':'Educacross H5 - Base','Ubbu\n(Base)':'Ubbu - Base','Binoculo By Tell Me\n(2024)':'Binoculo - 2024','Educacross Ed. Infantil\n(2024)':'Educacross Infantil - 2024','Educacross\n(2024)':'Educacross - 2024','Educacross AZ\n(2024)':'Educacross AZ - 2024','Educacross H5\n(2024)':'Educacross H5 - 2024','Ubbu\n(2024)':'Ubbu - 2024','Árvore\n(1 Módulo)':'Árvore 1 Módulo','Árvore\n(2 Módulos)':'Árvore 2 Módulos','Árvore\n(3 Módulos)':'Árvore 3 Módulos','total aluno/ano\nsem desconto':'total aluno sem desconto','total aluno/ano\ncom desconto sem complementar':'total aluno com desconto sem complementar','total aluno/ano\ncom desconto + Complementares':'total aluno com desconto com Complementares',})

            simulador = simul.copy()
            df_cliente = simulador.loc[simulador['CNPJ'] == cliente]
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
            df_cliente['H5 Plus)'] = df_cliente['H5 Plus'].where(df_cliente['H5 Plus'] == 0, 1)
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
            df_cliente['Livro de Inglês'] = df_cliente['Livro de Inglês'].where(df_cliente['Livro de Inglês'] == 0, 1)

            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ed. Infantil','INFANTIL')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Iniciais','FUNDAMENTAL ANOS INICIAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Finais','FUNDAMENTAL ANOS FINAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ensino Médio','ENSINO MÉDIO')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('PV ','PRÉ-VESTIBULAR')
            df_cliente=df_cliente.assign(Extra="")
            df_client = df_cliente.copy()

            lista = ['Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês']

            for item in lista:
                df_client.loc[df_client[item] == 1.0, item] = item
            COLUNAS = ['Série', 'Segmento','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra']
            p = pd.DataFrame(columns=COLUNAS)
            for i in lista:
                data = df_client[df_client[i] == i].groupby(['Série', 'Segmento', 'total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Squad','Tipo','Bimestre',i])['Extra'].count().reset_index()
                data = data.rename(columns={i: 'Produto'})
                p = pd.concat([p,data])
            p = p.sort_values(by=['Série'])
            p = p.reset_index()
            p = p.drop(columns=['index'])
            itens = pd.read_excel(planilha, sheet_name=sheetname)

            itens = itens[['MARCA',2024,'2024+','Produto','DESCRIÇÃO MAGENTO (B2C e B2B)','Bimestre','SEGMENTO','SÉRIE','PÚBLICO','TIPO DE FATURAMENTO']]
            itens = itens.rename(columns={'MARCA':'Marca','DESCRIÇÃO MAGENTO (B2C e B2B)':'Descrição Magento','SEGMENTO':'Segmento','SÉRIE':'Série','PÚBLICO':'Público','TIPO DE FATURAMENTO':'Faturamento'})
            itens = itens[(itens['Marca'] == marca) | (itens['Marca'] == 'CONEXIA') | (itens['Marca'] == 'MUNDO LEITOR') | (itens['Marca'] == 'MY LIFE')| (itens['Marca'] == 'HIGH FIVE')]
            
            pdt = pd.merge(p, itens, on=['Série','Bimestre','Segmento','Produto'], how='inner')
            cod_serial = pd.read_excel(planilha, sheet_name='cod_serial')
            pdt = pd.merge(pdt, cod_serial, on=['Série','Segmento','Bimestre','Público'], how='inner')
            pdt['Ano'] = '2024'
            pdt['SKU'] = pdt['Ano'] + pdt['Serial']
            pdt = pdt[['Série','Segmento','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra','Produto','Marca',2024,'2024+','Descrição Magento','Público','Faturamento','Serial','Categoria','Ano','SKU']]

            h = re.compile(r'[../\-]')
            pdt['CNPJ_off'] = [h.sub('', x) for x in pdt['CNPJ']]
            pdt['CNPJ_off'] = [x.lstrip('0') for x in pdt['CNPJ_off']]
            pdt['CNPJ_off'] = pdt['CNPJ_off'].astype(float)
            #pdt.to_excel(f'{output}/pdt_da_escola2.xlsx')
            
            cod_nome = pd.read_excel(planilha, sheet_name='nome')
            cod_nome['CNPJ_off'] = cod_nome['CNPJ_off'].astype(float)
            pdt = pd.merge(pdt, cod_nome, on=['CNPJ_off'], how='inner')
            #regra
            pdt = pdt[~((pdt['Marca'] == 'CONEXIA') & (pdt['Bimestre'].str.contains('ANUAL')))]

            pdt['Nome'] = 'SOLUÇÃO ' + pdt['Marca']  + ' - ' + pdt['Escola'] + ' - ' + pdt['Segmento'] + ' - ' + pdt['Série'] + ' - ' + pdt['Bimestre']
            pdt['SKU'] = pdt['Escola'] + pdt['Marca'] + pdt['Serial']
            pdt['SKU'] = pdt['SKU'].str.replace(' ','')
            
            operacoes = pdt[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome',2024,'2024+','Descrição Magento','Quantidade de alunos','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Total Série','Customer Group','Squad']]
            operacoes = operacoes.rename(columns = {2024:'Cód Itens'} )
            solucao = operacoes.copy()
            operacao = operacoes.copy()
            operacao = operacao[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome','Cód Itens','Descrição Magento','Quantidade de alunos','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Total Série','Customer Group','Squad']]
            operacao.to_excel('operacao.xlsx')

            
            solucao = solucao.groupby(['Escola','CNPJ','Série','Bimestre','Marca','Segmento','Ano','Público','Serial','SKU','Nome','Customer Group','Squad'])['2024+'].sum().reset_index()
            solucao['visibilidade'] = 'N'
            solucao['faturamento_produto'] = 'MATERIAL'
            solucao['cliente_produto'] = cliente_tipo
            solucao['ativar_restricao'] = 'S'
            
            solucao.to_csv('teste_solução.csv')

            categoria = pd.read_excel(planilha, sheet_name='categoriab2b')
            solucao = pd.merge(solucao,categoria, on=['Série'], how='inner')

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
            solucao['nome'] = solucao['nome'].str.replace('°','º')
            solucao['publico_produto'] = 'ALUNO'
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]

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
            df_brinde_final = df_brinde_final.sort_values(by=['Grupo do Cliente','Nome da Regra'])

            with st.spinner('Aguarde...'):
                time.sleep(2)

            st.success('Concluído com sucesso!', icon="✅")
            
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv(index=False).encode('UTF-8')

            escola = pdt['Escola'][0]
          


            col1, col2, col3 = st.columns(3)
            with col1:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    operacao.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                st.download_button(
                    label="Download do cadastro",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-operacao.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col2:
                solucao = convert_df(solucao)
                st.download_button(
                label="Download da solução",
                    data=solucao,
                    file_name=f'{today}-{escola}-solucao.csv',
                    mime='text/csv'
                )

            with col3:
                df_brinde_final = convert_df(df_brinde_final)
                st.download_button(
                label="Download do brinde",
                    data=df_brinde_final,
                    file_name=f'{today}-{escola}-brinde.csv',
                    mime='text/csv'
                )
                



    if choice == "B2C":
        st.info('Módulo B2C em construção :)')



    if choice == "SEB":

        marca = 'AZ' ## ou AZ SESC B2B ou AZ/SESC
        sheetname = 'itens'
        planilha = 'itens.xlsx'
        today = date.today().strftime('%d-%m-%Y')
        cliente_tipo = 'B2B'

        st.info("Simulador - SEB")

        #  29.271.264/0001-61
        cliente = st.text_input('Digite o CNPJ da escola:')
        # Carrega o arquivo
        file = st.file_uploader("Selecione um arquivo Excel", type=["xlsm"])
        
        if file is not None:
           # Lê o arquivo Excel
            simul = pd.read_excel(file, sheet_name='cálculos Anual')
            simul=simul.assign(Bimestre="ANUAL")

            simul = simul.rename(columns={'Construindo a Alfabetização':'Alfabetização','Itinerários Formativos Micro cursos     (2 IF)':'Itinerários','H5 - (3 Horas) Language Book + CLIL e PBL ':'H5 - 3 Horas','H5 - (2 horas)\nInternational Journey + \nApp H5':'H5 - 2 horas Journey','H5 Plus\n (3 horas extras)':'H5 Plus','My Life\n(Base)':'My Life - Base','My Life\n(2024)':'My Life - 2024','Binoculo By Tell Me\n(Base)':'Binoculo - Base','Educacross Ed. Infantil\n(Base)':'Educacross Infantil - Base','Educacross\n(Base)':'Educacross - Base','Educacross AZ\n(Base)':'Educacross AZ - Base','Educacross H5\n(Base)':'Educacross H5 - Base','Ubbu\n(Base)':'Ubbu - Base','Binoculo By Tell Me\n(2024)':'Binoculo - 2024','Educacross Ed. Infantil\n(2024)':'Educacross Infantil - 2024','Educacross\n(2024)':'Educacross - 2024','Educacross AZ\n(2024)':'Educacross AZ - 2024','Educacross H5\n(2024)':'Educacross H5 - 2024','Ubbu\n(2024)':'Ubbu - 2024','Árvore\n(1 Módulo)':'Árvore 1 Módulo','Árvore\n(2 Módulos)':'Árvore 2 Módulos','Árvore\n(3 Módulos)':'Árvore 3 Módulos','total aluno/ano\nsem desconto':'total aluno sem desconto','total aluno/ano\ncom desconto sem complementar':'total aluno com desconto sem complementar','total aluno/ano\ncom desconto + Complementares':'total aluno com desconto com Complementares',})

            simulador = simul.copy()
            df_cliente = simulador.loc[simulador['CNPJ'] == cliente]
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
            df_cliente['H5 Plus)'] = df_cliente['H5 Plus'].where(df_cliente['H5 Plus'] == 0, 1)
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
            df_cliente['Livro de Inglês'] = df_cliente['Livro de Inglês'].where(df_cliente['Livro de Inglês'] == 0, 1)

            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ed. Infantil','INFANTIL')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Iniciais','FUNDAMENTAL ANOS INICIAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Finais','FUNDAMENTAL ANOS FINAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ensino Médio','ENSINO MÉDIO')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('PV ','PRÉ-VESTIBULAR')
            df_cliente=df_cliente.assign(Extra="")
            df_client = df_cliente.copy()

            lista = ['Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês']

            for item in lista:
                df_client.loc[df_client[item] == 1.0, item] = item
            COLUNAS = ['Série', 'Segmento','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra']
            p = pd.DataFrame(columns=COLUNAS)
            for i in lista:
                data = df_client[df_client[i] == i].groupby(['Série', 'Segmento', 'total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Squad','Tipo','Bimestre',i])['Extra'].count().reset_index()
                data = data.rename(columns={i: 'Produto'})
                p = pd.concat([p,data])
            p = p.sort_values(by=['Série'])
            p = p.reset_index()
            p = p.drop(columns=['index'])
            itens = pd.read_excel(planilha, sheet_name=sheetname)

            itens = itens[['MARCA',2024,'2024+','Produto','DESCRIÇÃO MAGENTO (B2C e B2B)','Bimestre','SEGMENTO','SÉRIE','PÚBLICO','TIPO DE FATURAMENTO']]
            itens = itens.rename(columns={'MARCA':'Marca','DESCRIÇÃO MAGENTO (B2C e B2B)':'Descrição Magento','SEGMENTO':'Segmento','SÉRIE':'Série','PÚBLICO':'Público','TIPO DE FATURAMENTO':'Faturamento'})
            itens = itens[(itens['Marca'] == marca) | (itens['Marca'] == 'CONEXIA') | (itens['Marca'] == 'MUNDO LEITOR') | (itens['Marca'] == 'MY LIFE')| (itens['Marca'] == 'HIGH FIVE')]
            
            pdt = pd.merge(p, itens, on=['Série','Bimestre','Segmento','Produto'], how='inner')
            cod_serial = pd.read_excel(planilha, sheet_name='cod_serial')
            pdt = pd.merge(pdt, cod_serial, on=['Série','Segmento','Bimestre','Público'], how='inner')
            pdt['Ano'] = '2024'
            pdt['SKU'] = pdt['Ano'] + pdt['Serial']
            pdt = pdt[['Série','Segmento','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra','Produto','Marca',2024,'2024+','Descrição Magento','Público','Faturamento','Serial','Categoria','Ano','SKU']]

            h = re.compile(r'[../\-]')
            pdt['CNPJ_off'] = [h.sub('', x) for x in pdt['CNPJ']]
            pdt['CNPJ_off'] = [x.lstrip('0') for x in pdt['CNPJ_off']]
            pdt['CNPJ_off'] = pdt['CNPJ_off'].astype(float)
            #pdt.to_excel(f'{output}/pdt_da_escola2.xlsx')
            
            cod_nome = pd.read_excel(planilha, sheet_name='nome')
            cod_nome['CNPJ_off'] = cod_nome['CNPJ_off'].astype(float)
            pdt = pd.merge(pdt, cod_nome, on=['CNPJ_off'], how='inner')
            #regra
            #pdt = pdt[~((pdt['Marca'] == 'CONEXIA') & (pdt['Bimestre'].str.contains('ANUAL')))]

            pdt['Nome'] = 'SOLUÇÃO ' + pdt['Marca']  + ' - ' + pdt['Escola'] + ' - ' + pdt['Segmento'] + ' - ' + pdt['Série'] + ' - ' + pdt['Bimestre']
            pdt['SKU'] = pdt['Escola'] + pdt['Marca'] + pdt['Serial']
            pdt['SKU'] = pdt['SKU'].str.replace(' ','')
            
            operacoes = pdt[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome',2024,'2024+','Descrição Magento','Quantidade de alunos','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Total Série','Customer Group','Squad']]
            operacoes = operacoes.rename(columns = {2024:'Cód Itens'} )
            solucao = operacoes.copy()
            operacao = operacoes.copy()
            operacao = operacao[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome','Cód Itens','Descrição Magento','Quantidade de alunos','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Total Série','Customer Group','Squad']]
            operacao.to_excel('operacao.xlsx')

            
            solucao = solucao.groupby(['Escola','CNPJ','Série','Bimestre','Marca','Segmento','Ano','Público','Serial','SKU','Nome','Customer Group','Squad'])['2024+'].sum().reset_index()
            solucao['visibilidade'] = 'N'
            solucao['faturamento_produto'] = 'MATERIAL'
            solucao['cliente_produto'] = cliente_tipo
            solucao['ativar_restricao'] = 'S'
            
            solucao.to_csv('teste_solução.csv')

            categoria = pd.read_excel(planilha, sheet_name='categoriab2b')
            solucao = pd.merge(solucao,categoria, on=['Série'], how='inner')

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
            solucao['nome'] = solucao['nome'].str.replace('°','º')
            solucao['publico_produto'] = 'ALUNO'
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]

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
            df_brinde_final = df_brinde_final.sort_values(by=['Grupo do Cliente','Nome da Regra'])

            with st.spinner('Aguarde...'):
                time.sleep(2)

            st.success('Concluído com sucesso!', icon="✅")
            
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv(index=False).encode('UTF-8')

            escola = pdt['Escola'][0]
          


            col1, col2, col3 = st.columns(3)
            with col1:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    operacao.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                st.download_button(
                    label="Download do cadastro",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-operacao.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col2:
                solucao = convert_df(solucao)
                st.download_button(
                label="Download da solução",
                    data=solucao,
                    file_name=f'{today}-{escola}-solucao.csv',
                    mime='text/csv'
                )

            with col3:
                df_brinde_final = convert_df(df_brinde_final)
                st.download_button(
                label="Download do brinde",
                    data=df_brinde_final,
                    file_name=f'{today}-{escola}-brinde.csv',
                    mime='text/csv'
                )

       



    if choice == "PREMIUM/UNIQUE":
        marca = 'PREMIUM/UNIQUE' ## ou AZ SESC B2B ou AZ/SESC
        sheetname = 'itens'
        planilha = 'itens.xlsx'
        today = date.today().strftime('%d-%m-%Y')
        cliente_tipo = 'B2B'

        st.info("Simulador - SEB")

        #  29.271.264/0001-61
        cliente = st.text_input('Digite o CNPJ da escola:')
        # Carrega o arquivo
        file = st.file_uploader("Selecione um arquivo Excel", type=["xlsm"])
        
        if file is not None:
           # Lê o arquivo Excel
            simul = pd.read_excel(file, sheet_name='cálculos Anual')
            simul=simul.assign(Bimestre="ANUAL")

            simul = simul.rename(columns={'Construindo a Alfabetização':'Alfabetização','Itinerários Formativos Micro cursos     (2 IF)':'Itinerários','H5 - (3 Horas) Language Book + CLIL e PBL ':'H5 - 3 Horas','H5 - (2 horas)\nInternational Journey + \nApp H5':'H5 - 2 horas Journey','H5 Plus\n (3 horas extras)':'H5 Plus','My Life\n(Base)':'My Life - Base','My Life\n(2024)':'My Life - 2024','Binoculo By Tell Me\n(Base)':'Binoculo - Base','Educacross Ed. Infantil\n(Base)':'Educacross Infantil - Base','Educacross\n(Base)':'Educacross - Base','Educacross AZ\n(Base)':'Educacross AZ - Base','Educacross H5\n(Base)':'Educacross H5 - Base','Ubbu\n(Base)':'Ubbu - Base','Binoculo By Tell Me\n(2024)':'Binoculo - 2024','Educacross Ed. Infantil\n(2024)':'Educacross Infantil - 2024','Educacross\n(2024)':'Educacross - 2024','Educacross AZ\n(2024)':'Educacross AZ - 2024','Educacross H5\n(2024)':'Educacross H5 - 2024','Ubbu\n(2024)':'Ubbu - 2024','Árvore\n(1 Módulo)':'Árvore 1 Módulo','Árvore\n(2 Módulos)':'Árvore 2 Módulos','Árvore\n(3 Módulos)':'Árvore 3 Módulos','total aluno/ano\nsem desconto':'total aluno sem desconto','total aluno/ano\ncom desconto sem complementar':'total aluno com desconto sem complementar','total aluno/ano\ncom desconto + Complementares':'total aluno com desconto com Complementares',})

            simulador = simul.copy()
            df_cliente = simulador.loc[simulador['CNPJ'] == cliente]
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
            df_cliente['H5 Plus)'] = df_cliente['H5 Plus'].where(df_cliente['H5 Plus'] == 0, 1)
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
            df_cliente['Livro de Inglês'] = df_cliente['Livro de Inglês'].where(df_cliente['Livro de Inglês'] == 0, 1)

            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ed. Infantil','INFANTIL')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Iniciais','FUNDAMENTAL ANOS INICIAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Fund. Anos Finais','FUNDAMENTAL ANOS FINAIS')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('Ensino Médio','ENSINO MÉDIO')
            df_cliente['Segmento'] = df_cliente['Segmento'].str.replace('PV ','PRÉ-VESTIBULAR')
            df_cliente=df_cliente.assign(Extra="")
            df_client = df_cliente.copy()

            lista = ['Plataforma AZ','Materiais Impressos AZ','Alfabetização','Cantalelê','Mundo Leitor','4 Avaliações Nacionais','1 Simulado ENEM','5 Simulados ENEM','1 Simulado Regional','Itinerários','H5 - 3 Horas','H5 - 2 horas Journey','H5 Plus','My Life - Base','My Life - 2024','Binoculo - Base','Educacross Infantil - Base','Educacross - Base','Educacross AZ - Base','Educacross H5 - Base','Ubbu - Base','Binoculo - 2024','Educacross Infantil - 2024','Educacross - 2024','Educacross AZ - 2024','Educacross H5 - 2024','Ubbu - 2024','Árvore 1 Módulo','Árvore 2 Módulos','Árvore 3 Módulos','School Guardian','Tindin','Scholastic Earlybird and Bookflix','Scholastic Literacy Pro','Livro de Inglês']

            for item in lista:
                df_client.loc[df_client[item] == 1.0, item] = item
            COLUNAS = ['Série', 'Segmento','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra']
            p = pd.DataFrame(columns=COLUNAS)
            for i in lista:
                data = df_client[df_client[i] == i].groupby(['Série', 'Segmento', 'total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Squad','Tipo','Bimestre',i])['Extra'].count().reset_index()
                data = data.rename(columns={i: 'Produto'})
                p = pd.concat([p,data])
            p = p.sort_values(by=['Série'])
            p = p.reset_index()
            p = p.drop(columns=['index'])
            itens = pd.read_excel(planilha, sheet_name=sheetname)

            itens = itens[['MARCA',2024,'2024+','Produto','DESCRIÇÃO MAGENTO (B2C e B2B)','Bimestre','SEGMENTO','SÉRIE','PÚBLICO','TIPO DE FATURAMENTO']]
            itens = itens.rename(columns={'MARCA':'Marca','DESCRIÇÃO MAGENTO (B2C e B2B)':'Descrição Magento','SEGMENTO':'Segmento','SÉRIE':'Série','PÚBLICO':'Público','TIPO DE FATURAMENTO':'Faturamento'})
            itens = itens[(itens['Marca'] == marca) | (itens['Marca'] == 'CONEXIA') | (itens['Marca'] == 'MUNDO LEITOR') | (itens['Marca'] == 'MY LIFE')| (itens['Marca'] == 'HIGH FIVE')]
            pdt = pd.merge(p, itens, on=['Série','Bimestre','Segmento','Produto'], how='inner')
            cod_serial = pd.read_excel(planilha, sheet_name='cod_serial')
            pdt = pd.merge(pdt, cod_serial, on=['Série','Segmento','Bimestre','Público'], how='inner')
            pdt['Ano'] = '2024'
            pdt['SKU'] = pdt['Ano'] + pdt['Serial']
            pdt = pdt[['Série','Segmento','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Quantidade de alunos','Total Série','Razão Social','CNPJ','Bimestre','Squad','Tipo','Extra','Produto','Marca',2024,'2024+','Descrição Magento','Público','Faturamento','Serial','Categoria','Ano','SKU']]

            h = re.compile(r'[../\-]')
            pdt['CNPJ_off'] = [h.sub('', x) for x in pdt['CNPJ']]
            pdt['CNPJ_off'] = [x.lstrip('0') for x in pdt['CNPJ_off']]
            pdt['CNPJ_off'] = pdt['CNPJ_off'].astype(float)
            #pdt.to_excel(f'{output}/pdt_da_escola2.xlsx')
            
            cod_nome = pd.read_excel(planilha, sheet_name='nome')
            cod_nome['CNPJ_off'] = cod_nome['CNPJ_off'].astype(float)
            pdt = pd.merge(pdt, cod_nome, on=['CNPJ_off'], how='inner')
            #regra
            #pdt = pdt[~((pdt['Marca'] == 'CONEXIA') & (pdt['Bimestre'].str.contains('ANUAL')))]

            pdt['Nome'] = 'SOLUÇÃO ' + pdt['Marca']  + ' - ' + pdt['Escola'] + ' - ' + pdt['Segmento'] + ' - ' + pdt['Série'] + ' - ' + pdt['Bimestre']
            pdt['SKU'] = pdt['Escola'] + pdt['Marca'] + pdt['Serial']
            pdt['SKU'] = pdt['SKU'].str.replace(' ','')
            
            operacoes = pdt[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome',2024,'2024+','Descrição Magento','Quantidade de alunos','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Total Série','Customer Group','Squad']]
            operacoes = operacoes.rename(columns = {2024:'Cód Itens'} )
            solucao = operacoes.copy()
            operacao = operacoes.copy()
            operacao = operacao[['Escola','CNPJ','Ano','Marca','Serial','Segmento','Série','Bimestre','Público','SKU','Nome','Cód Itens','Descrição Magento','Quantidade de alunos','total aluno sem desconto','total aluno com desconto sem complementar','total aluno com desconto com Complementares','% Desconto Volume','% Desconto Extra','% Desconto Total','Total Série','Customer Group','Squad']]
            operacao.to_excel('operacao.xlsx')

            
            solucao = solucao.groupby(['Escola','CNPJ','Série','Bimestre','Marca','Segmento','Ano','Público','Serial','SKU','Nome','Customer Group','Squad'])['2024+'].sum().reset_index()
            solucao['visibilidade'] = 'N'
            solucao['faturamento_produto'] = 'MATERIAL'
            solucao['cliente_produto'] = cliente_tipo
            solucao['ativar_restricao'] = 'S'
            
            solucao.to_csv('teste_solução.csv')

            categoria = pd.read_excel(planilha, sheet_name='categoriab2b')
            solucao = pd.merge(solucao,categoria, on=['Série'], how='inner')

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
            solucao['nome'] = solucao['nome'].str.replace('°','º')
            solucao['publico_produto'] = 'ALUNO'
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]

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
            df_brinde_final = df_brinde_final.sort_values(by=['Grupo do Cliente','Nome da Regra'])

            with st.spinner('Aguarde...'):
                time.sleep(2)

            st.success('Concluído com sucesso!', icon="✅")
            
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv(index=False).encode('UTF-8')

            escola = pdt['Escola'][0]
          


            col1, col2, col3 = st.columns(3)
            with col1:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    operacao.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Configurar os parâmetros para o botão de download
                st.download_button(
                    label="Download do cadastro",
                    data=output.getvalue(),
                    file_name=f'{today}-{escola}-operacao.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col2:
                solucao = convert_df(solucao)
                st.download_button(
                label="Download da solução",
                    data=solucao,
                    file_name=f'{today}-{escola}-solucao.csv',
                    mime='text/csv'
                )

            with col3:
                df_brinde_final = convert_df(df_brinde_final)
                st.download_button(
                label="Download do brinde",
                    data=df_brinde_final,
                    file_name=f'{today}-{escola}-brinde.csv',
                    mime='text/csv'
                )
