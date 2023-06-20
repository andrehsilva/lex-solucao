from typing import Text
import streamlit as st
import pandas as pd
import io
import datetime
from datetime import datetime as dt
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
    st.sidebar.title('Script de solução - Simulador')


    page = ['B2B','B2C']
    choice = st.sidebar.selectbox('Selecione:',page)


    with open('template_simulador.xlsx', "rb") as template_file:
        template_byte = template_file.read()

        st.sidebar.download_button(label="Download arquivo template",
                            data=template_byte,
                            file_name="template_simulador.xlsx",
                            mime='application/octet-stream')



    if choice == 'B2B':

        today = date.today().strftime('%d-%m-%Y')
        cliente_tipo = 'B2B'
        
        st.info("Simulador B2B")

      

        marca = st.radio("Selecione a marca:",('AZ', 'AZ e High Five', 'High Five', 'My Lyfe'), horizontal=True)

        if marca == 'AZ':
            st.write('Selecionado: AZ')
            sheetname = 'itens_b2b_az' 
        elif marca == 'AZ e High Five':
            st.write('Selecionado: High Five')
            sheetname = 'itens_b2b_az'
        elif marca == 'High Five':
            st.write('Selecionado: High Five')
            sheetname = 'itens_b2b_hf'
        elif marca == 'My Lyfe':
            st.write('Selecionado: My Life')
            sheetname = 'itens_b2b_my'
        
        # CNPJ da escola
        cliente = st.text_input('Digite o CNPJ da escola:')

        marcas = ['AZ', 'CONEXIA', 'HIGH FIVE','MY LIFE']

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

            df_client = df_cliente.copy()

            lista = ['Materiais Impressos','Plataforma AZ','H5 Bilingual Education - Language Book + CLIL  e PBL','International Journey + App H5','Aulas Ao Vivo - ZOOM','Módulo de Comunicação','Liga das Corujinhas Games Educativos',
            'Educacross Games Matemática','Educacross Games Lingua Portuguesa','Educacross High Five','Cantalelê','My Life','UBBU','High Five Plus','4 Avaliações Nacionais','1 Simulado ENEM',
            '5 Simulados Enem','1 Simulado Regional','Itinerários Formativos Micro cursos (2 IF)','Mundo Leitor','ELT Aluno','Alfabetização ','Learning']

            for item in lista:
                df_cliente.loc[df_cliente[item] == 1, item] = item

            COLUNAS = ['CUSTOMER GROUP - ESCOLA','SQUAD ESCOLA','CNPJ ESCOLA','SÉRIE','SEGMENTO','TOTAL ALUNOS 2023','TOTAL PREÇO TABELA 2023','DESCONTO POR VOLUME','CUPOM EXTRA DE DESCONTO','PREÇO ACORDADO ESCOLA 2023 (ANUAL)','% DESCONTO POR SÉRIE 2023','TOTAL SEM DESCONTO','TOTAL COM DESCONTO SEM ELT','PREÇO  ANUAL FINAL SEM ELT','PREÇO FINAL COM  ELT (SE APLICÁVEL)','Produto','Extra']
            p = pd.DataFrame(columns=COLUNAS)

            for i in lista:
                data = df_cliente[df_cliente[i] == i].groupby(['CUSTOMER GROUP - ESCOLA','SQUAD','ESCOLA','CNPJ ESCOLA','SÉRIE','SEGMENTO','TOTAL ALUNOS 2023','TOTAL PREÇO TABELA 2023','DESCONTO POR VOLUME',
                                    'CUPOM EXTRA DE DESCONTO','PREÇO ACORDADO ESCOLA 2023 (ANUAL)','% DESCONTO POR SÉRIE 2023','TOTAL SEM DESCONTO','TOTAL COM DESCONTO SEM ELT','PREÇO  ANUAL FINAL SEM ELT','PREÇO FINAL COM  ELT (SE APLICÁVEL)',i])['Extra'].count().reset_index()
                data = data.rename(columns={i: 'Produto'})
                p = pd.concat([p,data])
         

            p['ANO'] = ''
            p['ANUAL'] = 'ANUAL'
            p['1º BIMESTRE'] = '1º BIMESTRE'
            p['2º BIMESTRE'] = '2º BIMESTRE'
            p['3º BIMESTRE'] = '3º BIMESTRE'
            p['4º BIMESTRE'] = '4º BIMESTRE'
            etapa = ['1º BIMESTRE','2º BIMESTRE','3º BIMESTRE','4º BIMESTRE','ANUAL']

            pb_aux = ['CUSTOMER GROUP - ESCOLA','SQUAD','ESCOLA','CNPJ ESCOLA','SÉRIE','SEGMENTO','TOTAL ALUNOS 2023','TOTAL PREÇO TABELA 2023','DESCONTO POR VOLUME','CUPOM EXTRA DE DESCONTO','% DESCONTO POR SÉRIE 2023','PREÇO ACORDADO ESCOLA 2023 (ANUAL)','TOTAL SEM DESCONTO','TOTAL COM DESCONTO SEM ELT','PREÇO  ANUAL FINAL SEM ELT','PREÇO FINAL COM  ELT (SE APLICÁVEL)','Produto','BIMESTRE']
            pb_t = pd.DataFrame(columns=pb_aux)

            for i in etapa:
                pb = p[p[i] == i].groupby(['CUSTOMER GROUP - ESCOLA','SQUAD','ESCOLA','CNPJ ESCOLA','SÉRIE','SEGMENTO','TOTAL ALUNOS 2023','TOTAL PREÇO TABELA 2023','DESCONTO POR VOLUME','CUPOM EXTRA DE DESCONTO','% DESCONTO POR SÉRIE 2023','PREÇO ACORDADO ESCOLA 2023 (ANUAL)','TOTAL SEM DESCONTO','TOTAL COM DESCONTO SEM ELT','PREÇO  ANUAL FINAL SEM ELT','PREÇO FINAL COM  ELT (SE APLICÁVEL)','Produto',i])['ANO'].count().reset_index()
                pb = pb.rename(columns={i: 'BIMESTRE'})
                pb_t = pd.concat([pb_t,pb])
            p = pb_t.copy()   

            #p['SÉRIE'].unique()

            p['ESCOLA'] = p['ESCOLA'].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
            p['ESCOLA'] = p['ESCOLA'].str.replace('EIRELI','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('COLEGIO','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('ESCOLA','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('CRECHE','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('INSTITUTO','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('EDUCANDARIO','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('COMUNIDADE','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('SOCIEDADE','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('CENTRO','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('EDUCACIONAL','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('EDUCACAO','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('ASSOCIACAO','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('INFANTIL','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('ENSINO','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('FUNDAMENTAL','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('MEDIO','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('LTDA','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('-','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('/','')
            p['ESCOLA'] = p['ESCOLA'].str.replace(' ','')
            p['ESCOLA'] = p['ESCOLA'].str.replace('.','')
            p['ANO'] = '2023'

            p = p.reset_index()
            p = p.drop(columns=['index'])
            p = p.rename(columns={'Produto':'PRODUTO'})
            p['DESCONTO POR VOLUME'] = p['DESCONTO POR VOLUME'].astype('float')
            p['CUPOM EXTRA DE DESCONTO'] = p['CUPOM EXTRA DE DESCONTO'].astype('float')
            #p.PRODUTO.unique()

            itens = pd.read_excel('itens.xlsx', sheet_name=sheetname)
            itens = itens.rename(columns={'UTILIZAÇÃO':'BIMESTRE'})
            itens = itens[itens['MARCA'].isin(marcas)]
            itens['PREÇO CATALOGO 2023'] = itens['PREÇO CATALOGO 2023'].astype('float')
        
            pdt = pd.merge(p, itens, on=['SÉRIE','SEGMENTO','BIMESTRE', 'PRODUTO'], how='inner')


            pdt['TOTAL DESCONTO'] = pdt['DESCONTO POR VOLUME'] + pdt['CUPOM EXTRA DE DESCONTO']
            pdt['VALOR COM DESCONTO'] = pdt.apply(lambda x: '{:.2f}'.format(x['PREÇO CATALOGO 2023'] - (x['PREÇO CATALOGO 2023'] * x['TOTAL DESCONTO'])), axis = 1)

            pdt = pdt.rename(columns={'2023+':'CODIGO ITENS', 'ANO':'ANO CATALOGO', 'BIMESTRE':'UTILIZAÇÃO','DESCRIÇÃO MAGENTO (B2C e B2B)':'DESCRIÇÃO ITEM (B2C e B2B)'})
            pdt = pdt[['ESCOLA','CNPJ ESCOLA','SÉRIE','SEGMENTO','UTILIZAÇÃO','ANO CATALOGO','MARCA','PRODUTO','DESCRIÇÃO ITEM (B2C e B2B)','CODIGO ITENS','TIPO','PÚBLICO','TOTAL ALUNOS 2023','TOTAL PREÇO TABELA 2023','PREÇO CATALOGO 2023','DESCONTO POR VOLUME','CUPOM EXTRA DE DESCONTO','TOTAL DESCONTO','TOTAL SEM DESCONTO','TOTAL COM DESCONTO SEM ELT','PREÇO  ANUAL FINAL SEM ELT','PREÇO FINAL COM  ELT (SE APLICÁVEL)','CUSTOMER GROUP - ESCOLA','SQUAD']]

            cod_serial = pd.read_excel('itens.xlsx', sheet_name='cod_serial')

            pdt = pd.merge(pdt, cod_serial, on=['SÉRIE','SEGMENTO','UTILIZAÇÃO','PÚBLICO'], how='inner')
            pdt['ANO'] = pdt['ANO CATALOGO'].astype(str) 
            pdt['SKU'] = pdt['ESCOLA'] + pdt['ANO'] + pdt['SERIAL']

            pdt = pdt[['ESCOLA','CNPJ ESCOLA','SÉRIE','SEGMENTO','UTILIZAÇÃO','SERIAL','ANO CATALOGO','MARCA','SKU','PRODUTO','DESCRIÇÃO ITEM (B2C e B2B)','CODIGO ITENS','TIPO','PÚBLICO','TOTAL ALUNOS 2023','TOTAL PREÇO TABELA 2023','PREÇO CATALOGO 2023','DESCONTO POR VOLUME','CUPOM EXTRA DE DESCONTO','TOTAL DESCONTO','TOTAL SEM DESCONTO','TOTAL COM DESCONTO SEM ELT','PREÇO  ANUAL FINAL SEM ELT','PREÇO FINAL COM  ELT (SE APLICÁVEL)','CUSTOMER GROUP - ESCOLA','SQUAD']]

            nome_ab = pd.read_excel('itens.xlsx', sheet_name='nome')
            nome_ab['CNPJ ESCOLA'] = nome_ab['CNPJ ESCOLA'].astype(float)

            pdt_nome = pdt.copy()
            #Regex ajuste de nome
            h = re.compile(r'[../\-]')
            pdt_nome['CNPJ ESCOLA'] = [h.sub('', x) for x in pdt_nome['CNPJ ESCOLA']]
            pdt_nome['CNPJ ESCOLA'] = [x.lstrip('0') for x in pdt_nome['CNPJ ESCOLA']]
            pdt_nome['CNPJ ESCOLA'] = pdt_nome['CNPJ ESCOLA'].astype(float)
            merge_nome = pd.merge(pdt_nome, nome_ab, on=['CNPJ ESCOLA'], how='inner')

            pdt['slice'] = pdt['MARCA'].str.slice(stop=2)
            pdt['SKU'] = merge_nome['nome_escola'] + '2023' + pdt['slice'] + pdt['SERIAL']
            pdt.drop(columns=['slice'],inplace=True)

            pdt_1 = pdt.loc[pdt['UTILIZAÇÃO'] == '1º BIMESTRE']
            pdt_2 = pdt.loc[pdt['UTILIZAÇÃO'] == '2º BIMESTRE']
            pdt_3 = pdt.loc[pdt['UTILIZAÇÃO'] == '3º BIMESTRE']
            pdt_4 = pdt.loc[pdt['UTILIZAÇÃO'] == '4º BIMESTRE']
            pdt_5 = pdt.loc[pdt['UTILIZAÇÃO'] == 'ANUAL']

            solucao_1_bimestre = pdt_1.groupby(['ESCOLA','CNPJ ESCOLA','SÉRIE','UTILIZAÇÃO','MARCA','SEGMENTO','ANO CATALOGO','PÚBLICO','SERIAL','SKU','CUSTOMER GROUP - ESCOLA','SQUAD'])['CODIGO ITENS'].sum().reset_index()
            solucao_2_bimestre = pdt_2.groupby(['ESCOLA','CNPJ ESCOLA','SÉRIE','UTILIZAÇÃO','MARCA','SEGMENTO','ANO CATALOGO','PÚBLICO','SERIAL','SKU','CUSTOMER GROUP - ESCOLA','SQUAD'])['CODIGO ITENS'].sum().reset_index()
            solucao_3_bimestre = pdt_3.groupby(['ESCOLA','CNPJ ESCOLA','SÉRIE','UTILIZAÇÃO','MARCA','SEGMENTO','ANO CATALOGO','PÚBLICO','SERIAL','SKU','CUSTOMER GROUP - ESCOLA','SQUAD'])['CODIGO ITENS'].sum().reset_index()
            solucao_4_bimestre = pdt_4.groupby(['ESCOLA','CNPJ ESCOLA','SÉRIE','UTILIZAÇÃO','MARCA','SEGMENTO','ANO CATALOGO','PÚBLICO','SERIAL','SKU','CUSTOMER GROUP - ESCOLA','SQUAD'])['CODIGO ITENS'].sum().reset_index()
            solucao_5_bimestre = pdt_5.groupby(['ESCOLA','CNPJ ESCOLA','SÉRIE','UTILIZAÇÃO','MARCA','SEGMENTO','ANO CATALOGO','PÚBLICO','SERIAL','SKU','CUSTOMER GROUP - ESCOLA','SQUAD'])['CODIGO ITENS'].sum().reset_index()

            if len(solucao_1_bimestre):
                solucao_1_bimestre['nome'] = 'SOLUÇÃO ' + solucao_1_bimestre['MARCA'] + ' - ' + merge_nome['nome_escola'] + ' - ' + solucao_1_bimestre['SEGMENTO'] + ' - ' + solucao_1_bimestre['SÉRIE'] + ' - ' + solucao_1_bimestre['UTILIZAÇÃO']
                #solucao_1_bimestre['PÚBLICO'] + ' - ' + '1º BIMESTRE'
                solucao_1_bimestre['visibilidade'] = 'N'
                solucao_1_bimestre['faturamento_produto'] = 'MATERIAL'
                solucao_1_bimestre['cliente_produto'] = cliente_tipo
                solucao_1_bimestre['ativar_restricao'] = 'S'                    


            if len(solucao_2_bimestre):
                solucao_2_bimestre['nome'] = 'SOLUÇÃO ' + solucao_2_bimestre['MARCA'] + ' - ' + merge_nome['nome_escola'] + ' - ' + solucao_2_bimestre['SEGMENTO'] + ' - ' + solucao_2_bimestre['SÉRIE'] + ' - ' + solucao_2_bimestre['UTILIZAÇÃO']
                #solucao_2_bimestre['PÚBLICO'] + ' - ' + '2º BIMESTRE'
                solucao_2_bimestre['visibilidade'] = 'N'
                solucao_2_bimestre['faturamento_produto'] = 'MATERIAL'
                solucao_2_bimestre['cliente_produto'] = cliente_tipo
                solucao_2_bimestre['ativar_restricao'] = 'S'


            if len(solucao_3_bimestre):
                solucao_3_bimestre['nome'] = 'SOLUÇÃO ' + solucao_3_bimestre['MARCA'] + ' - ' + merge_nome['nome_escola'] + ' - ' + solucao_3_bimestre['SEGMENTO'] + ' - '+ solucao_3_bimestre['SÉRIE'] + ' - ' + solucao_3_bimestre['UTILIZAÇÃO']
                #solucao_3_bimestre['PÚBLICO'] + ' - ' + '3º BIMESTRE'
                solucao_3_bimestre['visibilidade'] = 'N'
                solucao_3_bimestre['faturamento_produto'] = 'MATERIAL'
                solucao_3_bimestre['cliente_produto'] = cliente_tipo
                solucao_3_bimestre['ativar_restricao'] = 'S'


            if len(solucao_4_bimestre):
                solucao_4_bimestre['nome'] = 'SOLUÇÃO ' + solucao_4_bimestre['MARCA'] + ' - ' + merge_nome['nome_escola'] + ' - ' + solucao_4_bimestre['SEGMENTO'] + ' - ' + solucao_4_bimestre['SÉRIE'] + ' - ' + solucao_4_bimestre['UTILIZAÇÃO']
                solucao_4_bimestre['visibilidade'] = 'N'
                solucao_4_bimestre['faturamento_produto'] = 'MATERIAL'
                solucao_4_bimestre['cliente_produto'] = cliente_tipo
                solucao_4_bimestre['ativar_restricao'] = 'S'


            if len(solucao_5_bimestre):
                solucao_5_bimestre['nome'] = 'SOLUÇÃO ' + solucao_5_bimestre['MARCA'] + ' - ' + merge_nome['nome_escola'] + ' - ' + solucao_5_bimestre['SEGMENTO'] + ' - ' + solucao_5_bimestre['SÉRIE'] + ' - ' + solucao_5_bimestre['UTILIZAÇÃO']
                solucao_5_bimestre['visibilidade'] = 'N'
                solucao_5_bimestre['faturamento_produto'] = 'MATERIAL'
                solucao_5_bimestre['cliente_produto'] = cliente_tipo
                solucao_5_bimestre['ativar_restricao'] = 'S'

            res_solucao = pd.concat([solucao_1_bimestre,solucao_2_bimestre,solucao_3_bimestre,solucao_4_bimestre,solucao_5_bimestre])

            solucao = res_solucao.copy()
            solucao = solucao.rename(columns={'PÚBLICO':'grupo_de_atributo','SKU':'sku','ANO CATALOGO':'ano_produto','SÉRIE':'serie_produto','UTILIZAÇÃO':'utilizacao_produto','CODIGO ITENS':'itens','CATEGORIA':'categorias','CUSTOMER GROUP - ESCOLA':'grupos_permissao'})

            categoria = pd.read_excel('itens.xlsx', sheet_name='categoriab2b')
            solucao = pd.merge(solucao,categoria, on=['serie_produto'], how='inner')

            solucao['categorias'] = solucao['MARCA'] + '/' + solucao['categorias']
            solucao = solucao.sort_values(by=['utilizacao_produto','serie_produto'], ascending=True)

            solucao = solucao.rename(columns={'MARCA':'marca_produto'})

            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','itens','ativar_restricao','grupos_permissao']]

            solucao['itens'] = solucao['itens'].apply(lambda x: x[:-1])
            pdt['CODIGO ITENS'] = pdt['CODIGO ITENS'].apply(lambda x: x[:-1])

            solucao_sku = solucao[['sku','nome']]
            solucao_sku = solucao_sku.rename(columns={'nome':'NOME', 'sku':'SKU'})
            pdt_sku = pd.merge(solucao_sku,pdt, on=['SKU'], how='inner')

            solucao['nome'] = solucao['nome'].str.replace('INFANTIL','EI')
            solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
            solucao['nome'] = solucao['nome'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
            solucao['nome'] = solucao['nome'].str.replace('ENSINO MÉDIO','EM')

            pdt_sku['NOME'] = pdt_sku['NOME'].str.replace('INFANTIL','EI')
            pdt_sku['NOME'] = pdt_sku['NOME'].str.replace('FUNDAMENTAL ANOS INICIAIS','EFI')
            pdt_sku['NOME'] = pdt_sku['NOME'].str.replace('FUNDAMENTAL ANOS FINAIS','EFII')
            pdt_sku['NOME'] = pdt_sku['NOME'].str.replace('ENSINO MÉDIO','EM')

            solucao['serie_produto'] = solucao['serie_produto'].str.replace('°','º')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 1','1 ANO')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 2','2 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 3','3 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 4','4 ANOS')
            solucao['serie_produto'] = solucao['serie_produto'].str.replace('Grupo 5','5 ANOS')
            solucao['nome'] = solucao['nome'].str.replace('°','º')
            solucao = solucao.rename(columns={'itens':'items'})

            def encodes(text):
                return text.encode('ascii', errors='ignore').decode('utf-8')

            solucao['sku'] = solucao['sku'].apply(encodes)
            pdt_sku['SKU'] = pdt_sku['SKU'].apply(encodes)

            solucao['sku'] = solucao['sku'].str.replace(' ','')
            solucao['sku'] = solucao['sku'].str.replace('.','')
            pdt_sku['SKU'] = pdt_sku['SKU'].str.replace(' ','')
            pdt_sku['SKU'] = pdt_sku['SKU'].str.replace('.','')


            pdt_final = pdt_sku[['ESCOLA','CNPJ ESCOLA','ANO CATALOGO','SERIAL','SEGMENTO','SÉRIE','UTILIZAÇÃO','PÚBLICO','SKU','NOME','CODIGO ITENS','DESCRIÇÃO ITEM (B2C e B2B)','TOTAL ALUNOS 2023',
                                'PREÇO CATALOGO 2023','DESCONTO POR VOLUME','CUPOM EXTRA DE DESCONTO','TOTAL DESCONTO','TOTAL SEM DESCONTO','TOTAL COM DESCONTO SEM ELT','PREÇO  ANUAL FINAL SEM ELT','PREÇO FINAL COM  ELT (SE APLICÁVEL)',
                                'CUSTOMER GROUP - ESCOLA','SQUAD']]


            solucao['publico_produto'] = 'ALUNO'
            solucao = solucao[['grupo_de_atributo','nome','sku','visibilidade','ano_produto','faturamento_produto','marca_produto','publico_produto','serie_produto','utilizacao_produto','cliente_produto','categorias','items','ativar_restricao','grupos_permissao']]

            pdt_escola = pdt_final.copy()

            pdt_escola = pdt_escola[['SÉRIE','UTILIZAÇÃO','NOME','CODIGO ITENS','DESCRIÇÃO ITEM (B2C e B2B)','TOTAL ALUNOS 2023',
                                'PREÇO CATALOGO 2023','TOTAL DESCONTO','TOTAL COM DESCONTO SEM ELT']]

            pdt_escola['CODIGO ITENS'] = pdt_escola['CODIGO ITENS'].astype(int)
            pdt_escola = pdt_escola.rename(columns={'CODIGO ITENS':'2023'})
            merge_conf = pd.merge(pdt_escola,itens, how='inner',  on=['2023'] )

            pdt_escola['TOTAL ALUNOS 2023'] = pdt_escola['TOTAL ALUNOS 2023'].astype(float)
            pdt_escola['PREÇO CATALOGO 2023'] = pdt_escola['PREÇO CATALOGO 2023'].astype(float)
            pdt_escola['TOTAL ALUNOS 2023'] = pdt_escola['TOTAL ALUNOS 2023'].astype(float)

            pdt_escola['PREÇO COM DESCONTO'] = pdt_escola.apply(lambda x: '{:.2f}'.format(x['PREÇO CATALOGO 2023'] - (x['PREÇO CATALOGO 2023'] * x['TOTAL DESCONTO'])), axis = 1)

            df_input = pdt_final[['CNPJ ESCOLA','SKU','SÉRIE','UTILIZAÇÃO','DESCRIÇÃO ITEM (B2C e B2B)','CODIGO ITENS','CUSTOMER GROUP - ESCOLA']]

            df_brinde_input = pd.read_excel('itens.xlsx', sheet_name='brinde')

            df_input = df_input.rename(columns={'CNPJ ESCOLA':'CNPJ','SKU':'CODIGO MAGENTO','DESCRIÇÃO ITEM (B2C e B2B)':'DESCRIÇÃO ITENS','CODIGO ITENS':'CÓD. ITENS','CUSTOMER GROUP - ESCOLA':'CUSTOMER GROUP'})

            df_input['CÓD. ITENS'] = df_input['CÓD. ITENS'].astype('float')

            df_input = df_input[df_input['DESCRIÇÃO ITENS'].str.contains('KIT|AZ LIVRO ESPIRAL MULTIDISCIPLINAR|PACK')]

            df_brinde = pd.merge(df_input,df_brinde_input, on=['CÓD. ITENS'], how='inner')

            df_brinde_final = df_brinde.copy()
            df_brinde_final = df_brinde_final[['NOME DA REGRA','CUSTOMER GROUP','CODIGO MAGENTO','SKU BRINDE']]
            df_brinde_final = df_brinde_final.rename(columns={'CODIGO MAGENTO':'SKU CONDICAO'})
            df_brinde_final['Status'] = 'ATIVO'

            df_infantil = df_brinde_final.loc[df_brinde['NOME DA REGRA'].str.contains('Grupo')]
            df_infantil['Qtd Incremento'] = 11

            df_seg = df_brinde_final.loc[~df_brinde_final['NOME DA REGRA'].str.contains('Grupo')]
            df_seg['Qtd Incremento'] = 20

            df_brinde_final = pd.concat([df_infantil,df_seg])
            df_brinde_final['Qtd Condicao'] = 1

            df_brinde_final = df_brinde_final.rename(columns={'NOME DA REGRA':'Nome da Regra','CUSTOMER GROUP':'Grupo do Cliente','SKU CONDICAO':'Sku Condicao','SKU BRINDE':'Sku Brinde'})
            df_brinde_final = df_brinde_final[['Nome da Regra','Status','Grupo do Cliente','Sku Condicao','Qtd Condicao','Sku Brinde','Qtd Incremento']]

            df_brinde_final = df_brinde_final.sort_values(by=['Grupo do Cliente','Nome da Regra'])
            #df_brinde_final

            #with pd.ExcelWriter(f'{today}-{cliente_tipo}-{name_escola_completo}.xlsx') as writer:
                #solucao.to_excel(writer, sheet_name="Solução" , index=False)
                #pdt_final.to_excel(writer, sheet_name="Cadastro de itens" , index=False)
                #df_client.to_excel(writer, sheet_name="Produtos da escola" , index=False)
                #df_brinde_final.to_excel(writer, sheet_name="Brinde do professor" , index=False)
                #soma.to_excel(writer, sheet_name="Conferência" , index=False)

            #solucao.to_csv(f'{today}-{cliente_tipo}-{name_escola_completo}-cs.csv', index=False)

            #st.dataframe(solucao)

            with st.spinner('Aguarde...'):
                time.sleep(2)
                st.success('Concluído com sucesso!')

            @st.cache_data
            
            def convert_df(df):
                # IMPORTANT: Cache the conversion to prevent computation on every rerun
                return df.to_csv(index=False).encode('utf-8')
            
            col1, col2, col3 = st.columns(3)

            with col1:
                csv = convert_df(solucao)
                st.download_button(
                    label="Download da solução",
                    data=csv,
                    file_name=f'{today}-{cliente_tipo}-{name_escola_completo}-solucao.csv',
                    mime='text/csv',
                )

            with col2:
                brinde = convert_df(df_brinde_final)
                st.download_button(
                    label="Download do brinde",
                    data=brinde,
                    file_name=f'{today}-{cliente_tipo}-{name_escola_completo}-brinde.csv',
                    mime='text/csv',
                )

            with col3:
                cadastro = convert_df(pdt_final)
                st.download_button(
                    label="Download do cadastro",
                    data=cadastro,
                    file_name=f'{today}-{cliente_tipo}-{name_escola_completo}-cadastro.csv',
                    mime='text/csv',
                )

            st.dataframe(solucao)

            
    


    if choice == "B2C":
        
        st.info('Módulo B2C em construção :)')
            
        
                    


        #st.write(df_geral)
                    
    
