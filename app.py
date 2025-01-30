import pdfplumber
import pandas as pd
import streamlit as st
import io





def extrair_tabelas(arquivo:str, paginas:list):
    todas_as_tabelas = []

    with pdfplumber.open(arquivo) as pdf:
        for numero_da_pagina in paginas:
            pagina_selecionada = pdf.pages[numero_da_pagina]
            tabela = pagina_selecionada.extract_table(table_settings={"vertical_strategy": "lines_strict", 
                                                                      "horizontal_strategy": "lines_strict"})
            if tabela:
                todas_as_tabelas.extend(tabela)  # Pular cabeçalho após a primeira
            else:
                print(f'Não foi encontrado tabela na página: {numero_da_pagina+1}')
    return todas_as_tabelas

def remover_caracteres_ilegais(value):
    if isinstance(value, str):
        # Remove caracteres não imprimíveis e substitui por string vazia
        return ''.join(ch for ch in value if ch.isprintable())
    return value

@st.cache_data
def tabelas_para_xlsx(tabelas_extraidas):
    if tabelas_extraidas:
        cabecalho = tabelas_extraidas[0]
        df = pd.DataFrame(tabelas_extraidas[1:], columns=cabecalho)
        df = df.map(remover_caracteres_ilegais)
        print(df)
        df.to_excel("tabela_consolidada.xlsx", index=False)
        print("Tabela consolidada salva em 'tabela_consolidada.xlsx'")
    else:
        print("Nenhuma tabela encontrada.")
    return df





# Criar interface streamlit

st.write("# App *PDF2TABLE* :alien:")
    # Entrada de dados do usuário: 
# Especificar as páginas que contêm as tabelas
with st.form('my_form'):
    
    # ultima_pag_com_tabela = int(st.selectbox("Qual o número da página com o fim da tabela: ", ('1', '2', '3', '4', '5', '6', '7', '8', '9', '10')))
    ultima_pag_com_tabela = int(st.number_input("Qual o número da página com o fim da tabela: ", min_value=1, max_value=20, step=1))
    submitted = st.form_submit_button('Salvar')
st.write(ultima_pag_com_tabela)
# Criar uma lista com as páginas que possui a tabela
paginas_com_tabelas = list()
for i in range(0, ultima_pag_com_tabela):
    paginas_com_tabelas.append(i)

# Para o usuário selecionar o arquivo pdf
st.subheader("Selecione ou arraste o Termo de Referência em PDF")
uploaded_file = st.file_uploader("suba o arquivo pdf", type=['pdf'], label_visibility="hidden")
if uploaded_file is not None:
        
    # Função para extrair a tabela do pdf, retorna um DF
    tabelas = extrair_tabelas(uploaded_file, paginas_com_tabelas)
    # Função para transformar o DF em arquivo xlsx
    tabela_tratada = tabelas_para_xlsx(tabelas) 

    # Escrever a tabela extraida na tela
    st.write("Segue a extração da tabela: ")
    st.write(tabela_tratada)
    
    # Botão de download do excel para streamlit
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        tabela_tratada.to_excel(writer, sheet_name='Tabela_do_TR')
        # Fechar o ExcelWriter e devolver o arquivo Excel para o buffer
        writer.close()
        
        st.download_button(label="Baixe a tabela de itens em Excel", data=buffer, file_name="Itens_do_TR.xlsx", mime="application/vnd.ms-excel")

else:
    st.info(' Faça o upload de um Termo de Referência em PDF')

st.write("## FIM")

