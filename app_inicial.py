import pdfplumber
import pandas as pd
import re


caminho_do_arquivo = "TR_exemplo.pdf"

def extrair_tabela(arquivo:str, numero_da_pagina:int):
    # Para abrir o arquivo pdf
    with pdfplumber.open(arquivo) as pdf:
        # para ler apenas a primeira página do pdf: 
        pagina_selecionada = pdf.pages[numero_da_pagina]
        # Extrair a tabela com as configurações especificadas
        tabela = pagina_selecionada.extract_table(table_settings={"vertical_strategy": "lines_strict", 
                                                                        "horizontal_strategy": "lines_strict"})
        return tabela

def remover_caracteres_ilegais(value):
    # Remove caracteres não imprimíveis usando uma expressão regular
    if isinstance(value, str):
        return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', value)
    return value

def tabela_para_xlsx(tabela_extraida):
    # Verificar se a tabela foi extraída
    if tabela_extraida:
        # Converter a tabela para um DataFrame
        df = pd.DataFrame(tabela_extraida[1:], columns=tabela_extraida[0])  # Usar a primeira linha como cabeçalho
        
        # Remover caracteres ilegais de todas as células do DataFrame
        df = df.applymap(remover_caracteres_ilegais)

        # Exibir o DataFrame (opcional)
        print(df)
        
        # Salvar o DataFrame em um arquivo Excel
        df.to_excel("tabela_extraida.xlsx", index=False)
        
        # Informar que o arquivo foi salvo com sucesso
        print("Tabela salva em 'tabela_extraida.xlsx'")
        

    

tabela = extrair_tabela(caminho_do_arquivo, 0)
tabela_para_xlsx(tabela)

    