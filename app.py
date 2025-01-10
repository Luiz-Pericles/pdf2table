import pdfplumber
import pandas as pd
import re

caminho_do_arquivo = "TR_exemplo.pdf"

def extrair_tabelas(arquivo:str, paginas:list):
    todas_as_tabelas = []

    with pdfplumber.open(arquivo) as pdf:
        for numero_da_pagina in paginas:
            pagina_selecionada = pdf.pages[numero_da_pagina]
            tabela = pagina_selecionada.extract_table(table_settings={"vertical_strategy": "lines_strict", 
                                                                      "horizontal_strategy": "lines_strict"})
            if tabela:
                todas_as_tabelas.extend(tabela[1:] if todas_as_tabelas else tabela)  # Pular cabeçalho após a primeira

    return todas_as_tabelas

def remover_caracteres_ilegais(value):
    if isinstance(value, str):
        # Remove caracteres não imprimíveis e substitui por string vazia
        return ''.join(ch for ch in value if ch.isprintable())
    return value

def tabelas_para_xlsx(tabelas_extraidas):
    if tabelas_extraidas:
        cabecalho = tabelas_extraidas[0]
        df = pd.DataFrame(tabelas_extraidas[1:], columns=cabecalho)
        df = df.applymap(remover_caracteres_ilegais)
        print(df)
        df.to_excel("tabela_consolidada.xlsx", index=False)
        print("Tabela consolidada salva em 'tabela_consolidada.xlsx'")
    else:
        print("Nenhuma tabela encontrada.")

# Especificar as páginas que contêm as tabelas
paginas_com_tabelas = [ 0, 1, 2, 3 ]  # Índices zero para páginas 1, 2 e 3

tabelas = extrair_tabelas(caminho_do_arquivo, paginas_com_tabelas)
print(tabelas)
print(type(tabelas))
tabelas_para_xlsx(tabelas) 
