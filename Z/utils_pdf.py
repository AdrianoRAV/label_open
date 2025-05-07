import os
import shutil
import time
import fitz
import re
import pandas as pd
pasta_raiz = os.path.abspath(os.getcwd())
diretorio_downloads = os.path.join(pasta_raiz)

def renomear_pdf(novo_nome):
    for _ in range(10):  # Tenta por 10 segundos
        arquivos = [f for f in os.listdir(diretorio_downloads) if f.endswith(".pdf")]
        if arquivos:
            arquivo_baixado = arquivos[0]
            caminho_atual = os.path.join(diretorio_downloads, arquivo_baixado)
            novo_caminho = os.path.join(pasta_raiz, novo_nome + ".pdf")
            shutil.move(caminho_atual, novo_caminho)
            print(f"PDF renomeado e movido para: {novo_caminho}")
            return True
        time.sleep(1)
    print("Erro: Arquivo PDF não encontrado para renomeação.")
    return False


def pdf_to_excelucub(nome, nome_excel):
    # Caminho do arquivo PDF
    pdf_path = nome

    # Lista para armazenar os códigos de lacre
    lacres = []

    # Abre o PDF e percorre cada página
    with fitz.open(pdf_path) as pdf:
        for page_num in range(pdf.page_count):
            page = pdf[page_num]
            text = page.get_text()  # Extrai o texto da página

            # Usa regex para encontrar todos os códigos começando com "UC" ou "UB"
            lacre_codes = re.findall(r'\bU[CB]\d{9}\b', text)
            lacres.extend(lacre_codes)  # Adiciona os códigos à lista

    # Remove duplicatas caso existam
    lacres = list(set(lacres))

    # Salva os códigos em um DataFrame do pandas
    df = pd.DataFrame(lacres, columns=["Lacre"])

    # Exporta o DataFrame para um arquivo Excel
    df.to_excel(nome_excel, index=False)

    print(f"Códigos de lacre salvos no arquivo {nome_excel}")