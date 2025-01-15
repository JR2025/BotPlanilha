import os
import shutil
import re
from PyPDF2 import PdfReader
import openpyxl

def criar_planilha(caminho_saida, dados):
    """
    Cria ou atualiza uma planilha Excel com os dados fornecidos.

    :param caminho_saida: Caminho onde a planilha será salva.
    :param dados: Lista de dicionários contendo os dados para preencher a planilha.
    Cada dicionário deve ter as chaves: 'CNPJ', 'NFE', 'VALOR', 'CDGBARRA'.
    """
    if not os.path.exists(caminho_saida):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dados"
        ws.append(["CNPJ", "NFE", "VALOR", "CDGBARRA"])
    else:
        wb = openpyxl.load_workbook(caminho_saida)
        ws = wb.active

    for dado in dados:
        ws.append([dado['CNPJ'], dado['NFE'], dado['VALOR'], dado['CDGBARRA']])

    wb.save(caminho_saida)

def extrair_dados_pdf(caminho_pdf):
    """
    Extrai dados de um arquivo PDF.

    :param caminho_pdf: Caminho do arquivo PDF.
    :return: Lista de dicionários com os dados extraídos.
    """
    dados_extraidos = []
    reader = PdfReader(caminho_pdf)

    for page in reader.pages:
        texto = page.extract_text()

        # Regex para extrair CNPJ, NFE, VALOR e CDGBARRA
        cnpj_match = re.search(r'\b\d{14}\b', texto)
        nfe_match = re.findall(r'N[\u00baº]? da NF-e:\s*([\d, ]+)', texto)
        valor_match = re.search(r'17 - Valor Principal\s*([\d.,]+)', texto)
        cdgbarra_match = re.search(r'(\d{11}-\d\s+\d{11}-\d\s+\d{11}-\d\s+\d{11}-\d)', texto)

        if cnpj_match and nfe_match and valor_match and cdgbarra_match:
            nfe_numbers = re.findall(r'\d+', nfe_match[0])
            nfe_concat = ', '.join(nfe_numbers)

            # Limpeza do código de barras - remove os traços '-'
            cdgbarra_clean = re.sub(r'\s+|-', '', cdgbarra_match.group(1))

            dados_extraidos.append({
                'CNPJ': cnpj_match.group(0),
                'NFE': nfe_concat,
                'VALOR': valor_match.group(1).replace(',', '.'),
                'CDGBARRA': cdgbarra_clean  # Código de barras agora sem traços
            })

    return dados_extraidos


def mover_pdf(caminho_pdf, pasta_destino):
    """
    Move um arquivo PDF para uma pasta específica.

    :param caminho_pdf: Caminho do arquivo PDF.
    :param pasta_destino: Pasta de destino.
    """
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    shutil.move(caminho_pdf, os.path.join(pasta_destino, os.path.basename(caminho_pdf)))

def processar_pdfs(pasta_entrada, pasta_processados, caminho_planilha):
    """
    Processa todos os PDFs em uma pasta, extraindo dados e salvando-os em uma planilha.

    :param pasta_entrada: Pasta contendo os PDFs a serem processados.
    :param pasta_processados: Pasta para mover os PDFs processados.
    :param caminho_planilha: Caminho da planilha para salvar os dados extraídos.
    """
    dados_totais = []

    for arquivo in os.listdir(pasta_entrada):
        if arquivo.endswith('.pdf'):
            caminho_pdf = os.path.join(pasta_entrada, arquivo)
            dados = extrair_dados_pdf(caminho_pdf)

            if dados:
                dados_totais.extend(dados)  # Adiciona os dados extraídos à lista total
                mover_pdf(caminho_pdf, pasta_processados)  # Move o PDF processado
            else:
                print(f"Nenhum dado extraído do arquivo: {arquivo}")

    if dados_totais:
        criar_planilha(caminho_planilha, dados_totais)  # Cria ou atualiza a planilha com os dados extraídos

def iniciar_extracao(pasta_entrada, caminho_planilha):
    """
    Inicia o processo de extração de dados dos PDFs e cria a planilha.

    :param pasta_entrada: Caminho da pasta com os arquivos PDF.
    :param caminho_planilha: Caminho para salvar a planilha.
    """
    pasta_processados = os.path.join(pasta_entrada, "processados")
    processar_pdfs(pasta_entrada, pasta_processados, caminho_planilha)
