import os
import re
import PyPDF2
import unicodedata
from collections import defaultdict

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
import xml.etree.ElementTree as ET
from tkinter import Tk
from tkinter.filedialog import askdirectory


# Listas de palavras para remoção
palavras_para_removern = ["GREENPLAC", "DURATEX", "ARAUCO", "GUARARAPES", "ESSENCIAL", "WOOD", "MULTIMARCAS"]
palavras_para_remover2n = ["MADEIRAS"]
palavras_para_remover = []
palavras_para_remover2 = []

# Função para extrair dados
def extrair_gplan_pdf(pasta_vendedor):
    print("g")
    # Dicionário para armazenar os resultados
    resultado = defaultdict(int)

    # Iterar sobre todos os arquivos na pasta do vendedor
    for arquivo in os.listdir(pasta_vendedor):
        if "zMDF" in arquivo and arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(pasta_vendedor, arquivo)
            
            # Abrir e processar o PDF
            with open(caminho_pdf, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                texto_completo = ''
                
                # Concatenar o texto de todas as páginas
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    texto_completo += page.extract_text()

            # Dividir o texto em linhas para pegar a segunda linha
            linhas = texto_completo.split('\n')
            if len(linhas) >= 2:
                nome = linhas[1]  # A segunda linha contém o nome

                # Remove os acentos
                nome_sem_acentos = remover_acentos(nome)
                
                # Remover todas as ocorrências das palavras da lista do texto
                for palavra in palavras_para_remover:
                    nome_sem_acentos = nome_sem_acentos.replace(palavra, "").strip()
                
                # Remover palavras específicas na primeira posição
                primeira_palavra = nome_sem_acentos.split()[0] if nome_sem_acentos.split() else ""
                if primeira_palavra in palavras_para_remover2:
                    nome_sem_acentos = nome_sem_acentos[len(primeira_palavra):].strip()

                # Remover espaços extras entre palavras
                nome_sem_acentos = re.sub(r'\s+', ' ', nome_sem_acentos).strip()    

                
                # Procurar os códigos e somar as quantidades
                quantidades = re.findall(r'Código(\d) \*', texto_completo)
                quantidade_chapas = sum(map(int, quantidades))

                # Atualizar os resultados
                resultado[nome_sem_acentos] += quantidade_chapas
                

    # print(texto_completo)
    print(resultado)
    # Retornar o dicionário
    return dict(resultado)

# Função para remover acentos e converter para maiúsculas
def remover_acentos(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    ).upper()

def extrair_nesting_pdf(pasta, nc_files_data=None):    
    print("n")
    if not nc_files_data:
        corte = os.path.join(pasta, "Nesting")
        caminho_xml = os.path.join(corte, 'InfoOutput.xml')
        # Parse do XML
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        nc_files_data = []
        for file in root.findall(".//NcFiles/File"):
            file_name = file.get("name")
            nc_files_data.append([file_name])
    
    # Expressão regular para encontrar os nomes das chapas
    padrao = r'\d+_\d+_\d+_(.+? \d+mm)'
    
    # Contar as ocorrências de cada tipo de chapa
    contador = defaultdict(int)
    
    # Loop através dos dados fornecidos
    for item in nc_files_data:
        texto_completo = item[0]  # item é uma lista com uma string, então usamos item[0]
        
        # Encontrar chapas usando a expressão regular
        chapas_encontradas = re.findall(padrao, texto_completo)
        
        for chapa in chapas_encontradas:
            # Converter para maiúsculas
            chapa_upper = remover_acentos(chapa)
        
            # Remover palavras específicas de todo o texto da chapa
            for palavra in palavras_para_remover:
                chapa_upper = chapa_upper.replace(palavra, "").strip()
            
            # Remover palavras específicas na primeira posição
            primeira_palavra = chapa_upper.split()[0] if chapa_upper.split() else ""
            if primeira_palavra in palavras_para_remover2:
                chapa_upper = chapa_upper[len(primeira_palavra):].strip()
                
            # Remover espaços extras entre palavras
            chapa_upper = re.sub(r'\s+', ' ', chapa_upper).strip()    
                
            # Incrementar o contador
            contador[chapa_upper] += 1
    
    # print(texto_completo)   
    print(contador)
    return dict(contador)

def gerar_pdf_com_tabela(pasta_vendedor, pasta):
    try:
        resultado_nesting = extrair_nesting_pdf(pasta)
    except Exception as e:
        print('deu um erro ae:', e)  

    try:
        resultado_gplan = extrair_gplan_pdf(pasta_vendedor)
    except Exception as e:
        print('deu um erro ae:', e)  

    resultado_nesting = {k: resultado_nesting[k] for k in sorted(resultado_nesting) if k != 'total'}
    resultado_gplan = {k: resultado_gplan[k] for k in sorted(resultado_gplan) if k != 'total'}

    total_nesting = sum(resultado_nesting.values())
    total_gplan = sum(resultado_gplan.values())

    todas_chaves = sorted(set(resultado_nesting.keys()).union(set(resultado_gplan.keys())))

    caminho_arquivo_pdf = os.path.join(pasta_vendedor, "Contagem de Chapa.pdf")
    c = canvas.Canvas(caminho_arquivo_pdf, pagesize=letter)
    width, height = letter

    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 40, "Contagem de Chapas Nesting/GPlan:")

    c.setFont("Helvetica-Bold", 10)

    # Novas posições ajustadas
    pos_produto = 30
    pos_nesting = 400  # mais próximo do final
    pos_gplan = 460    # bem perto do nesting

    c.drawString(pos_produto, height - 70, "Produto")
    c.drawString(pos_nesting, height - 70, "NEST.")
    c.drawString(pos_gplan, height - 70, "GPL.")

    c.line(30, height - 75, width - 30, height - 75)

    y_position = height - 90

    for idx, chave in enumerate(todas_chaves):
        quantidade_nesting = resultado_nesting.get(chave, 0)
        quantidade_gplan = resultado_gplan.get(chave, 0)

        if quantidade_nesting == quantidade_gplan:
            c.setFillColor(colors.white)
        else:
            c.setFillColor(colors.lightpink)

        c.rect(30, y_position , width - 60, 15, fill=1)

        c.setFont("Helvetica", 10)
        c.setFillColor(colors.black)

        c.drawString(pos_produto + 5, y_position + 3, chave)
        c.drawRightString(pos_nesting + 20, y_position + 3, str(quantidade_nesting))
        c.drawRightString(pos_gplan + 20, y_position + 3, str(quantidade_gplan))

        y_position -= 15

    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(colors.black)
    c.drawString(pos_produto + 5, y_position, "TOTAL:")
    c.drawRightString(pos_nesting + 20, y_position, str(total_nesting))
    c.drawRightString(pos_gplan + 20, y_position, str(total_gplan))

    c.save()
    return caminho_arquivo_pdf


# Exemplo de uso
# pasta_vendedor = "./PROJJE_IKAD CCB SALAS ADM\\VENDEDOR"
# arquivo_pdf = gerar_pdf_com_tabela(pasta_vendedor)
# print(f"Resultado gerado no arquivo PDF: {arquivo_pdf}")


def main():

    # Abrir a janela de seleção de arquivo
    Tk().withdraw()  # Evitar que a janela principal do Tkinter apareça
    pasta = askdirectory(title="Selecione a pasta do projeto")
    print(pasta)

    pasta_vendedor = os.path.join(pasta, "VENDEDOR")
    arquivo_pdf = gerar_pdf_com_tabela(pasta_vendedor, pasta)
    print(f"Resultado gerado no arquivo PDF: {arquivo_pdf}")

if __name__ == "__main__":
    main()

