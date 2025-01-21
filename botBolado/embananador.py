import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.units import inch
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename
import os
import random

def formatar_valores(data):
    """ Remove o .0 dos valores flutuantes e formata os números inteiros """
    formatted_data = []
    for row in data:
        new_row = []
        for cell in row:
            if isinstance(cell, float) and cell.is_integer():
                new_row.append(int(cell))
            else:
                new_row.append(cell)
        formatted_data.append(new_row)
    return formatted_data

def gerar_relatorio_pecas(df, arquivo_excel, nome=None):   
    print("RP")
    # Substitui NaN por string vazia
    df = df.fillna('')

    # Aplicar a lógica de exclusão de linhas
    mask_exclude = (
        ((df['PEÇA DESCRIÇÃO'].str.contains('_PAINEL_DUP_', na=False)) & (df['ESPESSURA'].isin([15, 18]))) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_PRAT_DUP_CORTE', na=False)) |
        (df['PEÇA DESCRIÇÃO'].str.contains('AFAST_DUP_CORTE', na=False)) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_PAINEL_ENG_CORTE', na=False)) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_ENGROSSO_', na=False)) |
        ((df['PEÇA DESCRIÇÃO'].str.contains('_ENG', na=False)) & (df['ESPESSURA'] == '6'))
    )
    
    # Filtrar o DataFrame removendo as linhas que atendem à condição mask_exclude
    df_filtered = df[~mask_exclude]

    # Filtrar e organizar os dados
    relatorio_pecas = df_filtered[['PEÇA DESCRIÇÃO', 'CLIENTE - DADOS DO CLIENTE', 'ALTURA (X)', 'PROF (Y)', 'ESPESSURA', 'AMBIENTE', 'DESENHO']]

    relatorio_pecas = relatorio_pecas.rename(columns={
        'PEÇA DESCRIÇÃO': 'PEÇA DESCRIÇÃO',
        'CLIENTE - DADOS DO CLIENTE': 'CLIENTE',
        'ALTURA (X)': 'ALT (X)',
        'PROF (Y)': 'PROF (Y)',
        'ESPESSURA': 'ESP (Z)',
        'AMBIENTE': 'AMBIENTE',
        'DESENHO': 'DESENHO'
    })
    
    # Adicionar a coluna VISTO
    relatorio_pecas['VISTO'] = ''
    
    # Organizar por ALTURA (X) de forma decrescente e depois por LARGURA (Y) de forma decrescente
    relatorio_pecas = relatorio_pecas.sort_values(by=['ALT (X)', 'PROF (Y)'], ascending=[False, False])

    # Adicionar a coluna NUMERADOR
    relatorio_pecas['NUM'] = range(1, len(relatorio_pecas) + 1)

    # Formatar valores
    data = [relatorio_pecas.columns.tolist()] + formatar_valores(relatorio_pecas.values.tolist())
    
    # Salvar o relatório como PDF
    pasta_arquivo = os.path.dirname(arquivo_excel)  # Obtém o diretório onde o arquivo Excel está localizado
    if nome is None:
        nome = obter_nome(pasta_arquivo)
    print('pasta_arquivo', pasta_arquivo)
    pasta_arquivo = pasta_arquivo+"\VENDEDOR" 
    file_name = os.path.join(pasta_arquivo, f'Relatorio_Pecas_{nome}.pdf')
    print(file_name)
    # Ajustar margens para usar mais espaço na página
    margins = {'rightMargin': 0.3 * inch, 'leftMargin': 0.3 * inch, 
               'topMargin': 0.3 * inch, 'bottomMargin': 0.3 * inch}
    doc = SimpleDocTemplate(file_name, pagesize=landscape(letter), **margins)
    elements = []

    # Criar a tabela
    table = Table(data)

    # Define o estilo da tabela
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.black),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),   # Justificar à esquerda para 'PEÇA'
        ('ALIGN', (2, 1), (2, -1), 'RIGHT'),  # Justificar à direita para 'ALTURA (X)'
        ('ALIGN', (3, 1), (3, -1), 'RIGHT'),  # Justificar à direita para 'PROF (Y)'
        ('ALIGN', (4, 1), (4, -1), 'RIGHT'),  # Justificar à direita para 'ESPESSURA'
        ('ALIGN', (5, 1), (5, -1), 'LEFT'),   # Justificar à esquerda para 'AMBIENTE'
        ('ALIGN', (6, 1), (6, -1), 'LEFT'),   # Justificar à esquerda para 'DESENHO'
        ('ALIGN', (7, 1), (7, -1), 'RIGHT'),  # Justificar à direita para 'NUMERADOR'
        ('ALIGN', (1, 1), (1, -1), 'LEFT'),   # Justificar à esquerda para 'CLIENTE'
        ('BOX', (0, 0), (-1, -1), 0.2, colors.black),  # Espessura das bordas
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),  # Espessura das linhas de grade
        ('TOPPADDING', (0, 0), (-1, -1), 10),  # Adiciona espaçamento na parte Superior
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Adiciona espaçamento na parte inferior
        ('BOTTOMPADDING', (1, 1), (1, -1), -1),  # Adiciona espaçamento na parte inferior para a coluna 'CLIENTE'
        ('BOTTOMPADDING', (5, 1), (5, -1), -1),  # Adiciona espaçamento na parte inferior para a coluna 'AMBIENTE'
        ('BOTTOMPADDING', (0, 0), (0, -1), 0),  # Adiciona espaçamento na parte inferior para a coluna 'PECA'
    ])

    # Alternar fundo cinza e branco
    num_rows = len(data)
    for i in range(1, num_rows):
        if i % 2 == 1:
            style.add('BACKGROUND', (0, i), (-1, i), colors.Color(0.9, 0.9, 0.9))  # Cinza claro
        else:
            style.add('BACKGROUND', (0, i), (-1, i), colors.white)
    
    table.setStyle(style)

    # Ajustar a largura das colunas para caber no texto
    largura_total = landscape(letter)[0] - (margins['leftMargin'] + margins['rightMargin'])  # Largura total disponível na página
    
    # Definir larguras específicas para as colunas
    largura_colunas = [
        2.0 * inch,  # 'PEÇA'
        2.3 * inch,  # 'CLIENTE'
        0.6 * inch,  # 'ALTURA (X)'
        0.6 * inch,  # 'PROF (Y)'
        0.6 * inch,  # 'ESPESSURA'
        1.9 * inch,  # 'AMBIENTE'
        0.9 * inch,  # 'DESENHO'
        0.8 * inch,  # 'VISTO'
        0.4 * inch,  # 'NUMERADOR'
    ]
    
    # Defina as alturas das linhas
    row_heights = [0.2 * inch] * len(data)  # Define a altura de todas as linhas

    # Configure manualmente a altura das linhas
    for i, height in enumerate(row_heights):
        table._argH[i] = height

    # Ajustar as larguras das colunas para garantir que caibam na largura total disponível
    while sum(largura_colunas) > largura_total:
        for i in range(len(largura_colunas)):
            if largura_colunas[i] > 0.5 * inch:  # Minimizar até um certo ponto
                largura_colunas[i] -= 0.1 * inch
    
    table._argW = largura_colunas

    # Ajustar o tamanho da fonte para que o texto caiba
    pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
    table.setStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),  # Tamanho padrão para todas as colunas
        ('FONTSIZE', (1, 1), (1, -1), 6),    # Tamanho da fonte para a coluna 'CLIENTE'
        ('FONTSIZE', (5, 1), (5, -1), 6),    # Tamanho da fonte para a coluna 'AMBIENTE'
        ('FONTSIZE', (0, 0), (0, -1), 8),    # Tamanho da fonte para a coluna 'PECA'
    ])

    elements.append(table)
    doc.build(elements)

def obter_nome(pasta_arquivo):
    try:
        # Extrair o nome da pasta
        nome_pasta = os.path.basename(pasta_arquivo)  # Obtém o nome da pasta
        print('Nome da pasta:', nome_pasta)
        
        # Verificar se o nome da pasta contém pelo menos 2 partes separadas por espaços
        partes = nome_pasta.split(' ')
        if len(partes) >= 2:
            # Extrair as partes após o primeiro espaço
            nome = ' '.join(partes[1:])  # Pega todas as partes após o primeiro espaço
            print('Nome extraído:', nome)
            return nome
        else:
            raise ValueError(f"O nome da pasta não contém o formato esperado. Pasta: {nome_pasta}")
    except Exception as e:
        print(f"Erro ao ler o arquivo PDF: {e}")
        return ""
    
def contar_pecas(df):
    # Substitui NaN por string vazia
    df = df.fillna('')

    # Aplicar a lógica de exclusão de linhas
    mask_exclude = (
        ((df['PEÇA DESCRIÇÃO'].str.contains('_PAINEL_DUP_', na=False)) & (df['ESPESSURA'].isin([15, 18]))) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_PRAT_DUP_CORTE', na=False)) |
        (df['PEÇA DESCRIÇÃO'].str.contains('AFAST_DUP_CORTE', na=False)) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_PAINEL_ENG_CORTE', na=False)) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_ENGROSSO_', na=False)) |
        ((df['PEÇA DESCRIÇÃO'].str.contains('_ENG', na=False)) & (df['ESPESSURA'] == '6'))
    )

    # Filtrar o DataFrame removendo as linhas que atendem à condição mask_exclude
    df_filtered = df[~mask_exclude]
    
    # Contar a quantidade de peças (número de linhas no DataFrame filtrado)
    quantidade_pecas = len(df_filtered)
    
    # Contar a quantidade total de peças (sem filtro)
    quantidade_total = len(df)
    
    return quantidade_pecas, quantidade_total

def criar_arquivo_com_pecas(df, arquivo):
    print("NP")
    print(arquivo)
    
    pasta_arquivo = os.path.dirname(arquivo) 

    # Obter as quantidades de peças
    quantidade_pecas, quantidade_total = contar_pecas(df)
    
    # Criar o nome do arquivo
    nome = f"zTotal_Pecas__{quantidade_pecas}__.txt"
    pasta_arquivo = os.path.join(pasta_arquivo, "VENDEDOR") 
    nome_arquivo =  os.path.join(pasta_arquivo, nome)
    
    # Criar o conteúdo do arquivo
    # conteudo = f"Quantidade de peças do projeto {nome_projeto} tem {quantidade_pecas} peças acabadas e {quantidade_total} peças de corte."
    conteudo = f"TOTAL PECAS: __{quantidade_pecas}__"
    
    # Escrever no arquivo
    with open(nome_arquivo, 'w') as arquivo:
        arquivo.write(conteudo)
    
    print(f"Arquivo {nome_arquivo} criado com sucesso.")

import pandas as pd
from fpdf import FPDF

def arquivo_ripado(df, arquivo, nome=None):
    
    # Filtrar as linhas que têm "TIRA_RIPADO" na coluna 'PEÇA DESCRIÇÃO'
    df_tira_ripado = df[df['PEÇA DESCRIÇÃO'] == '_TIRA_RIPADO']

    # Se não houver registros, não gera o PDF e imprime uma mensagem
    if df_tira_ripado.empty:
        print("Nenhum registro com '_TIRA_RIPADO' ou '45G' encontrado. Nenhum PDF gerado.")
        return

    pasta_arquivo = os.path.dirname(arquivo) 
    if nome is None:
        nome = obter_nome(pasta_arquivo)
    pasta_arquivo = os.path.join(pasta_arquivo, "VENDEDOR")
    pdf_nome = f'cRelatorio Frente_{nome}.pdf'
    
    # Função para calcular o valor da coluna "Roteiro"
    def calcular_roterio(row):
        serra = 4
        calculo = (row['PROF (Y)'] - serra) / 2
        return f"Abrir em {calculo:.0f}mm"

    # Criar a coluna "Roteiro"
    df_tira_ripado['Roteiro'] = df_tira_ripado.apply(calcular_roterio, axis=1)

    # Criar o PDF com FPDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Definir título
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(50, 3, txt=f"Relatório Tira Ripado - {nome}", ln=True, align='C')
    pdf.ln()
    
    # Definir cabeçalho
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 5, "PEÇA DESCRIÇÃO", border=1, align='C')
    pdf.cell(50, 5, "CLIENTE", border=1, align='C')
    pdf.cell(10, 5, "MAT", border=1, align='C')
    pdf.cell(10, 5, "ALT", border=1, align='C')
    pdf.cell(10, 5, "PROF", border=1, align='C')
    pdf.cell(50, 5, "DESCRIÇÃO DO MATERIAL", border=1, align='C')
    pdf.cell(30, 5, "Roteiro", border=1, align='C')
    pdf.ln()

    # Definir as cores de fundo alternadas
    fundo_cinza = [220, 220, 220]  # Cor cinza claro
    fundo_branco = [255, 255, 255]  # Cor branca
    cor_fundo = fundo_branco  # Começar com fundo branco
    for index, row in df_tira_ripado.iterrows():
        pdf.set_font('Arial', '', 8)

        # Alterna a cor de fundo
        pdf.set_fill_color(*cor_fundo)

        # Adiciona os dados da linha
        pdf.cell(30, 6, str(row['PEÇA DESCRIÇÃO']), border=1, align='C', fill=True)
        pdf.cell(50, 6, str(row['CLIENTE - DADOS DO CLIENTE']), border=1, align='C', fill=True)
        pdf.cell(10, 6, str(row['CÓDIGO MATERIAL']), border=1, align='C', fill=True)
        pdf.cell(10, 6, str(row['ALTURA (X)']), border=1, align='C', fill=True)
        pdf.cell(10, 6, str(row['PROF (Y)']), border=1, align='C', fill=True)
        pdf.cell(50, 6, str(row['DESCRIÇÃO DO MATERIAL']), border=1, align='C', fill=True)
        pdf.cell(30, 6, str(row['Roteiro']), border=1, align='C', fill=True)
        pdf.ln()

        # Alterna a cor de fundo para a próxima linha
        cor_fundo = fundo_cinza if cor_fundo == fundo_branco else fundo_branco

        # Se a página estiver cheia, criar uma nova
        if pdf.get_y() > 275:  # Ajuste esse valor conforme necessário
            pdf.add_page()
            pdf.set_font('Arial', 'B', 7)
            pdf.cell(30, 5, "PEÇA DESCRIÇÃO", border=1, align='C')
            pdf.cell(50, 5, "CLIENTE", border=1, align='C')
            pdf.cell(10, 5, "MAT", border=1, align='C')
            pdf.cell(10, 5, "ALT", border=1, align='C')
            pdf.cell(10, 5, "PROF", border=1, align='C')
            pdf.cell(50, 5, "DESCRIÇÃO DO MATERIAL", border=1, align='C')
            pdf.cell(30, 5, "Roteiro", border=1, align='C')
            pdf.ln()
            
    # Salvar o PDF
    pdf.output(pasta_arquivo+"/"+pdf_nome)

    print(f"PDF gerado com sucesso: {pdf_nome}")

### funcao p importar  
import pdfplumber
import glob
import json
import re
from collections import Counter

def gerar_aciete(pasta_arquivo):

    def ler_pdf(pasta_arquivo, nome_arquivo):
        # Usando glob para pegar todos os arquivos que correspondem ao padrão
        caminho_pdf = glob.glob(os.path.join(pasta_arquivo, nome_arquivo))
        
        if not caminho_pdf:  # Verificação se o arquivo não foi encontrado
            print("Nenhum arquivo encontrado com o padrão:", nome_arquivo)
            return False

        caminho = caminho_pdf[0]
            
        conteudo = ""
        try:
            with pdfplumber.open(caminho) as pdf:
                # Itera sobre todas as páginas
                for pagina in pdf.pages:
                    # Extrai o texto da página
                    conteudo += pagina.extract_text()
            return conteudo
        
        except Exception as e:
            print(f"Erro ao ler o arquivo PDF {nome_arquivo}: {e}")
            return False
            
    def get_op(pasta_arquivo):
        # Usando o glob para procurar o arquivo com o padrão correto
        texto = glob.glob(os.path.join(pasta_arquivo, "planoCorte_Moveo_Ecomobile_OP_*_Cut.xls"))

        if texto:
            # Pegando o nome do arquivo encontrado
            nome_arquivo = os.path.basename(texto[0])
            
            # Usando regex para procurar o número após 'OP_'
            match = re.search(r'OP_(\d+)_', nome_arquivo)
            
            if match:
                return match.group(1)
        
        return None
   
    def get_totTiraRipado(pasta_vendedor):
        texto = ler_pdf(pasta_vendedor, "Listagem_Pecas.pdf")
        if texto:
            ripado = re.findall(r'_TIRA RIPADO', texto, flags=re.IGNORECASE)
            # Encontrar as ocorrências de 'Abrir em Xmm', onde X é qualquer valor numérico
            valores = re.findall(r'Abrir em (\d+)mm', texto)
            
            # Contar as ocorrências dos valores encontrados
            contagem_valores = Counter(valores)
            
            # Criar a string formatada para os valores
            valores_formatados = " ".join([f"({contagem}){valor}mm" for valor, contagem in contagem_valores.items()])
            
            # Retorna o número de ocorrências de '_TIRA RIPADO' e os valores formatados
            return len(ripado), valores_formatados
        
        return None, None
            
    def get_totPainelRouter(pasta_vendedor):
        texto = ler_pdf(pasta_vendedor, "Router.pdf")
        if texto:
            router = re.findall(r'_PAINEL ROUTER', texto, flags=re.IGNORECASE)
            return len(router)  # Retorna o número de ocorrências encontradas
        return None
    
    def get_totEngrosso(pasta_vendedor):
        # Lê o conteúdo do PDF
        texto = ler_pdf(pasta_vendedor, "Composto.pdf")
        
        if texto:
            # Procura por "DUPLADO" ou "ENGROSSADO"
            engrosso = re.findall(r'(DUPLADO|ENGROSSADO)', texto, flags=re.IGNORECASE)
            afastador = re.findall(r'AFASTADOR', texto, flags=re.IGNORECASE)
            # Busca os valores "15" e "18"
            esp = ""
            if '15mm' in texto:
                esp += "31mm"
            if '18mm' in texto:
                if esp:
                    esp += ", "  # Adiciona a vírgula se já encontrou "15"
                esp += "37mm"
            soma = len(engrosso) + len(afastador) * 2 
            return soma, esp  # Retorna o número de ocorrências e os valores encontrados de esp
        return None, None

    def get_totFrente45(pasta_vendedor):
        texto = ler_pdf(pasta_vendedor, "Frentes.pdf")
        if texto:
            usinagem = re.findall(r'(CORTE 45G|PERFIL 45G)', texto, flags=re.IGNORECASE)
            return len(usinagem)  # Retorna o número de ocorrências encontradas
        return None
    
    def get_totVarios(pasta_vendedor):
        texto = ler_pdf(pasta_vendedor, "ListagemCompleta.pdf")
        if texto:
            # Define a lista de regex para diferentes campos
            regexData = [
                { 'regex': r'(\d+)\s*(ML|M2|UN)\s*(ser_cor_lam_45g|lam_topo|ser_lam_lar|ser_lam_est)', 'field': 'fitagem' },
                { 'regex': r'(\d+)\s*(ML|M2|UN)\s*furo_cnc_(20mm|10mm|3mm|15mm|1mm|5mm|8mm)', 'field': 'furosist' },
                { 'regex': r'(\d+)\s*(ML|M2|UN)\s*(usi_rebaixo_7|usi_rebaixo_4|usi_rasgo_7|usi_rasgo_4)', 'field': 'canal' },
                { 'regex': r'(\d+)\s*(ML|M2|UN)\s*furo_cnc_35mm', 'field': 'furodob' },
                { 'regex': r'(\d+)\s*(ML|M2|UN)\s*ser_corte_015', 'field': 'cortes' },
                { 'regex': r'(\d+)\s*(ML|M2|UN)\s*servico_(instal_perfil|corte_barra)_015', 'field': 'cortePerfil' }
            ]
            
            # Dicionário para armazenar as somas de cada campo
            result = {
                'fitagem': 0,
                'furosist': 0,
                'canal': 0,
                'furodob': 0,
                'cortes': 0,
                'cortePerfil': 0,
                'tipoCanal': ''
            }

            # Flag para identificar se "usi_rasgo" ou "usi_rebaixo" aparecem para 7MM e 4MM
            rasgo7_found = False
            canal7_found = False
            rasgo4_found = False
            canal4_found = False

            # Itera sobre cada regexData e faz a contagem para cada campo
            for item in regexData:
                matches = re.findall(item['regex'], texto, flags=re.IGNORECASE)
                
                # Para cada campo, soma os valores encontrados
                for match in matches:
                    # A quantidade que deve ser somada é o primeiro grupo (o número encontrado)
                    result[item['field']] += int(match[0])  # match[0] é o número encontrado
                    
                    # Verifica se as condições para tipoCanal são atendidas
                    if item['field'] == 'canal':
                        if 'usi_rasgo_7' in match[2]:
                            rasgo7_found = True
                        if 'usi_rebaixo_7' in match[2]:
                            canal7_found = True
                        if 'usi_rasgo_4' in match[2]:
                            rasgo4_found = True
                        if 'usi_rebaixo_4' in match[2]:
                            canal4_found = True

            # Define o valor de tipoCanal com base nas condições
            if rasgo7_found and canal7_found:
                result['tipoCanal'] = 'CANAL 7MM E REBAIXO 7MM'
            elif rasgo4_found and canal4_found:
                result['tipoCanal'] = 'CANAL 4MM E REBAIXO 4MM'
            elif rasgo7_found and canal4_found:
                result['tipoCanal'] = 'CANAL 7MM E REBAIXO 4MM'
            elif rasgo4_found and canal7_found:
                result['tipoCanal'] = 'CANAL 4MM E REBAIXO 7MM'
            elif rasgo7_found:
                result['tipoCanal'] = 'CANAL 7MM'
            elif canal7_found:
                result['tipoCanal'] = 'REBAIXO 7MM'
            elif rasgo4_found:
                result['tipoCanal'] = 'CANAL 4MM'
            elif canal4_found:
                result['tipoCanal'] = 'REBAIXO 4MM'
                        
            print(result)
            # Retorna o dicionário com as somas de cada campo
            return result
        
        return result
    
    def get_ttpecas(pasta_vendedor):
            # Caminho completo para o arquivo de texto
            caminho_arquivo = glob.glob(os.path.join(pasta_vendedor, "zTotal_Pecas*.txt"))

            if caminho_arquivo:
                try:
                    # Abre o arquivo e lê seu conteúdo
                    with open(caminho_arquivo[0], 'r', encoding='utf-8') as file:
                        texto = file.read()

                    # Verifica se o texto foi lido corretamente
                    if texto:
                        # Usando regex para buscar o padrão '__<número>__'
                        match = re.search(r'__(\d+)__', texto)
                        
                        if match:
                            return match.group(1)  # Retorna o número encontrado
                        else:
                            print("Padrão não encontrado no texto.")
                    else:
                        print("Arquivo vazio ou não leu o conteúdo corretamente.")
                except Exception as e:
                    print(f"Erro ao ler o arquivo: {e}")
            else:
                print("Nenhum arquivo encontrado com o padrão especificado.")
            return None
    
    try:
        pasta_vendedor = os.path.join(pasta_arquivo, "VENDEDOR")
        nome_projeto = obter_nome(pasta_arquivo)  # Suponha que essa função está retornando um nome válido
        
        aceite_data = {}
        
        op = get_op(pasta_arquivo)
        ripado, abrirem = get_totTiraRipado(pasta_vendedor)
        router = get_totPainelRouter(pasta_vendedor)
        engrosso, esp = get_totEngrosso(pasta_vendedor)
        usinagem = get_totFrente45(pasta_vendedor)
        result = get_totVarios(pasta_vendedor)
        ttpecas = get_ttpecas(pasta_vendedor)

        opzinha = str(op) if op is not None else ""
        
        aceite_data["projeto"] = str(nome_projeto) if nome_projeto is not None else ""
        aceite_data["opField"] = f"OP {opzinha}"
        aceite_data["ripado"] = str(ripado) if ripado is not None else "0"
        aceite_data["abrirem"] = str(abrirem) if abrirem is not None else ""
        aceite_data["router"] = str(router) if router is not None else "0"
        aceite_data["engrosso"] = str(engrosso) if engrosso is not None else "0"
        aceite_data["usinagem"] = str(usinagem) if usinagem is not None else "0"
        aceite_data["fitagem"] = str(result.get("fitagem", ""))
        aceite_data["espengrosso"] = str(esp) if esp is not None else ""
        aceite_data["furosist"] = str(result.get("furosist", ""))
        aceite_data["canal"] = str(result.get("canal", ""))
        aceite_data["furodob"] = str(result.get("furodob", ""))
        aceite_data["cortes"] = str(result.get("cortes", ""))
        aceite_data["corteperfil"] = str(result.get("cortePerfil", ""))
        aceite_data["tipoCanal"] = str(result.get("tipoCanal", ""))
        aceite_data["ttpecas"] = str(ttpecas) if ttpecas is not None else ""

        # Verifique se os dados estão corretos antes de gravar
        if aceite_data:
            caminho_json = os.path.join(pasta_arquivo, "VENDEDOR", f"aceite_OP-{opzinha}_{nome_projeto}.banana")
            
            # Verifique se a pasta de destino existe, caso contrário, crie
            pasta_destino = os.path.dirname(caminho_json)
            if not os.path.exists(pasta_destino):
                os.makedirs(pasta_destino)

            # Salva o arquivo JSON
            with open(caminho_json, "w") as f:
                json.dump(aceite_data, f, ensure_ascii=False, indent=4)

            # Mostra o caminho do arquivo gerado
            print(f"Arquivo JSON gerado em: {caminho_json}")
        else:
            print("Erro: aceite_data está vazio, não foi possível gerar o arquivo JSON.")
    
    except Exception as e:
        print(f"Erro ao gerar json: {e}")
###

def mostrar_erro(mensagem):
    root = Tk()
    root.withdraw()  # Oculta a janela principal
    messagebox.showerror("Erro", mensagem)
    root.destroy()  # Destroi a janela após mostrar o erro

def main(arquivo=None):
    print("  _______    ______   __    __   ______   __    __   ______   ")
    print(" |       \  /      \ |  \  |  \ /      \ |  \  |  \ /      \  ")
    print(" | $$$$$$$\|  $$$$$$\| $$\ | $$|  $$$$$$\| $$\ | $$|  $$$$$$\ ")
    print(" | $$__/ $$| $$__| $$| $$$\| $$| $$__| $$| $$$\| $$| $$__| $$ ")
    print(" | $$    $$| $$    $$| $$$$\ $$| $$    $$| $$$$\ $$| $$    $$ ")
    print(" | $$$$$$$\| $$$$$$$$| $$\$$ $$| $$$$$$$$| $$\$$ $$| $$$$$$$$ ")
    print(" | $$__/ $$| $$  | $$| $$ \$$$$| $$  | $$| $$ \$$$$| $$  | $$ ")
    print(" | $$    $$| $$  | $$| $$  \$$$| $$  | $$| $$  \$$$| $$  | $$ ")
    print("  \$$$$$$$  \$$   \$$ \$$   \$$ \$$   \$$ \$$   \$$ \$$   \$$ ")
    
    if not arquivo:
        # Abrir a janela de seleção de arquivo
        Tk().withdraw()  # Evitar que a janela principal do Tkinter apareça
        arquivo = askopenfilename(title="Selecione o arquivo Projeto_producao.xls", filetypes=[("Excel files", "*.xls;*.xlsx")])
    print(arquivo)
    if arquivo:
        try:
            # Ler o arquivo Excel
            df = pd.read_excel(arquivo)
            
            # Perguntar ao usuário qual relatório deseja gerar
            #print("Qual relatório você deseja gerar?")
            #print("1. Relatório de Peças")
            opcao = "1"  # input("Digite o número da opção: ")
            if opcao == '1':
                diretorio = os.path.dirname(arquivo)
                pasta_vendedor = os.path.join(diretorio, 'VENDEDOR')
                if not os.path.exists(pasta_vendedor):
                    os.makedirs(pasta_vendedor)
                
                nome = obter_nome(diretorio)
                gerar_relatorio_pecas(df, arquivo, nome)
                criar_arquivo_com_pecas(df, arquivo)
                arquivo_ripado(df, arquivo, nome)
            else:
                print("Opção inválida.")
        except Exception as e:
            mostrar_erro(f"Erro ao ler o arquivo ou gerar o relatório: {e}")

if __name__ == "__main__":
    main()
