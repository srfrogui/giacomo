import os
import shutil
import time
import tkinter as tk
import pyautogui as ag
from tkinter import filedialog, messagebox, IntVar
from tkinterdnd2 import TkinterDnD, DND_FILES
import xlrd
from xlutils.copy import copy
import logging
from PIL import Image, ImageTk
import threading
import PyPDF2
import win32gui
from fpdf import FPDF
import subprocess
import xml.etree.ElementTree as ET
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle

from contar_chapas import extrair_gplan_pdf, extrair_nesting_pdf

imagem={
    'salvar':'./img/salvar_salvar.png',
    'val_plan_cort':'./img/val_plan_cort.png',
    'val_explort_nc':'./img/val_export_nc.png',
    'abrir_nesting':'./img/abrir_nesting.png',
    'listas2':'./img/listras2.png',
    "importar": './img/btt_importar.png',
    "carregar_arquivo": './img/btt_carregar_arq.png',
    "otimizar": './img/btt_otimizar.png',
    "plano_corte": './img/btt_pla_cort.png',
    "converter": './img/btt_converter.png',
    "sel_lista": './img/btt_sel_lista.png',
    "exportar_plan_cort": './img/btt_exportar_plan_cort.png',
    "expor_nc_conf": './img/btt_exp_nc_conf.png',
    "pausa_carregar_arquivo":'./img/pausa_carregar_arquivo.png',
    "pausa_exportar_plan_corte":'./img/pausa_exportar_plano_corte.png',
    "listras":'./img/listras.png',
    "erro_etiqueta":'./img/val_erroetiquta.png',
    "view_etiqueta":'./img/val_etiqueta.png',
    "bt_etiqueta":'./img/btt_estiqueta.png',
}

# "resultado": './img/btt_resultado.png',
# "parametro": './img/btt_parametro.png',
# "maquina": './img/btt_maquina.png',
# "estatisticas": './img/btt_estatisticas.png',
# "configuracao": './img/btt_config.png',
# "carregar_pasta": './img/btt_carregar_pasta.png',
# "exportar_nc": './img/btt_exp_nc.png',
# "sel_pasta_pla_cort": './img/btt_sel_past_pla_cort.png',
# "ok_plan_cort": './img/btt_ok_plan_cort.png',
# "sel_past_nc": './img/btt_sel_past_nc.png',
# "excluir": './img/anc_excluir.png',
# "excluir_list": './img/btt_excluir_item_list.png',
# "pausa_exportar_nc":'./img/pausa_exportar_nc.png',
# "pausa_otimizando":'./img/pausa_otimizando.png',
# "pausa_por_otimizando":'./img/pausa_pos_otimizando.png',

def extrair_nome(caminho):
    # Extrai o nome da última pasta do caminho
    nome_pasta = os.path.basename(caminho)
    return nome_pasta

def on_drop(event):
    # Capturar o caminho da pasta arrastada
    folder_path = event.data
    folder_path = folder_path.replace("{", "").replace("}", "")
    print(f'Pasta solta: {folder_path}')
    
    # Extrair os caminhos das pastas usando split
    folder_path = folder_path.split("C:/")[1:]
    folder_path_list = ["C:/" + path.strip() for path in folder_path]
    # Adicionar as pastas extraídas à lista
    pastas.extend(folder_path_list)

    # Exibir as pastas na interface
    for pasta in folder_path_list:
        adicionar_pasta_interface(pasta)

def selecionar_pasta_manual():
    # Abre o diálogo para selecionar uma pasta manualmente
    pasta_selecionada = filedialog.askdirectory(title="Selecione uma pasta")
    if pasta_selecionada:
        # Adicionar a pasta manualmente
        pastas.append(pasta_selecionada)
        adicionar_pasta_interface(pasta_selecionada)
        
def adicionar_pasta_interface(pasta):
    if pasta not in pastas:
        pastas.append(pasta)
    update_pasta_view()
    
def remove_pasta(path):
    # Remove a pasta da lista e atualiza a visualização
    if path in pastas:
        pastas.remove(path)
        update_pasta_view()
    
def update_pasta_view():
    # Limpa o frame e recria as labels de pastas com função de clique para remoção
    for widget in frame_pastas.winfo_children():
        widget.destroy()
        
    # Definindo o número de colunas
    num_colunas = 2
    num_itens = len(pastas)

    for index, path in enumerate(pastas):
        nome_pasta = path.split('/')[-1]  # Apenas o nome da pasta
        column = index // (num_itens // num_colunas + 1)  # Define a coluna
        row = index % (num_itens // num_colunas + 1)  # Define a linha

        # Cria um label para a pasta
        label = tk.Label(frame_pastas, text=nome_pasta, font=("Arial", 10), relief=tk.RAISED, cursor="hand2")
        label.grid(row=row, column=column, padx=5, pady=5)  # Usando grid para disposição em colunas

        # Adiciona eventos de clique e hover
        label.bind("<Button-1>", lambda e, p=path: remove_pasta(p))
        label.bind("<Enter>", lambda event, p=path: caminho_label.config(text=p))  # Mostra caminho completo ao passar o mouse
        label.bind("<Leave>", lambda event: caminho_label.config(text=""))  # Limpa o texto ao sair

def aguarde(imagem, confianca=0.95, timeout=500, intervalo=2, inverter=False):
    timeO = 0
    acao = "sumir" if inverter else "aparecer"
    
    print(f"Esperando {acao} {imagem} .")

    while timeO < timeout:
        try:
            imagem_encontrada = ag.locateCenterOnScreen(imagem, confidence=confianca)
            if not inverter:
                print(f"Imagem {imagem} {'sumiu' if inverter else 'encontrada'}.")
                return True
            
        except Exception as e:
            if inverter:
                print(f"Imagem {imagem} sumiu.")
                return True
            
        timeO += 1
        
        print(f"Aguardando a imagem {imagem} {acao}: time {timeO}/{timeout}")
        time.sleep(intervalo)  # Aguarda o intervalo antes de tentar novament
        
    print(f"Timeout atingido. A imagem não {acao}.")
    raise ValueError(f"Timeout atingido. A imagem não {acao}.")

def procurar(imagem, confianca=0.98, limite=0.9):
    print(f"Procurando imagem... {imagem} - {os.path.exists(imagem)}")
    #print(f"Caminho relativo: {os.path.relpath(imagem)}")
    
    while confianca >= limite:
        try:
            localizacao = ag.locateCenterOnScreen(imagem, confidence=confianca)
            if localizacao:
                print(f"Imagem encontrada com confiança {confianca}")
                return localizacao
        except Exception as e:
            print(f"Erro ao procurar imagem com confiança {confianca}: {e}")
            confianca -= 0.01  # Diminui a confiança gradualmente
            
    print("Imagem não encontrada dentro do limite de confiança.")
    return None

def clicar(imagem, ajusteX=0, ajusteY=0, right=None):
    localizacao = procurar(imagem)

    if localizacao:  # Verifica se a imagem foi encontrada
        print(f"Imagem localizada em {localizacao}")
        x, y = localizacao.x, localizacao.y
        posicao_certa = (x + ajusteX, y + ajusteY)
        if right:
            ag.click(posicao_certa, button='right')
        else:
            ag.click(posicao_certa)
        print(f'Clique realizado na posição: {posicao_certa}')
    else:
        print(f"Imagem {imagem} não encontrada. Não foi possível clicar.")

def salvar(pasta=None, nome=None):
    aguarde(imagem['salvar'])
    if nome:
        ag.write(nome)
        print(nome)
    time.sleep(0.2)
    if pasta:
        ag.hotkey('ctrl', 'f')
        ag.hotkey(['shift', 'tab'] * 2)
        ag.press('enter')
        time.sleep(0.2)
        print(pasta)
        ag.write(pasta)
        ag.press('enter') #dentro da pasta
    time.sleep(1)
    clicar(imagem['salvar'], ajusteX=-100)
    
    
def log_message(message):
    if text_log:
        text_log.insert(tk.END, message + '\n')
        text_log.see(tk.END)


#===========================================================================

def obter_caminhos(pasta):
    vendedor = os.path.join(pasta, 'VENDEDOR')
    
    # Cria a pasta 'Nesting' dentro da pasta do cliente
    caminho_nesting = os.path.join(pasta, 'Nesting')
    os.makedirs(caminho_nesting, exist_ok=True)

    # Cria a pasta 'plano de corte' dentro da pasta 'Nesting'
    caminho_plano_corte = os.path.join(caminho_nesting, 'Plano de Corte')
    os.makedirs(caminho_plano_corte, exist_ok=True)
    
    # Exibe o conteúdo da pasta para debug
    print(f"Conteúdo da pasta {pasta}: {os.listdir(pasta)}")
    
    caminho_arquivo = None  # Inicializa a variável

    # Verifica todos os arquivos na pasta
    for arquivo in os.listdir(pasta):
        print(f"Analisando arquivo: {arquivo}")  # Para debug
        if arquivo.endswith(".xls") and "planoCorte_Moveo_Ecomobile_OP_" in arquivo:
            caminho_arquivo = os.path.join(pasta, arquivo)
            break  # Encontrei o arquivo, não preciso continuar o loop

    # Verifica se o arquivo foi encontrado
    if caminho_arquivo is None:
        raise ValueError("Cut nao encontrado")

    return caminho_arquivo, caminho_nesting, caminho_plano_corte, vendedor

def importar_optimiza(pasta, cut):
    cutut = os.path.basename(cut)
    aguarde(imagem['carregar_arquivo'])
    clicar(imagem['carregar_arquivo'])
    salvar(pasta, cutut)
    clicar(imagem["converter"])
    aguarde(imagem['pausa_carregar_arquivo'])
    clicar(imagem['sel_lista'])
    clicar(imagem['otimizar'])

def exportar_plano_corte(nesting, corte):
    aguarde(imagem['plano_corte'])
    time.sleep(2)
    aguarde(imagem['listras'])
    aguarde(imagem['listas2'], inverter=True)
    
    #verifica etiqueta carregada
    clicar(imagem['bt_etiqueta'])
    aguarde(imagem['view_etiqueta'])
    
    if procurar(imagem['erro_etiqueta']):
        ag.hotkey('alt','f4')
        mostrar_mensagem_erro('Imagem Etiqueta nao caregada!!!!')     
        return True
    else: 
        ag.hotkey('alt','f4')
        clicar(imagem['plano_corte'])
        print('1/4')
        time.sleep(1)
        if not procurar(imagem['val_plan_cort']):
            clicar(imagem['plano_corte'])
            print('2/4')
            time.sleep(1)
            if not procurar(imagem['val_plan_cort']):
                clicar(imagem['plano_corte'])
                print('3/4')
                time.sleep(1)
                if not procurar(imagem['val_plan_cort']):
                    clicar(imagem['plano_corte'])
                    print('4/4')
                    time.sleep(1)
                    if not procurar(imagem['val_plan_cort']):
                        time.sleep(1)
                        print('fds')
                        raise Exception('ovo')
        clicar(imagem['val_plan_cort'], ajusteX=430, ajusteY=350) #botao selecionar
        salvar(corte)
        clicar(imagem['val_plan_cort'], ajusteX=430, ajusteY=420) #botao exportar
        aguarde(imagem['pausa_exportar_plan_corte'])
        ag.press('enter')
        clicar(imagem['exportar_plan_cort'])
        clicar(imagem['expor_nc_conf'], ajusteY= -260) #botao selecionar 
        salvar(nesting)
        clicar(imagem['expor_nc_conf']) #botao exportar 
        aguarde(imagem['val_explort_nc']) #aguarde exportacao
        ag.hotkey('alt', 'f4') # fecha janela exportacao
        return False

def limpar_lista():
    clicar(imagem['importar'])
    aguarde(imagem['carregar_arquivo'])
    clicar(imagem['carregar_arquivo'], ajusteY=-100, right=True)
    clicar(imagem['carregar_arquivo'], ajusteY=-60, ajusteX=20)
    ag.press('enter')

def gerar_pdfs(pasta, vendedor):
    # Listar todas as imagens .bmp na pasta
    imagens = [arquivo for arquivo in os.listdir(pasta) if arquivo.endswith(".bmp")]
    
    # Organizar as imagens em ordem crescente
    imagens.sort(key=lambda x: (int(x.split('_')[0]), int(x.split('_')[1].split('.')[0])))

    # Agrupar imagens por material (primeira parte do nome)
    grupos = {}
    for imagem in imagens:
        material = imagem.split('_')[0]
        if material not in grupos:
            grupos[material] = []
        grupos[material].append(imagem)

    # Criar PDFs para cada grupo
    for material, arquivos in grupos.items():
        pdf = FPDF('P', 'mm', 'A4')
        pdf.set_auto_page_break(auto=True, margin=10)
        largura, altura = 210, 297  # Tamanho da página A4 em mm

        for i in range(0, len(arquivos), 2):  # Processar 4 imagens por iteração (2 páginas)
            pdf.add_page()
            
            # Adicionar primeira imagem na metade superior da página
            if i < len(arquivos):
                img1_path = os.path.join(pasta, arquivos[i])
                img1 = Image.open(img1_path)
                img1.thumbnail((int(largura * 3.78), int(altura * 1.89)))  # Ajustar DPI para a metade superior
                pdf.image(img1_path, x=10, y=10, w=190)

            # Adicionar segunda imagem na metade inferior da página
            if i + 1 < len(arquivos):
                img2_path = os.path.join(pasta, arquivos[i + 1])
                img2 = Image.open(img2_path)
                img2.thumbnail((int(largura * 3.78), int(altura * 1.89)))  # Ajustar DPI para a metade inferior
                pdf.image(img2_path, x=10, y=150, w=190)

        # Salvar o PDF
        pdf_path = os.path.join(vendedor, f"yG2Nesting_{material}.pdf")
        pdf.output(pdf_path)
        print(f"PDF criado: {pdf_path}")

def obter_nome(pasta_arquivo):
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

def compress_to_rar(vendedor, cliente):
    
    rar_filename = os.path.join(os.path.dirname(vendedor), f'VENDEDOR_{cliente}.rar')
    
    # Chama o WinRAR via linha de comando
    rar_path = os.path.join(os.environ['ProgramFiles'], 'WinRAR', 'Rar.exe')
    command = [rar_path, 'a', '-ep1', rar_filename, vendedor] # Adicionando '-ep1'
    try:
        subprocess.run(command, check=True)
        print(f"Arquivos RAR criados: {rar_filename}")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao criar arquivo RAR: {e}") 

def gerar_relatorio_pdf1(corte, vendedor, pasta):
    caminho_pdf = vendedor + '/xRelatorio_Chapas.pdf'
    caminho_xml = os.path.join(os.path.dirname(corte), 'InfoOutput.xml')
    pastanome =  os.path.basename(pasta)
    
    # Parse do XML
    tree = ET.parse(caminho_xml)
    root = tree.getroot()

    # Configurar o documento PDF com margem superior reduzida
    doc = SimpleDocTemplate(caminho_pdf, pagesize=letter, topMargin=20)  # topMargin ajustado para 20
    story = []

    # Definir estilos para o conteúdo
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CustomTitle', fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor("#333333"), spaceAfter=12))
    styles.add(ParagraphStyle(name='Subtitle', fontName='Helvetica', fontSize=14, textColor=colors.HexColor("#666666"), spaceAfter=10))
    styles.add(ParagraphStyle(name='CustomNormal', fontName='Helvetica', fontSize=12, textColor=colors.black, spaceAfter=6))

    # Título do Relatório
    story.append(Paragraph(f"Relatório {pastanome}", styles['CustomTitle']))

    # Adicionando os Arquivos NcFiles
    story.append(Paragraph("Arquivos NcFiles:", styles['Subtitle']))
    
    nc_files_data = []
    #nc_files_data.append(["Arquivo"])
    
    for file in root.findall(".//NcFiles/File"):
        file_name = file.get("name")
        nc_files_data.append([file_name])

    # Criar tabela de arquivos NcFiles
    nc_table = Table(nc_files_data, colWidths=[500], rowHeights=30)
    nc_table.setStyle(TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#333333")),  # Cor do texto do título da coluna
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento do texto
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f2f2f2")),  # Cor de fundo da linha de cabeçalho
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#000")),  # Bordas das células
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#f9f9f9")]),  # Cores alternadas nas linhas
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 18),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')  # Alinha o texto para o topo das células
    ]))
    story.append(nc_table)

    # # Adicionando o arquivo de NestingOutputJsonFile
    # nesting_output = root.find(".//NestingOutputJsonFile")
    # if nesting_output is not None and nesting_output.get("name"):
    #     story.append(Paragraph(f"Nesting Output Json File: {nesting_output.get('name')}", styles['CustomNormal']))
    # else:
    #     story.append(Paragraph("Nesting Output Json File: Não especificado", styles['CustomNormal']))

    # Adicionando os Arquivos InputSourceFiles
    story.append(Paragraph("Arquivos InputSourceFiles:", styles['Subtitle']))
    
    input_files_data = []
    #input_files_data.append(["Arquivo"])

    for file in root.findall(".//InputSourceFiles/File"):
        file_name = file.get("name")
        input_files_data.append([file_name])

    # Criar tabela de arquivos InputSourceFiles
    input_table = Table(input_files_data, colWidths=[500], rowHeights=30)
    input_table.setStyle(TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#333333")),  # Cor do texto do título da coluna
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento do texto
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f2f2f2")),  # Cor de fundo da linha de cabeçalho
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#000")),  # Bordas das células
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#f9f9f9")]),  # Cores alternadas nas linhas
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 18),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')  # Alinha o texto para o topo das células
    ]))
    story.append(input_table)

    # Gerar o PDF
    doc.build(story)
    print("Relatório gerado com sucesso!")
    
def gerar_relatorio_pdf(corte, vendedor, pasta):
    # Caminhos de arquivos
    caminho_xml = os.path.join(os.path.dirname(corte), 'InfoOutput.xml')
    pastanome = os.path.basename(pasta)
    # Parse do XML
    tree = ET.parse(caminho_xml)
    root = tree.getroot()

    story = []
    # Definir estilos para o conteúdo
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CustomTitle', fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor("#333333"), spaceAfter=12))
    styles.add(ParagraphStyle(name='Subtitle', fontName='Helvetica', fontSize=14, textColor=colors.HexColor("#666666"), spaceAfter=10))
    styles.add(ParagraphStyle(name='CustomNormal', fontName='Helvetica', fontSize=12, textColor=colors.black, spaceAfter=6))


    nc_files_data = []
    for file in root.findall(".//NcFiles/File"):
        file_name = file.get("name")
        nc_files_data.append([file_name])
    
    input_files_data = []
    for file in root.findall(".//InputSourceFiles/File"):
        file_name = file.get("name")
        input_files_data.append([file_name])

    # Extrair dados de nesting e gplan
    try:
        resultado_nesting = extrair_nesting_pdf(pasta, nc_files_data)
    except Exception as e:
        print('Erro ao extrair dados de Nesting:', e)
        resultado_nesting = {}

    try:
        resultado_gplan = extrair_gplan_pdf(vendedor)
    except Exception as e:
        print('Erro ao extrair dados de GPlan:', e)
        resultado_gplan = {}

    # Organizar as chaves em ordem alfabética e garantir que o 'total' fique no final
    resultado_nesting = {k: resultado_nesting[k] for k in sorted(resultado_nesting) if k != 'total'}
    resultado_gplan = {k: resultado_gplan[k] for k in sorted(resultado_gplan) if k != 'total'}

    # Calcular os totais para cada dicionário
    total_nesting = sum(resultado_nesting.values())
    total_gplan = sum(resultado_gplan.values())

    # Preparar os dados para gerar a tabela no formato de texto
    todas_chaves = sorted(set(resultado_nesting.keys()).union(set(resultado_gplan.keys())))

    # Criar o objeto doc e a lista story
    caminho_arquivo_pdf = os.path.join(vendedor, "wRelatorio_Chapas.pdf")
    doc = SimpleDocTemplate(caminho_arquivo_pdf, pagesize=letter, topMargin=20)

    # Título do Relatório
    story.append(Paragraph(f"Relatório {pastanome}", styles['CustomTitle']))

    # Adicionando a tabela de Nesting/GPlan
    story.append(Paragraph("Contagem de Chapas Nesting/GPlan:", styles['Subtitle']))

    table_data = []
    # Cabeçalho da tabela
    table_data.append(["Produto", "NESTING", "GPLAN"])

    # Adicionando os dados das chapas com linha alternada
    for idx, chave in enumerate(todas_chaves):
        quantidade_nesting = resultado_nesting.get(chave, 0)
        quantidade_gplan = resultado_gplan.get(chave, 0)

        # Adicionar a linha com a cor de fundo definida
        table_data.append([chave, str(quantidade_nesting), str(quantidade_gplan)])

    # Adicionar o total ao final
    table_data.append(["TOTAL:", str(total_nesting), str(total_gplan)])

    # Criar a tabela de Nesting/GPlan
    table = Table(table_data, colWidths=[300, 50, 50], rowHeights=13)

    # Estilizando a tabela
    table.setStyle(TableStyle([ 
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#333333")),  # Cor do texto
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alinhamento centralizado
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f2f2f2")),  # Cor de fundo do cabeçalho
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#000")),  # Grid de separação
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#f9f9f9")]),  # Linhas alternadas de fundo
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),  # Fonte
        ('FONTSIZE', (0, 0), (-1, -1), 10),  # Tamanho da fonte
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Alinhamento vertical do texto
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),  # Reduzir padding inferior se necessário
        ('TOPPADDING', (0, 0), (-1, -1), 0),  # Reduzir padding inferior se necessário
    ]))

    # Aplicando a cor de fundo dinâmica para as linhas
    for i, row in enumerate(table_data[1:-1]):  # Ignorando o cabeçalho
        if float(row[1]) != float(row[2]):  # Se os valores de NESTING e GPLAN não forem iguais
            # Alterando a cor de fundo da linha para as divergentes
            if float(row[1]) > float(row[2]):
                table.setStyle(TableStyle([('BACKGROUND', (1, i + 1), (1, i + 1), colors.lightpink)])) # Cor de fundo para divergência da segunda coluna
            else:
                table.setStyle(TableStyle([('BACKGROUND', (2, i + 1), (2, i + 1), colors.lightpink)])) # Cor de fundo para divergência da terceira coluna
                
    # Alterar a cor da última linha (total) para um cinza mais escuro
    last_row_index = len(table_data) - 1  # Índice da última linha (total)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, last_row_index), (-1, last_row_index), colors.lightgrey)  # Cor de fundo cinza mais escuro
    ]))

    # Adicionar a tabela no story
    story.append(table)

        # Adicionando os Arquivos NcFiles
    story.append(Paragraph("Arquivos NcFiles:", styles['Subtitle']))

    # Criar tabela de arquivos NcFiles
    nc_table = Table(nc_files_data, colWidths=[500], rowHeights=25)
    nc_table.setStyle(TableStyle([ 
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#333333")),  
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f2f2f2")),  
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#000")),  
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#f9f9f9")]),  
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 15),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')  
    ]))
    story.append(nc_table)

    # Adicionando os Arquivos InputSourceFiles
    story.append(Paragraph("Arquivos InputSourceFiles:", styles['Subtitle']))

    # Criar tabela de arquivos InputSourceFiles
    input_table = Table(input_files_data, colWidths=[500], rowHeights=20)
    input_table.setStyle(TableStyle([ 
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#333333")),  
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f2f2f2")),  
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#000")),  
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#f9f9f9")]),  
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')  
    ]))
    story.append(input_table)

    # Gerar o PDF
    doc.build(story)
    print("Relatório gerado com sucesso!")
    
    
log_file_N = 'Nesting.log'

def processo_nesting(pasta):
    with open(log_file_N, 'a') as log:
        log.write(f'Processando pasta: {pasta}\n')
        
    cut, nesting, corte, vendedor = obter_caminhos(pasta)
    cliente = obter_nome(pasta)
    
    # Etapas do processamento
    if var_importa.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Importando e optimizando...\n')
        log_message("Importando e optimizando...")
        importar_optimiza(pasta, cut)
    
    if var_exportar.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Exportando planos de corte...\n')
        log_message("Exportando planos de corte...")
        erro = exportar_plano_corte(nesting, corte)
    if not var_exportar.get():
        erro = None
    if var_limpar_lista.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Limpando lista...')
        log_message("Limpando lista...")
        limpar_lista()
        
    if var_relatorio_pdf.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Gerando PDFs...')
        if erro:
            print('Naogerar, erro encontrado sem etiqueta')
            log_message('Naogerar, erro encontrado sem etiqueta')
        else:
            log_message("Gerando PDFs...")
            gerar_pdfs(corte, vendedor)
        
    if gerar_pdf_html.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Gerando relatorio InfoOutput...')
        if erro:
            print('Naogerar, erro encontrado sem etiqueta')
            log_message('Naogerar, erro encontrado sem etiqueta')
        else:
            log_message("Gerando relatorio InfoOutput...")
            gerar_relatorio_pdf(corte, vendedor, pasta)
    
    if compress_vend.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Compactando VENDEDOR...')
        log_message("Compactando VENDEDOR...")
        compress_to_rar(vendedor, cliente)
    
    if var_del_lista.get():
        try:
            if erro:
                print('Mantendo na lista projeto com erro')
            else:
                remove_pasta(pasta)
        except: #hehe 
            remove_pasta(pasta)

    with open(log_file_N, 'a') as log:
        log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
        log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')
            
def atualizar_frame_pastas():
    # Clear the frame content
    for widget in frame_pastas.winfo_children():
        widget.destroy()

    # Re-add each remaining folder to the frame
    for pasta in pastas:
        tk.Label(frame_pastas, text=pasta, fg="blue").pack(anchor='w', padx=5)
        
def mostrar_mensagem_erro(mensagem):
    # Show a Tkinter window to notify the user and wait for their response
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Atenção", f"{mensagem}. Clique OK para continuar.")
    root.destroy()

def main():
    def ok():
        if not pastas :
            messagebox.showinfo("Atenção", "Selecione uma pasta.")
            return    
        
        for pasta in pastas:       
            try:
                
                # processo_nesting(pasta)
                threading.Thread(target=processo_nesting, args=(pasta,)).start()
                log_message("Processo Nesting Iniciado...")


                    
            except ValueError as e:  # Captura especificamente o erro lançado na função clicar
                with open(log_file_N, 'a') as log:
                    log.write(f'Erro ao processar pasta {pasta}: {e}\n')
                    log_message(f'Erro ao processar pasta {pasta}: {e}\n')
                    mostrar_mensagem_erro("Aviso: Erro ao processar pasta {pasta}: {e}\n")
                continue
            except Exception as e:
                with open(log_file_N, 'a') as log:
                    log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                    mostrar_mensagem_erro("Aviso: Erro generico ao processar pasta {pasta}: {e}\n")
                break
            
    def on_close():
        janela.destroy()
        os._exit(0)   

    global caminho_label, frame_pastas, pastas, text_log
    pastas=[]
    
    # Criar janela principal
    global janela
    janela = TkinterDnD.Tk()
    janela.title("Seleção de Pastas")
    janela.geometry("500x400")  
    
    janela.protocol("WM_DELETE_WINDOW", on_close)

    # Frame para exibir as pastas
    frame_pastas = tk.Frame(janela)
    frame_pastas.pack(pady=20, padx=10)

    # Label para exibir o caminho completo
    caminho_label = tk.Label(janela, text="", fg="blue")
    caminho_label.pack(pady=5)

    global var_importa, var_exportar, var_limpar_lista, var_relatorio_pdf, gerar_pdf_html, compress_vend, var_del_lista
    # Variáveis para checkboxes (definidas após criar a janela)
    var_importa = IntVar(value=1)
    var_exportar= IntVar(value=1)
    var_limpar_lista = IntVar(value=1)
    var_relatorio_pdf = IntVar(value=1)
    gerar_pdf_html = IntVar(value=1)
    compress_vend = IntVar(value=1)
    var_del_lista = IntVar(value=0)
    
    # Frame para organizar os checkboxes horizontalmente
    checkbox_frame = tk.Frame(janela)
    checkbox_frame.pack(pady=3)

    # Checkboxes adicionados ao frame com layout grid
    tk.Checkbutton(checkbox_frame, text="Importa", variable=var_importa).grid(row=0, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Exportar", variable=var_exportar).grid(row=0, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Limpar Lista", variable=var_limpar_lista).grid(row=0, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Materiais PCorte", variable=var_relatorio_pdf).grid(row=1, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Relatorio PCorte", variable=gerar_pdf_html).grid(row=1, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Comp Vendedor", variable=compress_vend).grid(row=1, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="*Remover Lista", variable=var_del_lista).grid(row=1, column=3, sticky='w')

    # Frame para organizar os botões
    frame_botoes = tk.Frame(janela)
    frame_botoes.pack(pady=10)
    
    # Botão para seleção manual de pastas
    botao_selecionar = tk.Button(frame_botoes, text="Selecionar Pasta Manualmente", command=selecionar_pasta_manual)
    botao_selecionar.pack(side=tk.LEFT, padx=10)

    # Botão OK para finalizar
    botao_ok = tk.Button(frame_botoes, text="OK", command=ok)
    botao_ok.pack(side=tk.LEFT, padx=10)
    
    text_log = tk.Text(janela, height=200, width=80)
    text_log.pack(pady=10, padx=10)
    
    # Configuração de drag-and-drop
    janela.drop_target_register(DND_FILES)
    janela.dnd_bind('<<Drop>>', on_drop)

    # Iniciar a interface
    janela.mainloop()
    
if __name__ == '__main__': 
    main()
    
#  | |__     __ _    __ _    __ _ 
#  | '_ \   / _` |  / _` |  / _` | 
#  | |_) | | (_| | | (_| | | (_| |
#  |_.__/   \__,_|  \__,_|  \__,_| 
    
# ░░▄█▀▄█▀▀▀▀▄░░░░░░▄▀▀█▄░▀█▄░░█▄░░░▀█░░░░
# ░▄█░▄▀░░▄▄▄░█░░░▄▀▄█▄░▀█░░█▄░░▀█░░░░█░░░
# ▄█░░█░░░▀▀▀░█░░▄█░▀▀▀░░█░░░█▄░░█░░░░█░░░
# ██░░░▀▄░░░▄█▀░░░▀▄▄▄▄▄█▀░░░▀█░░█▄░░░█░░░