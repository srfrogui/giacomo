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

def gerar_relatorio_pecas(df, arquivo_excel):   
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
    
    print('pasta_arquivo', pasta_arquivo)
    nome = obter_nome(pasta_arquivo)
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
    # Obter o nome do projeto
    nome_projeto = obter_nome(pasta_arquivo)
    
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
    
### funcao p importar  
import pdfplumber
import glob
import json
import re

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
            return len(ripado)  # Retorna o número de ocorrências encontradas
        return None
            
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
                { 'regex': r'(\d+)\s*(ML|M2|UN)\s*(usi_rebaixo_4mm|usi_rasgo_7mm)', 'field': 'canal' },
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

            # Flag para identificar se "usi_rasgo_7mm" ou "usi_rebaixo_4mm" aparecem
            rasgo_found = False
            canal_led_found = False

            # Itera sobre cada regexData e faz a contagem para cada campo
            for item in regexData:
                matches = re.findall(item['regex'], texto, flags=re.IGNORECASE)
                
                # Para cada campo, soma os valores encontrados
                for match in matches:
                    # A quantidade que deve ser somada é o primeiro grupo (o número encontrado)
                    result[item['field']] += int(match[0])  # match[0] é o número encontrado
                    
                    # Verifica se as condições para tipoCanal são atendidas
                    if item['field'] == 'canal':
                        if 'usi_rasgo_7mm' in match[2]:
                            rasgo_found = True
                        if 'usi_rebaixo_4mm' in match[2]:
                            canal_led_found = True

            # Define o valor de tipoCanal com base nas condições
            if rasgo_found and canal_led_found:
                result['tipoCanal'] = 'RASGO E CANAL LED'
            elif rasgo_found:
                result['tipoCanal'] = 'RASGO'
            elif canal_led_found:
                result['tipoCanal'] = 'CANAL LED'
            
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
        ripado = get_totTiraRipado(pasta_vendedor)
        router = get_totPainelRouter(pasta_vendedor)
        engrosso, esp = get_totEngrosso(pasta_vendedor)
        usinagem = get_totFrente45(pasta_vendedor)
        result = get_totVarios(pasta_vendedor)
        ttpecas = get_ttpecas(pasta_vendedor)

        opzinha = str(op) if op is not None else ""
        
        aceite_data["opField"] = f"OP {opzinha}"
        aceite_data["ripado"] = str(ripado) if ripado is not None else ""
        aceite_data["router"] = str(router) if router is not None else ""
        aceite_data["engrosso"] = str(engrosso) if engrosso is not None else ""
        aceite_data["usinagem"] = str(usinagem) if usinagem is not None else ""
        aceite_data["fitagem"] = str(result.get("fitagem", ""))
        aceite_data["espengrosso"] = str(esp) if esp is not None else ""
        aceite_data["furosist"] = str(result.get("furosist", ""))
        aceite_data["canal"] = str(result.get("canal", ""))
        aceite_data["furodob"] = str(result.get("furodob", ""))
        aceite_data["cortes"] = str(result.get("cortes", ""))
        aceite_data["corteperfil"] = str(result.get("cortePerfil", ""))
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
                gerar_relatorio_pecas(df, arquivo)
                criar_arquivo_com_pecas(df, arquivo)
            else:
                print("Opção inválida.")
        except Exception as e:
            mostrar_erro(f"Erro ao ler o arquivo ou gerar o relatório: {e}")

if __name__ == "__main__":
    main()
