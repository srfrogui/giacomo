from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from PIL import Image
from tkinter import Tk, Label, Button, Entry, filedialog, StringVar, ttk, messagebox
import os
import pandas as pd
import win32com.client
import glob

class GeradorPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de PDF de Projeto Produção")

        # Variáveis para armazenar entrada do usuário
        self.pasta_projeto = StringVar()
        self.coluna_filtro = StringVar()
        self.valores_filtro = StringVar()

        # Layout da interface
        Label(root, text="Selecione a pasta do projeto:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_pasta = Entry(root, textvariable=self.pasta_projeto, width=50)
        self.entry_pasta.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        Button(root, text="Selecionar", command=self.selecionar_pasta).grid(row=0, column=2, padx=10, pady=5)

        Label(root, text="Selecione a coluna para filtrar:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.combo_colunas = ttk.Combobox(root, textvariable=self.coluna_filtro, values=[
            "pcpitem", "desenho", "ambiente", "localizador", "componente", "modulo",
            "esp_material", "comprimento", "largura", "veio_material", "pcpped","cod_material"
        ])
        self.combo_colunas.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        Label(root, text="Digite os valores para filtrar (separados por vírgula):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_valores = Entry(root, textvariable=self.valores_filtro, width=50)
        self.entry_valores.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        Button(root, text="Gerar PDF", command=self.gerar_pdf).grid(row=3, column=1, padx=10, pady=20)

    def selecionar_pasta(self):
        """Seleciona a pasta do projeto."""
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_projeto.set(pasta)

    def gerar_pdf(self):
        # Validar entradas
        pasta = self.pasta_projeto.get()
        coluna = self.coluna_filtro.get()
        valores = self.valores_filtro.get()
        
        
        """CRIA ARQUIVO LIVEL"""
        caminho_ela = glob.glob(os.path.join(pasta, "planoCorte_Moveo_Ecomobile_OP_*.xls"))
        
        if not caminho_ela:
            print("Erro: Nenhum arquivo encontrado com o padrão especificado.")
            return  # Retorna para evitar falhas posteriores
        
        caminho_el = caminho_ela[0]
        caminho_excel = os.path.join(pasta, "Cut_Livel.xls")
        caminho_ex = os.path.normpath(caminho_excel)
        try:
            os.remove(caminho_excel)
        except Exception:
            pass
        print(caminho_ex)
        try:
            # Iniciar o Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Tornar invisível durante a execução

            # Imprimir o caminho do arquivo encontrado (para depuração)
            print(f"Abrindo arquivo: {caminho_el}")

            # Abrir o arquivo
            workbook = excel.Workbooks.Open(caminho_el)
            if not workbook:
                print("Erro: Não foi possível abrir o arquivo Excel.")
                return

            # Salvar como .xls (Excel 97-2003)
            workbook.SaveAs(caminho_ex, FileFormat=56)  # Salva como .xls (Excel 97-2003)
            print(f"Arquivo salvo com sucesso em: {caminho_ex}")

            # Fechar o arquivo e o Excel
            workbook.Close(SaveChanges=False)
            excel.Quit()  # Certifique-se de que o Excel será fechado corretamente
        
            """Lê os dados do Excel, aplica os filtros e gera o PDF."""
            try:

                if not pasta or not coluna or not valores:
                    messagebox.showerror("Erro", "Preencha todos os campos antes de gerar o PDF.")
                    return

                # Caminho do arquivo Excel
                caminho_imagens = os.path.join(pasta, "Gplan")
                if not os.path.exists(caminho_excel):
                    messagebox.showerror("Erro", f"Arquivo não encontrado: {caminho_excel}")
                    return

                # Ler o arquivo Excel
                df = pd.read_excel(caminho_excel)

                os.remove(caminho_excel)
                
                # Verificar se a coluna existe no DataFrame
                if coluna not in df.columns:
                    messagebox.showerror("Erro", f"A coluna '{coluna}' não existe no arquivo.")
                    return

                # Aplicar filtro
                valores_filtrados = [v.strip() for v in valores.split(",")]
                
                # Extrair somente o pecaID (parte após a barra) para comparação
                df['PECA ID'] = df['localizador'].astype(str).str.split('/').str[-1]
                
                # Aplicar filtro adequado com base na coluna
                if coluna == "localizador":
                    # Filtro que verifica se qualquer parte do valor corresponde
                    df_filtrado = df[df[coluna].str.contains('|'.join(valores_filtrados), na=False)]
                else:
                    # Filtro para verificar correspondência exata
                    df_filtrado = df[df[coluna].astype(str).isin(valores_filtrados)]

                if df_filtrado.empty:
                    messagebox.showinfo("Aviso", "Nenhum dado encontrado com os filtros aplicados.")
                    return

                # Gerar PDF
                nome_pdf = os.path.join(pasta, "Itens_Faltantes.pdf")
                self.criar_pdf(df_filtrado, nome_pdf, caminho_imagens)

                messagebox.showinfo("Sucesso", f"PDF gerado com sucesso: {nome_pdf}")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
                
        except Exception as e:
            print(f"Erro ao processar o arquivo Excel: {e}")

    def criar_pdf(self, df_filtrado, nome_pdf, caminho_imagens):
            """Cria o PDF a partir do DataFrame filtrado."""
            # Dicionário de mapeamento de colunas
            mapeamento_colunas = {
                "cliente": "CLIENTE",
                "desc_material": "MATERIAL",
                "esp_material": "ESP",
                "quantidade": "QTN",
                "comprimento": "ALTURA (X)",
                "largura": "PROF (Y)",
                "pcpitem": "PCP",
                "desenho": "DESENHO",
                "componente": "PEÇA DESCRIÇÃO",
                "furacao_f1": "PROG. 1",
                "furacao_f2": "PROG. 2",    
                "ambiente":"AMBIENTE",
                "cod_material":"COD",
            }

            # Aplica o mapeamento das colunas ao DataFrame
            df_filtrado = df_filtrado.rename(columns=mapeamento_colunas)

            # Criação do objeto canvas para o PDF
            c = canvas.Canvas(nome_pdf, pagesize=landscape(letter))  # Definir página em paisagem
            largura, altura = landscape(letter)  # A4 paisagem em pontos (792x612)

            # Definindo o tamanho das colunas e altura das linhas
            largura_coluna = 55  # Ajuste da largura das colunas
            altura_linha = 18    # Ajuste da altura das linhas
            inicio_x = 20
            inicio_y = altura - 30
            y_pos = inicio_y 

            # Títulos das colunas
            colunas_interesse = [
                "PCP", "CLIENTE", "MATERIAL", "ESP", "QTN",
                "ALTURA (X)", "PROF (Y)", "DESENHO",
                "PEÇA DESCRIÇÃO", "PROG. 1", "PROG. 2", "PECA ID", "COD"
            ]
            
            # Preparando os dados da tabela (extração do DataFrame)
            dados_tabela = []
            dados_tabela.append(colunas_interesse)  # Adiciona os cabeçalhos

            # Substitui NaN por uma string vazia
            df_filtrado = df_filtrado.fillna("")

            # Limitar o campo "CLIENTE" a no máximo 18 caracteres
            df_filtrado["AMBIENTE"] = df_filtrado["AMBIENTE"].str[:10]
            df_filtrado["CLIENTE"] = df_filtrado["CLIENTE"].str[:18]
            
            # Adicionar os dados filtrados
            for _, row in df_filtrado.iterrows():
                linha = [str(row.get(coluna, "")) for coluna in colunas_interesse]
                dados_tabela.append(linha)
                
            # Define a altura mínima para criar uma nova página
            altura_minima = 10  

            # Cabeçalho da tabela
            dados_tabela = []
            dados_tabela.append(colunas_interesse)  # Adiciona os cabeçalhos

            # Define largura das colunas como antes
            colWidths = [largura_coluna] * len(colunas_interesse)
            colWidths[0] = 30
            colWidths[1] = 100
            colWidths[2] = 150  # Aumenta a largura da segunda coluna (MATERIAL)
            colWidths[3] = 20
            colWidths[4] = 20
            colWidths[6] = 45
            colWidths[8] = 65
            colWidths[9] = 60
            colWidths[10] = 50
            colWidths[11] = 50

            # Adicionar cabeçalho ao PDF
            tabela_cabecalho = Table([colunas_interesse], colWidths=colWidths, rowHeights=altura_linha)
            tabela_cabecalho.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 0), (-1, 0), 8)
            ]))

            # Escrever cabeçalho na página
            tabela_cabecalho.wrapOn(c, largura, altura)
            tabela_cabecalho.drawOn(c, inicio_x, y_pos)
            y_pos -= altura_linha

            # Gerar dados linha por linha
            for _, row in df_filtrado.iterrows():
                linha = [str(row.get(coluna, "")) for coluna in colunas_interesse]
                
                # Criar uma tabela apenas para a linha atual
                tabela_linha = Table([linha], colWidths=colWidths, rowHeights=altura_linha)
                tabela_linha.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('FONTSIZE', (0, 0), (-1, -1), 7)
                ]))
                
                # Verificar se a próxima linha cabe na página
                if y_pos - altura_linha < altura_minima:
                    c.showPage()  # Criar nova página
                    y_pos = altura - 30  # Reiniciar posição Y
                    # Adicionar o cabeçalho novamente
                    tabela_cabecalho.wrapOn(c, largura, altura)
                    tabela_cabecalho.drawOn(c, inicio_x, y_pos)
                    y_pos -= altura_linha  # Atualizar posição Y após cabeçalho
                
                # Desenhar a linha no PDF
                tabela_linha.wrapOn(c, largura, altura)
                tabela_linha.drawOn(c, inicio_x, y_pos)
                y_pos -= altura_linha  # Atualizar posição Y após desenhar a linha
                
            y_pos -= altura_linha - 36
            # Definir o número máximo de imagens por linha (5 imagens)
            max_images_per_row = 5

            # Definir as dimensões da imagem e do quadrado
            image_width = 130
            image_height = 130
            square_width = image_width + 20
            square_height = image_height + 30  # Inclui o espaço para o texto
            
            afastamento_esquerda= 13
            # Inicializar as posições
            x_pos_imagem = inicio_x + afastamento_esquerda  # Posição inicial X para as imagens
            y_pos_imagem = y_pos - 150  # Posição inicial Y para as imagens (abaixo da tabela)
            counter = 0  # Contador para o número de imagens na linha

            # Começar o loop para inserir as imagens
            for i, row in df_filtrado.iterrows():
                imagem_nome = row["DESENHO"]
                imagem_id = row["PECA ID"]
                imagem_path = os.path.join(caminho_imagens, f"{imagem_nome}.bmp")
                
                if os.path.exists(imagem_path):
                    # Desenhar o quadrado ao redor da imagem e do texto
                    c.setLineWidth(1)
                    c.rect(x_pos_imagem - 10, y_pos_imagem - 10, square_width, square_height, stroke=1, fill=0)

                    # Desenhar o nome da imagem e o número dentro do quadrado
                    c.setFont("Helvetica", 12)
                    c.drawString(x_pos_imagem, y_pos_imagem + image_height + 5, f"{imagem_nome} - {imagem_id}")

                    # Desenhar a imagem no PDF
                    c.drawImage(imagem_path, x_pos_imagem, y_pos_imagem, width=image_width, height=image_height)

                    # Atualiza a posição para a próxima imagem
                    x_pos_imagem += square_width  # Avança para a próxima posição X
                    counter += 1  # Incrementa o contador de imagens na linha

                    # Se atingiu o número máximo de imagens por linha, ajusta para a próxima linha
                    if counter == max_images_per_row:
                        x_pos_imagem = inicio_x + afastamento_esquerda  # Volta para a posição inicial da linha
                        y_pos_imagem -= square_height  # Avança para a próxima linha de imagens
                        counter = 0  # Reseta o contador

                    # Verificar se a posição Y para as imagens excedeu a página
                    if y_pos_imagem < 40:
                        c.showPage()  # Adicionar uma nova página
                        y_pos_imagem = altura - 200  # Reiniciar a posição Y para as imagens
                        x_pos_imagem = inicio_x + afastamento_esquerda  # Reiniciar a posição X para o início da página
                        counter = 0  # Reseta o contador quando uma nova página é criada

            # Salvar o PDF
            c.save()

if __name__ == "__main__":
    root = Tk()
    app = GeradorPDFApp(root)
    root.mainloop()
