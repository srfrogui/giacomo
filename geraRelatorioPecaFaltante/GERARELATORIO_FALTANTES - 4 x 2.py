import time
import webbrowser
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
import barcode
from barcode.writer import ImageWriter
import io
import tempfile

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
            "pcpitem", "desenho", "ambiente", "localizador", "componente",
            "esp_material", "comprimento", "largura", "veio_material", "pcpped","cod_material"
        ])
        self.combo_colunas.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        Label(root, text="Digite os valores para filtrar (separados por vírgula):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_valores = Entry(root, textvariable=self.valores_filtro, width=50)
        self.entry_valores.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        Button(root, text="Gerar PDF", command=self.gerar_pdf).grid(row=3, column=1, padx=10, pady=20)
        Button(root, text="Revelar no Explorer", command=self.abrir_pasta_pdf).grid(row=4, column=1, padx=10, pady=5, sticky="w")

    def abrir_pasta_pdf(self):
        if hasattr(self, 'caminho_pdf_gerado') and os.path.exists(self.caminho_pdf_gerado):
            pasta_pdf = os.path.dirname(self.caminho_pdf_gerado)
            os.startfile(pasta_pdf)  # Só funciona no Windows
        else:
            messagebox.showwarning("Aviso", "Nenhum PDF foi gerado ainda.")

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

                nome_projeto = os.path.basename(pasta)
                # Gerar PDF
                nome_pdf = os.path.join(pasta, f"Itens_Faltantes_{nome_projeto}_{int(time.time())}.pdf")
                self.criar_pdf(df_filtrado, nome_pdf, caminho_imagens, nome_projeto)

                self.caminho_pdf_gerado = nome_pdf  # Salva para uso no botão "Revelar no Explorer"
                
                webbrowser.open(f'file:///{os.path.abspath(nome_pdf)}')
                # messagebox.showinfo("Sucesso", f"PDF gerado com sucesso: {nome_pdf}")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
                
        except Exception as e:
            print(f"Erro ao processar o arquivo Excel: {e}")

    def criar_pdf(self, df_filtrado, nome_pdf, caminho_imagens, nome_projeto):
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
            "furacao_f3": "PROG. 3",
            "ambiente":"AMBIENTE",
            "cod_material":"COD",
        }

        # Aplica o mapeamento das colunas ao DataFrame
        df_filtrado = df_filtrado.rename(columns=mapeamento_colunas)

        # Criação do objeto canvas para o PDF
        c = canvas.Canvas(nome_pdf, pagesize=landscape(letter))
        largura, altura = landscape(letter)

        # Configurações de layout
        largura_pagina = largura
        altura_pagina = altura
        
        # Definir margens
        margem_esquerda = 20
        margem_superior = altura - 20
        margem_inferior = 10

        # Título do projeto
        c.setFont("Helvetica-Bold", 14)
        c.drawString(margem_esquerda, margem_superior, f"PROJETO >> {nome_projeto}")
        
        # Posição inicial para o conteúdo
        y_pos = margem_superior - 10

        # Configurações para cada bloco de peça
        # 4 x 2 
        largura_bloco = 180  # Aumentada para acomodar mais informações 
        altura_bloco = 220   # Aumentada para acomodar códigos de barras
        margem_bloco = 10

        # 5 x 3 
        # largura_bloco = 140  # Aumentada para acomodar mais informações 
        # altura_bloco = 180   # Aumentada para acomodar códigos de barras
        # margem_bloco = 10

        # Posições iniciais
        x_pos = margem_esquerda
        y_pos_atual = y_pos

        # Processar cada peça
        for i, row in df_filtrado.iterrows():
            # Verificar se precisa de nova página
            if y_pos_atual - altura_bloco < margem_inferior:
                c.showPage()  # Nova página
                y_pos_atual = margem_superior - 40
                x_pos = margem_esquerda

            # Verificar se precisa ir para a próxima linha
            if x_pos + largura_bloco > largura_pagina - margem_esquerda:
                x_pos = margem_esquerda
                y_pos_atual -= altura_bloco + margem_bloco
                
                # Verificar se precisa de nova página após mudar de linha
                if y_pos_atual - altura_bloco < margem_inferior:
                    c.showPage()
                    y_pos_atual = margem_superior - 40
                    x_pos = margem_esquerda

            # Criar bloco para a peça atual
            self.criar_bloco_peca(c, row, x_pos, y_pos_atual, largura_bloco, altura_bloco, caminho_imagens)
            
            # Avançar para a próxima coluna
            x_pos += largura_bloco + margem_bloco

        # Salvar o PDF
        c.save()

    def criar_bloco_peca(self, c, row, x, y, largura_bloco, altura_bloco, caminho_imagens):
            """Cria um bloco com informações da peça e imagem."""
            
            # Informações conforme solicitado
            info_linhas = [
                f"PCP: {row.get('PCP', '')} Desenho: {row.get('DESENHO', '')}",
            ]

            # Adicionar furações apenas se tiverem conteúdo e não forem 'nan'
            furacoes_com_conteudo = []
            for i in range(1, 4):
                valor = row.get(f'PROG. {i}', '')
                if valor and str(valor).strip() and str(valor) != 'nan':
                    furacoes_com_conteudo.append((i, valor))

            # Adicionar furações com conteúdo à lista
            for num, valor in furacoes_com_conteudo:
                info_linhas.append(f"Furação {num}: {valor}")

            # Adicionar informações restantes
            outras_info = [
                f"Mat: {row.get('MATERIAL', '')}",
                f"Peça: {row.get('PEÇA DESCRIÇÃO', '')}",
                f"Dim: {row.get('ALTURA (X)', '')} x {row.get('PROF (Y)', '')} x {row.get('ESP', '')}",
            ]
            
            # Calcular quantas linhas extras serão necessárias devido à quebra de texto
            linhas_extras = 0
            for linha in outras_info:
                if len(linha) > 35:
                    # Calcular quantas linhas serão necessárias para esta informação
                    palavras = linha.split()
                    linha_atual = ""
                    for palavra in palavras:
                        if len(linha_atual + palavra) < 35:
                            linha_atual += palavra + " "
                        else:
                            linhas_extras += 1
                            linha_atual = palavra + " "
                    linhas_extras += 1  # Última linha
                else:
                    linhas_extras += 1
            
            # Calcular altura dinâmica
            altura_base = 20  # Altura para primeira linha
            altura_por_furacao = 15  # Altura por furação (texto + código de barras)
            altura_por_linha_quebrada = 10  # Altura por linha quebrada
            
            # Calcular altura total necessária
            altura_info = (
                altura_base +  # Primeira linha
                (len(furacoes_com_conteudo) * altura_por_furacao) +  # Furações com código de barras
                (linhas_extras * altura_por_linha_quebrada)  # Linhas das outras informações
            )
            
            # Ajustar altura mínima e máxima
            altura_info = max(80, min(altura_info, 150))  # Mínimo 80, máximo 150
            
            # Ajustar altura da imagem conforme a altura da informação
            altura_imagem = altura_bloco - altura_info - 10
            
            # Cor de fundo para o bloco de informações
            # c.setFillColor(colors.lightgrey)
            c.rect(x, y - altura_info, largura_bloco, altura_info, stroke=0)
            
            # Informações da peça
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 10)
            
            # Escrever informações
            linha_altura = y - 9
            
            # Primeira linha (PCP e Desenho)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x + 5, linha_altura, info_linhas[0])
            linha_altura -= 12
            
            # Furações com código de barras
            for linha in info_linhas[1:]:  # Pular a primeira linha (já escrita)
                if linha.startswith("Furação"):
                    c.setFont("Helvetica", 9)
                    # Extrair o valor da furação
                    code_value = linha.split(": ")[1] if ": " in linha else ""
                    
                    # Desenhar o texto da furação
                    c.drawString(x + 5, linha_altura, linha)
                    
                    if code_value and code_value.strip() != '':
                        try:
                            # Gerar código de barras Code128
                            code128 = barcode.get_barcode_class('code128')
                            barcode_image = code128(code_value, writer=ImageWriter())
                            
                            # Criar arquivo temporário
                            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                                temp_filename = temp_file.name
                            
                            # Salvar código de barras no arquivo temporário
                            barcode_image.write(temp_filename, options={
                                'module_height': 4,
                                'module_width': 0.2,
                                'font_size': 4,
                                'text_distance': 2,
                                'quiet_zone': 1,
                                'write_text': False
                            })
                            
                            # Posicionar o código de barras
                            barcode_x = x + 95
                            barcode_y = linha_altura - 6
                            barcode_width = 80
                            barcode_height = 15
                            
                            # Inserir no PDF usando o arquivo temporário
                            c.drawImage(temp_filename, barcode_x, barcode_y, 
                                        width=barcode_width, height=barcode_height)
                            
                            # Limpar arquivo temporário
                            os.unlink(temp_filename)
                            
                        except Exception as e:
                            # Fallback: desenhar retângulo se houver erro
                            print(f"Erro ao gerar código de barras: {e}")
                            c.setFillColor(colors.black)
                            c.rect(x + 60, linha_altura - 8, 80, 12, fill=1, stroke=0)
                            c.setFillColor(colors.white)
                            c.setFont("Helvetica", 9)
                            c.drawString(x + 65, linha_altura - 6, code_value[:15])
                            c.setFillColor(colors.black)
                    
                    linha_altura -= 15  # Espaço para furação + código de barras
            
            # Demais informações (com quebra de linha se necessário)
            c.setFont("Helvetica", 10)
            for linha in outras_info:
                # Quebrar linha muito longa
                if len(linha) > 35:
                    partes = []
                    parte_atual = ""
                    for palavra in linha.split():
                        if len(parte_atual + palavra) < 35:
                            parte_atual += palavra + " "
                        else:
                            partes.append(parte_atual.strip())
                            parte_atual = palavra + " "
                    partes.append(parte_atual.strip())
                    
                    for parte in partes:
                        c.drawString(x + 5, linha_altura, parte)
                        linha_altura -= 10
                else:
                    c.drawString(x + 5, linha_altura, linha)
                    linha_altura -= 10
            
            # Borda ao redor do bloco de informações
            c.setStrokeColor(colors.black)
            c.setLineWidth(0.5)
            c.rect(x, y - altura_info, largura_bloco, altura_info, fill=0, stroke=1)
            
            # Imagem da peça
            imagem_nome = row.get("DESENHO", "")
            imagem_path = os.path.join(caminho_imagens, f"{imagem_nome}.bmp")
            
            y_imagem = y - altura_info - altura_imagem - 0 # esse - é o gap entre bloco de texto e imagem
            
            if os.path.exists(imagem_path):
                try:
                    # Borda para a imagem
                    c.rect(x, y_imagem, largura_bloco, altura_imagem, fill=0, stroke=1)
                    
                    # Centralizar a imagem no espaço disponível
                    img_largura = largura_bloco - 10
                    img_altura = altura_imagem - 10
                    
                    c.drawImage(imagem_path, x + 5, y_imagem + 5, 
                            width=img_largura, height=img_altura, preserveAspectRatio=True)
                    
                    # Legenda abaixo da imagem
                    # c.setFont("Helvetica", 7)
                    # c.drawString(x + 5, y_imagem - 10, f"Desenho: {imagem_nome}")
                    
                except Exception as e:
                    # Em caso de erro ao carregar a imagem
                    c.setFont("Helvetica", 8)
                    c.drawString(x + 10, y_imagem + altura_imagem/2, "Imagem não disponível")
                    print(f"Erro ao carregar imagem {imagem_path}: {e}")
            else:
                # Mensagem quando a imagem não existe
                c.setFont("Helvetica", 8)
                c.drawString(x + 10, y_imagem + altura_imagem/2, "Imagem não encontrada")
                c.rect(x, y_imagem, largura_bloco, altura_imagem, fill=0, stroke=1)
if __name__ == "__main__":
    root = Tk()
    app = GeradorPDFApp(root)
    root.mainloop()
