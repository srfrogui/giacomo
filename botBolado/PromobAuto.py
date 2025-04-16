import os
import shutil
import time
import tkinter as tk
import pyautogui as ag
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import xlrd
from xlutils.copy import copy
import logging
from PIL import Image, ImageTk
import threading
import PyPDF2
import win32gui

from embananador import criar_arquivo_com_pecas, gerar_relatorio_pecas

def extrair_nome(caminho):
    # Extrai o nome da última pasta do caminho
    nome_pasta = os.path.basename(caminho)
    return nome_pasta

def aguarde(imagem, confianca=0.95, timeout=50, intervalo=2, inverter=False):
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

def clicar(imagem, ajusteX=0, ajusteY=0):
    localizacao = procurar(imagem)

    if localizacao:  # Verifica se a imagem foi encontrada
        print(f"Imagem localizada em {localizacao}")
        x, y = localizacao.x, localizacao.y
        posicao_certa = (x + ajusteX, y + ajusteY)
        ag.click(posicao_certa)
        print(f'Clique realizado na posição: {posicao_certa}')
    else:
        print(f"Imagem {imagem} não encontrada. Não foi possível clicar.")

def criar_pasta_vendedor(pasta, nome):
    vendedor_pasta = os.path.join(pasta, nome)
    if not os.path.exists(vendedor_pasta):
        os.makedirs(vendedor_pasta)  # Cria a pasta 'VENDEDOR'
    return vendedor_pasta

def mostrar_mensagem_erro(mensagem):
    # Show a Tkinter window to notify the user and wait for their response
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Atenção", f"{mensagem}. Clique OK para continuar.")
    root.destroy()

def salvar(pasta, nome=None):
    aguarde('./img/salvar_salvar.png')
    if nome:
        ag.write(nome)
        print(nome)
    time.sleep(0.2)
    ag.hotkey('ctrl', 'f')
    ag.hotkey(['shift', 'tab'] * 2)
    ag.press('enter')
    time.sleep(0.2)
    print(pasta)
    ag.write(pasta)
    ag.press('enter') #dentro da pasta
    time.sleep(1)
    clicar('./img/salvar_salvar.png', ajusteX=-100)

# ovo ovo oovo ovoo ovo --------------------------------------
def projeto_producao(pasta):
    # Define os caminhos das pastas de origem e destino
    source_folder = os.path.join(pasta, 'Gplan')  # Pasta de origem
    destination_folder = pasta  # Pasta de destino
    
    # Caminho do arquivo original
    original_file_path = os.path.join(source_folder, 'Projeto_producao.xls')

    # Verifica se o arquivo original existe
    if not os.path.exists(original_file_path):
        print(f"Erro: O arquivo '{original_file_path}' não foi encontrado.")
        log_message(f"Erro: O arquivo '{original_file_path}' não foi encontrado.")
        return

    # Caminho do arquivo copiado
    copied_file_path = os.path.join(destination_folder, 'Projeto_producao.xls')

    # Copiar o arquivo original para a pasta de destino
    shutil.copy(original_file_path, copied_file_path)

    # Carregar a planilha copiada (.xls)
    workbook = xlrd.open_workbook(copied_file_path, formatting_info=True)
    sheet = workbook.sheet_by_index(0)

    # Copiar o conteúdo para poder editar (já que xlrd é só leitura)
    writable_wb = copy(workbook)
    writable_sheet = writable_wb.get_sheet(0)

    # Encontrar os índices das colunas 'OBSERVAÇÕES-PROMOB' e 'AMBIENTE'
    col_observacoes = None
    col_ambiente = None
    header_row = sheet.row_values(0)

    for idx, value in enumerate(header_row):
        if value == 'OBSERVAÇÕES-PROMOB':
            col_observacoes = idx
        if value == 'AMBIENTE':
            col_ambiente = idx

    # Variável para controlar se houve alterações
    alteracoes_realizadas = False

    if col_observacoes is not None and col_ambiente is not None:
        # Verifica se existem valores na coluna 'OBSERVAÇÕES-PROMOB'
        has_values = any(sheet.cell_value(row, col_observacoes) for row in range(1, sheet.nrows))

        # Se houver valores, limpar a coluna 'AMBIENTE' e atualizar
        if has_values:
            for row in range(1, sheet.nrows):
                writable_sheet.write(row, col_ambiente, '')  # Limpa a célula

            # Atualizar as células da coluna 'AMBIENTE' com os valores de 'OBSERVAÇÕES-PROMOB'
            for row in range(1, sheet.nrows):
                obs_value = sheet.cell_value(row, col_observacoes)
                if obs_value:  # Se não for vazio ou None
                    writable_sheet.write(row, col_ambiente, obs_value)
                    alteracoes_realizadas = True  # Marcar que houve alteração

            # Salvar o arquivo modificado
            writable_wb.save(copied_file_path)
            
            if alteracoes_realizadas:
                print("Planilha copiada e atualizada com sucesso!")
                log_message("Planilha copiada e atualizada com sucesso!")
            else:
                print("Nenhuma alteração foi feita na tabela, pois não há valores na coluna 'OBSERVAÇÕES-PROMOB'.")
                log_message("Nenhuma alteração necessaria")
        else:
            print("Nenhum valor encontrado na coluna 'OBSERVAÇÕES-PROMOB'. A coluna 'AMBIENTE' não foi limpa.")
            log_message("Nenhum valor na coluna 'OBSERVAÇÕES-PROMOB'.")
    else:
        print("As colunas 'OBSERVAÇÕES-PROMOB' ou 'AMBIENTE' não foram encontradas.")
        log_message("As colunas 'OBSERVAÇÕES-PROMOB' ou 'AMBIENTE' não foram encontradas.")
    
def processo_gplan(pasta):
    clicar('./img/proce_gplan.png', ajusteY=-40)
    clicar('./img/proce_gplan.png')
    ag.press('down')
    ag.press('enter')
    salvar(pasta)
    aguarde('./img/proce_vizu.png', timeout=1000, intervalo=10)
    ag.press('right')
    ag.press('enter')
    aguarde('./img/proce_cncgen.png')
    ag.press('enter')
    time.sleep(1)
    ag.hotkey('alt', 'F4')

def process_pdf(vendedor_pasta, x_adjustV, name):
    clicar('./img/proce_pdf.png', ajusteY=20, ajusteX=x_adjustV)
    aguarde('./img/load_pdf.png')
    if procurar('./img/validacao_pdf_visto.png'):
        clicar('./img/proce_pdf.png')  # gera pdf caso ache o validacao
        salvar(vendedor_pasta, nome=name)
        if name == "Router":
            #log_message("Aviso: TEM ROUTER SALVE O DESENHO!")
            return True  
        else:
            return False      

def processo_dinheirinho(pasta):
    aguarde('./img/proce_money.png')
    def process_pedido_fabrica(vendedor_pasta):
        clicar('./img/proce_money.png', ajusteY=-40)
        clicar('./img/proce_money.png')
        ag.press(['down'] * 3 + ['right'] + ['down'] * 3 + ['enter'])
        aguarde('./img/proce_pdf.png')
        clicar('./img/proce_pdf.png')
        salvar(vendedor_pasta, nome='PedidoFabrica')
        time.sleep(0.5)
        fechar_processo()
        
    def process_pedido_vidro(vendedor_pasta):
        clicar('./img/proce_money.png', ajusteY=-40)
        clicar('./img/proce_money.png')
        ag.press(['down'] * 3 + ['right'] + ['down'] * 4 + ['enter'])
        aguarde('./img/proce_pdf.png')
        if procurar('./img/prd_vidros.png'):
            clicar('./img/proce_pdf.png')  # gera pdf caso ache o validacao
            salvar(vendedor_pasta, nome='PedidoVidro')
            log_message('Vidros Gerado!')
        time.sleep(0.5)
        fechar_processo()
        
    def process_listagem_completa(vendedor_pasta):
        clicar('./img/proce_money.png', ajusteY=-40)
        clicar('./img/proce_money.png')
        ag.press(['down'] * 3 + ['right'] + ['down'] * 2 + ['enter'])
        aguarde('./img/proce_pdf.png')
        clicar('./img/proce_pdf.png')
        salvar(vendedor_pasta, nome='ListagemCompleta')
        time.sleep(0.5)
        fechar_processo()

    def process_pcp(pasta):
        clicar('./img/proce_money.png', ajusteY=-40)
        clicar('./img/proce_money.png')
        ag.press(['down'] * 3 + ['right'] + ['down'] * 5 + ['right'] + ['down'] * 3 + ['enter'])
        aguarde('./img/proce_pdf.png')
        clicar('./img/proce_pdf.png', ajusteX=200)
        ag.press('down')
        ag.press('enter')
        salvar(pasta)  # Salva o PCP
        time.sleep(0.5)
        ag.press('right')
        ag.press('enter')
        fechar_processo()

    def process_todos_relatorios(vendedor_pasta):
        clicar('./img/proce_money.png', ajusteY=-40)
        clicar('./img/proce_money.png')
        ag.press(['down'] * 3 + ['right'] + ['down'] * 6 + ['right'] + ['enter'])
        aguarde('./img/proce_pdf.png')
        # process_pdf(vendedor_pasta, 0, 'Listagem_Pecas')
        # process_pdf(vendedor_pasta, 480, 'Frentes')
        variavel = process_pdf(vendedor_pasta, 600, 'Router')
        process_pdf(vendedor_pasta, 700, 'Perfil')
        process_pdf(vendedor_pasta, 800, 'Composto')
        fechar_processo()
        return variavel
        
    def verificar_e_apagar_pdf(vendedor_pasta): # fazer com oo outro pdf LISTAGEM.pdf >> Abrir em 33mm

    #AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
        # caminho_listagempecas = os.path.join(vendedor_pasta, 'Listagem_Pecas.pdf')
        # if os.path.exists(caminho_listagempecas):
        #     try:
        #         with open(caminho_listagempecas, 'rb') as arquivo:
        #             leitorr = PyPDF2.PdfReader(arquivo)
        #             conteudor = ""

        #             for pagina in leitorr.pages:
        #                 conteudor += pagina.extract_text() or ""

        #         if "Abrir em " not in conteudor:
        #             for tentativar in range(10):  # Tenta por 10 vezes
        #                 try:
        #                     os.remove(caminho_listagempecas)
        #                     print(f"{caminho_listagempecas} apagado, pois não contém os termos necessários.")
        #                     break
        #                 except PermissionError:
        #                     print(f"Tentativa {tentativar + 1} falhou. O arquivo está em uso. Tentando novamente...")
        #                     time.sleep(1)  # Aguardando 1 segundo antes de tentar novamente
        #             else:
        #                 print(f"Falha ao remover {caminho_listagempecas} após várias tentativas.")
        #         else:
        #             print(f"{caminho_listagempecas} mantido, pois contém os termos necessários.")
        #     except Exception as e:
        #         print(f"Erro ao ler o PDF Listagem_Pecas.pdf: {e}")
        # else:
        #     print(f"Arquivo {caminho_listagempecas} não encontrado.")
            
        
    #AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA    
       
        # caminho_frentes = os.path.join(vendedor_pasta, 'Frentes.pdf')
        # print("Caminho:frentes", caminho_frentes)
        # if os.path.exists(caminho_frentes):
        #     try:
        #         with open(caminho_frentes, 'rb') as arquivo:
        #             leitor = PyPDF2.PdfReader(arquivo)
        #             conteudo = ""

        #             for pagina in leitor.pages:
        #                 conteudo += pagina.extract_text() or ""

        #         if "CORTE 45G" not in conteudo and "PERFIL 45G" not in conteudo and "Articulador" not in conteudo and "Utilitário" not in conteudo:
        #             for tentativa in range(10):  # Tenta por 10 vezes
        #                 try:
        #                     os.remove(caminho_frentes)
        #                     print(f"{caminho_frentes} apagado, pois não contém os termos necessários.")
        #                     break
        #                 except PermissionError:
        #                     print(f"Tentativa {tentativa + 1} falhou. O arquivo está em uso. Tentando novamente...")
        #                     time.sleep(1)  # Aguardando 1 segundo antes de tentar novamente
        #             else:
        #                 print(f"Falha ao remover {caminho_frentes} após várias tentativas.")
        #         else:
        #             print(f"{caminho_frentes} mantido, pois contém os termos necessários.")
        #     except Exception as e:
        #         print(f"Erro ao ler o PDF: {e}")
        # else:
        #     print(f"Arquivo {caminho_frentes} não encontrado.")
        print('minipa')
        
    def fechar_processo():
        ag.hotkey('alt', 'F4')
        
    # Chamando as funções otimizadas
    vendedor_pasta = criar_pasta_vendedor(pasta, 'VENDEDOR')
    #process_pedido_fabrica(vendedor_pasta)
    #process_pedido_vidro(vendedor_pasta)
    process_listagem_completa(vendedor_pasta)
    process_pcp(pasta)
    variavel = process_todos_relatorios(vendedor_pasta)
    time.sleep(1)
    verificar_e_apagar_pdf(vendedor_pasta)
    return variavel

log_file_P = 'promob.log'

def processo_completin():
    for pasta in list(pastas):
        try:
            with open(log_file_P, 'a') as log:
                log_message(f'Processando pasta: {pasta}\n')
            
            for pasta in pastas:
                if var_dinheirinho.get():
                    log_message("Iniciando Relatorios...")
                    leite = processo_dinheirinho(pasta)
                if var_gplan.get():
                    log_message("Gerando Gplan...")
                    processo_gplan(pasta)
                if var_producao.get():
                    log_message("Copiando Porojeto_producao ...")
                    projeto_producao(pasta)
                if var_RPecas.get():
                    log_message("Gerando Relatorio Pecas ...")
                    gerar_relatorio_pecas(pasta+"Projeto_producao.xls")
                if var_NPecas.get():
                    log_message("Gerando Relatorio Pecas ...")
                    criar_arquivo_com_pecas(pasta+"Projeto_producao.xls")
                if leite:
                    #mostrar_mensagem_erro("Aviso: TEM ROUTER SALVE O DESENHO!")
                    ag.alert(text="Aviso: TEM ROUTER SALVE O DESENHO!", title="SALVE IMEDIATAMENTE", button="SALVADO!")
            with open(log_file_P, 'a') as log:
                log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
                log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')
                
            pastas.remove(pasta)
            atualizar_frame_pastas()
            
        except ValueError as e:
            with open(log_file_P, 'a') as log:
                log.write(f'Erro ao processar pasta {pasta}: {e}\n')
                log_message(f'Erro ao processar pasta {pasta}: {e}\n')
                mostrar_mensagem_erro(f"Aviso: Erro ao processar pasta {pasta}: {e}\n")
            break
        except Exception as e:
            with open(log_file_P, 'a') as log:
                log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                log_message(f'Erro generico ao processar pasta {pasta}: {e}\n')
                mostrar_mensagem_erro(f"Aviso: Erro generico ao processar pasta {pasta}: {e}\n")
            continue

 
    log_message("Processo Finalizado...")
    
def log_message(message):
    if text_log:
        text_log.insert(tk.END, message + '\n')
        text_log.see(tk.END)

def manda_pra_frente(window_title_part):
    def windowEnumerationHandler(hwnd, top_windows):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

    top_windows = []
    win32gui.EnumWindows(windowEnumerationHandler, top_windows)
    
    found = False
    for hwnd, title in top_windows:
        if window_title_part.lower() in title.lower():
            print(f"Trazendo a janela: {title}")
            win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
            time.sleep(0.1)  # Atraso
            win32gui.SetForegroundWindow(hwnd)
            found = True
            break

    if not found:
        print(f"Janela contendo '{window_title_part}' não encontrada.")

def atualizar_frame_pastas():
    # Clear the frame content
    for widget in frame_pastas.winfo_children():
        widget.destroy()

    # Re-add each remaining folder to the frame
    for pasta in pastas:
        tk.Label(frame_pastas, text=pasta, fg="blue").pack(anchor='w', padx=5)

def main():
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
        # Extrair o nome da pasta
        nome_pasta = pasta.split('/')[-1]
        pasta_label = tk.Label(frame_pastas, text=nome_pasta, font=("Arial", 12), relief=tk.RAISED)

        # Evento para mostrar o caminho completo ao passar o mouse sobre o nome
        pasta_label.bind("<Enter>", lambda event, p=pasta: caminho_label.config(text=p))
        pasta_label.pack(pady=2)

    def ok():
        logging.info('Botão OK pressionado. Fechando a aplicação.')  # Fecha a janela e encerra o loop do tkinter
        if not pastas :
            messagebox.showinfo("Atenção", "Selecione uma pasta.")
            return
        threading.Thread(target=processo_completin).start()
        log_message("Processo Iniciado...")
        
    def on_close():
        janela.destroy()
        os._exit(0)   
        
    global caminho_label, frame_pastas, pastas, text_log
    pastas = []

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

    global var_gplan, var_dinheirinho, var_producao, var_RPecas, var_NPecas
    # Checkboxes para selecionar quais processos executar
    var_gplan = tk.BooleanVar(value=True)
    var_dinheirinho = tk.BooleanVar(value=True)
    var_producao = tk.BooleanVar(value=True)
    var_RPecas = tk.BooleanVar(value=False)
    var_NPecas = tk.BooleanVar(value=True)
    
    # Frame para organizar os checkboxes horizontalmente
    checkbox_frame = tk.Frame(janela)
    checkbox_frame.pack(pady=3)

    # Checkboxes adicionados ao frame com layout grid
    tk.Checkbutton(checkbox_frame, text="Processo Dinheirinho", variable=var_dinheirinho).grid(row=0, column=0, sticky='w', padx=5)
    tk.Checkbutton(checkbox_frame, text="Processo GPlan", variable=var_gplan).grid(row=0, column=1, sticky='w', padx=5)
    tk.Checkbutton(checkbox_frame, text="Projeto Produção", variable=var_producao).grid(row=0, column=2, sticky='w', padx=5)
    tk.Checkbutton(checkbox_frame, text="Relatorio Pecas", variable=var_RPecas).grid(row=1, column=1, sticky='w', padx=5)
    tk.Checkbutton(checkbox_frame, text="Contar Pecas", variable=var_NPecas).grid(row=1, column=2, sticky='w', padx=5)

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
    
    
def oloco():
    print('banana')
    
if __name__ == '__main__': 
    main()
    
