# Início do script
from GAuto import novo_projeto, importa, importar_projeto, abrir_parametro, configurar_optimizacao, imprimir_loop, gerar_gvision, abrir_producao, verificar_arquivos, extrair_nome, log_file_G
from PromobAuto import processo_dinheirinho, processo_gplan, atualizar_frame_pastas, mostrar_mensagem_erro, log_file_P
from embananador import criar_arquivo_com_pecas, gerar_relatorio_pecas, gerar_aciete
from G2Auto import obter_caminhos, importar_optimiza, exportar_plano_corte , limpar_lista, log_file_N, clicar, aguarde, procurar, obter_nome,compress_to_rar, gerar_relatorio_pdf, gerar_pdfs
from Moveu import moveu, log_file_M
from contar_chapas import gerar_pdf_com_tabela
from arrasta_banana import main as arrasta_banana

import multiprocessing
import time
import win32gui
import os
import xlrd
import tkinter as tk
import threading
import pandas as pd
import shutil
from xlutils.copy import copy
from tkinter import filedialog, messagebox, IntVar
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkcalendar import DateEntry
import sys

print("Rodando sem try exept pra ver as bomba dos erro")
print("Diretório atual:", os.getcwd())  # Mostra o diretório atual
# Defina o diretório de trabalho para o diretório do script
script_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
os.chdir(script_dir)
print("Diretório Script:", script_dir)

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
    check_pasta_count()
    
def remove_pasta(path):
    # Remove a pasta da lista e atualiza a visualização
    if path in pastas:
        pastas.remove(path)
        update_pasta_view()
        check_pasta_count() 
    
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

# def processar_pastas_gplan(pasta):
#     global log_file_G
#     log_file_G = 'gplano.log'
#     # Abrir o GPlan uma vez
#     #abrir_gplan() ## quando abre o gplan por aqui por algum motivo ele n consegue gerar_gvision
#     gplano, cuti = verificar_arquivos(pasta)
#     with open(log_file_G, 'a') as log:
#         log.write(f'Processando pasta: {pasta}\n')
        
#     # Etapas do processamento
#     if var_novo_projeto.get():
#         log_message("Iniciando NovoProjeto...")
#         novo_projeto()
#     if var_importa.get():
#         log_message("Iniciando Importacao 1/2...")
#         importa()
#     if var_importar_projeto.get():
#         log_message("Importando Projeto 2/2...")
#         importar_projeto(pasta, gplano, cuti)
#     if var_abrir_parametro.get():
#         abrir_parametro()
#     if var_configurar_otimizacao.get():
#         log_message("Configurando optimizacao...")
#         configurar_optimizacao()
#     if var_imprimir_loop.get():
#         log_message("Gerando PDF's...")
#         imprimir_loop(pasta)
#     if var_gerar_gvision.get():
#         log_message("Gerando GVision...")
#         gerar_gvision(pasta)
#     if var_abrir_producao.get():
#         log_message("Retornando Loop...")
#         abrir_producao()
    
#     with open(log_file_G, 'a') as log:
#         log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
#         log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')
        

# def processo_completin():
#     global log_file_P
#     log_file_P = 'processamento.log'
#     with open(log_file_P, 'a') as log:
#         log_message(f'Processando pasta: {pasta}\n')
    
#     for pasta in pastas:
#         if var_dinheirinho.get():
#             log_message("Iniciando Relatorios...")
#             processo_dinheirinho(pasta)
#         if var_gplan.get():
#             log_message("Gerando Gplan...")
#             processo_gplan(pasta)
#         if var_producao.get():
#             log_message("Copiando Porojeto_producao ...")
#             projeto_producao(pasta)
#         if var_RPecas.get():
#             log_message("Gerando Relatorio peças ...")
#             ovo(pasta + "/Projeto_producao.xls")
    
    
#     with open(log_file_P, 'a') as log:
#         log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
#         log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')
        
#================================================================================================================================================================================================================================        

def processar_pastas_gplan(pasta):
    # Abrir o GPlan uma vez
    #abrir_gplan() ## quando abre o gplan por aqui por algum motivo ele n consegue gerar_gvision
    gplano, cuti = verificar_arquivos(pasta)
    with open(log_file_G, 'a') as log:
        log.write(f'Processando pasta: {pasta}\n')
    log_message(f'Processando pasta: {pasta}\n')

    # Etapas do processamento
    if var_novo_projeto.get():
        with open(log_file_G, 'a') as log:
            log.write("Iniciando NovoProjeto...")
        log_message("Iniciando NovoProjeto...")
        novo_projeto()
    if var_importa.get():
        with open(log_file_G, 'a') as log:
            log.write("Iniciando Importacao 1/2...")
        log_message("Iniciando Importacao 1/2...")
        importa()
    if var_importar_projeto.get():
        with open(log_file_G, 'a') as log:
            log.write("Iniciando Importacao 2/2...")
        log_message("Importando Projeto 2/2...")
        importar_projeto(pasta, gplano, cuti)
    if var_abrir_parametro.get():
        abrir_parametro()
    if var_configurar_otimizacao.get():
        with open(log_file_G, 'a') as log:
            log.write("Configurando optimizacao...")
        log_message("Configurando optimizacao...")
        configurar_optimizacao()
    if var_imprimir_loop.get():
        with open(log_file_G, 'a') as log:
            log.write("Gerando PDF's...")
        log_message("Gerando PDF's...")
        imprimir_loop(pasta)
    if var_gerar_gvision.get():
        with open(log_file_G, 'a') as log:
            log.write("Gerando GVision...")
        log_message("Gerando GVision...")
        gerar_gvision(pasta)
    if var_abrir_producao.get():
        with open(log_file_G, 'a') as log:
            log.write("Retornando Loop...")
        log_message("Retornando Loop...")
        abrir_producao()
    
    with open(log_file_G, 'a') as log:
        log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
    log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')
#================================================================================================================================================================================================

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


def processo_completin(pasta):
    with open(log_file_P, 'a') as log:
        log.write(f'Processando pasta: {pasta}\n')
    log_message(f'Processando pasta: {pasta}\n')
    
    if var_dinheirinho.get():
        with open(log_file_P, 'a') as log:
            log.write("Iniciando Relatorios...")
        log_message("Iniciando Relatorios...")
        router = processo_dinheirinho(pasta)
    if var_gplan.get():
        with open(log_file_P, 'a') as log:
            log.write("Gerando Gplan...")
        log_message("Gerando Gplan...")
        processo_gplan(pasta)
    if router:
        #mostrar_mensagem_erro("Aviso: TEM ROUTER SALVE O DESENHO!")
        log_message("ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER")
        log_message("ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER")
        log_message("ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER")
        log_message("ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER")
        log_message("ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER")
        log_message("ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER ROUTER")
        
    with open(log_file_P, 'a') as log:
        log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
    log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')


def processo_completin_loopavel(pasta):
    with open(log_file_P, 'a') as log:
        log.write(f'Processando L pasta: {pasta}\n')
    log_message(f'Processando L pasta: {pasta}\n')
    print("ovo")
    if var_producao.get():
        with open(log_file_P, 'a') as log:
            log.write("Copiando Porojeto_producao ...")
        log_message("Copiando Porojeto_producao ...")
        projeto_producao(pasta)
    
    try:
        arquivo_xls = os.path.join(pasta, 'Projeto_producao.xls')
        df = pd.read_excel(arquivo_xls)
    except FileNotFoundError:
        with open(log_file_P, 'a') as log:
            log.write(f"Oi! Arquivo não encontrado: {arquivo_xls}\n")
        log_message(f"Oi! Arquivo não encontrado: {arquivo_xls}\n")
    
    if var_RPecas.get():
        with open(log_file_P, 'a') as log:
            log.write("Gerando Relatorio Pecas ...")
        log_message("Gerando Relatorio Pecas ...")
        gerar_relatorio_pecas(df, arquivo_xls)

    if var_NPecas.get():
        with open(log_file_P, 'a') as log:
            log.write("Contando Pecas ...")
        log_message("Contando Pecas ...")
        criar_arquivo_com_pecas(df, arquivo_xls)
        
    with open(log_file_P, 'a') as log:
        log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
        log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')

#================================================================================================================================================================================================
def processo_nesting(pasta):
    with open(log_file_N, 'a') as log:
        log.write(f'Processando pasta: {pasta}\n')
        
    cut, nesting, corte, vendedor = obter_caminhos(pasta)
    try:
        print(f'Valor de var_importa_n: {var_importa_n.get()}')
    except Exception as e:
        print(e)
        
    # Etapas do processamento
    if var_importa_n.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Importando e optimizando...\n')
        log_message("Importando e optimizando...")
        importar_optimiza(pasta, cut)
    
    if var_exportar_n.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Exportando planos de corte...\n')
        log_message("Exportando planos de corte...")
        exportar_plano_corte(nesting, corte)
    
    if var_limpar_lista_n.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Limpando lista...')
        log_message("Limpando lista...")
        limpar_lista()
        
    if var_relatorio_pdf_n.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Gerando PDFs...')
        log_message("Gerando PDFs...")
        gerar_pdfs(corte, vendedor)
        
    if gerar_pdf_html_n.get():
        with open(log_file_N, 'a') as log:
            log.write(f'Gerando relatorio InfoOutput...')
        log_message("Gerando relatorio InfoOutput...")
        gerar_relatorio_pdf(corte, vendedor, pasta)
    
    with open(log_file_N, 'a') as log:
        log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
        log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')







#================================================================================================================================================================================================


    
def log_message(message):
    if text_log:
        text_log.insert(tk.END, message + '\n')
        text_log.see(tk.END)
 
def main(): 
    
    def manda_pra_frente(window_title_part):
        def windowEnumerationHandler(hwnd, top_windows):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

        results = []
        top_windows = []
        win32gui.EnumWindows(windowEnumerationHandler, top_windows)
        
        found = False
        for hwnd, title in top_windows:
            if window_title_part.lower() in title.lower():
                print(f"Trazendo a janela: {title}")
                win32gui.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                time.sleep(0.1)  # Atraso
                win32gui.SetForegroundWindow(hwnd)
                found = True
                return True
            
        if not found:
            print(f"Janela contendo '{window_title_part}' não encontrada.")
            return False
            
    def mostrar_mensagem_erro(mensagem):
        # Show a Tkinter window to notify the user and wait for their response
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Atenção", f"{mensagem}. Clique OK para continuar.")
        root.destroy()
        
    def atualizar_frame_pastas():
        # Clear the frame content
        for widget in frame_pastas.winfo_children():
            widget.destroy()

        # Re-add each remaining folder to the frame
        for pasta in pastas:
            tk.Label(frame_pastas, text=pasta, fg="blue").pack(anchor='w', padx=5)
    
    def selecionar_todos_promob():
        estado = var_selecionar_todos_promob.get()
        var_gplan.set(estado)
        var_dinheirinho.set(estado)
        var_producao.set(estado)
        var_RPecas.set(estado)
        var_NPecas.set(estado)

    def toggle_chrome_inputs():
        if var_moveu.get():
            entry_vendedor.config(state='normal')
            entry_cliente.config(state='normal')
            date_prazo.config(state='normal')
        else:
            entry_vendedor.config(state='disabled')
            entry_cliente.config(state='disabled')
            date_prazo.config(state='disabled')
                       
    def selecionar_todos_gplan():
        estado = var_selecionar_todos_gplan.get()
        var_novo_projeto.set(estado)
        var_importa.set(estado)
        var_importar_projeto.set(estado)
        var_abrir_parametro.set(estado)
        var_configurar_otimizacao.set(estado)
        var_imprimir_loop.set(estado)
        var_gerar_gvision.set(estado)
        var_abrir_producao.set(estado)
        
    def selecionar_todos_nesting():
        estado = var_selecionar_todos_nesting.get()
        var_importa_n.set(estado)
        var_exportar_n.set(estado)
        var_limpar_lista_n.set(estado)
        var_relatorio_pdf_n.set(estado)
        gerar_pdf_html_n.set(estado)

    def ok():
        
        if not pastas:
            messagebox.showinfo("Atenção", "Selecione uma pasta.")
            return
        
        if var_moveu.get():
            if not entry_vendedor.get() or not entry_cliente.get() or not date_prazo.get_date():
                messagebox.showinfo("Atenção", "Os campos Vendedor, Cliente e Prazo são obrigatórios.")
                return
        
        if pastas:
            pasta = pastas[0]
            if var_gplan.get() or var_dinheirinho.get():
                # try:
                processo_completin(pasta)
                # pro1 = threading.Thread(target=processo_completin, args=(pasta,))
                # pro1.start()
                # pro1.join()
                    
                # except ValueError as e:
                #     with open(log_file_P, 'a') as log:
                #         log.write(f'Erro ao processar pasta {pasta}: {e}\n')
                #         log_message(f'Erro ao processar pasta {pasta}: {e}\n')
                #         mostrar_mensagem_erro(f"Aviso: Erro ao processar pasta {pasta}: {e}\n")
                # except Exception as e:
                #     with open(log_file_P, 'a') as log:
                #         log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                #         log_message(f'Erro generico ao processar pasta {pasta}: {e}\n')
                #     mostrar_mensagem_erro(f"Aviso: Erro generico ao processar pasta {pasta}: {e}\n")        
            
        for pasta in pastas:   
            
            if var_producao.get() or var_RPecas.get() or var_NPecas.get():
                # try:
                processo_completin_loopavel(pasta)
                # pro2 = threading.Thread(target=processo_completin_loopavel, args=(pasta,))
                # pro2.start()
                # pro2.join()
                    
                # except ValueError as e:  # Captura especificamente o erro lançado na função clicar
                #     with open(log_file_P, 'a') as log:
                #         log.write(f'Erro ao processar pasta {pasta}: {e}\n')
                #         log_message(f'Erro ao processar pasta {pasta}: {e}\n')
                #         mostrar_mensagem_erro("Aviso: Erro ao processar pasta {pasta}: {e}\n")
                #     break
                # except Exception as e:
                #     with open(log_file_P, 'a') as log:
                #         log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                #         mostrar_mensagem_erro("Aviso: Erro generico ao processar pasta {pasta}: {e}\n")
                #     break
                
            if var_moveu.get():
                # try:
                
                    
                log_message("Processo Moveu Iniciado...")
                with open(log_file_M, 'a') as log:
                    log.write(f'Processando pasta: {pasta}\n')
                prazo = date_prazo.get_date().strftime("%d%m%Y")
                nomemov = (entry_vendedor.get() + " - " + entry_cliente.get()).upper()
                print(prazo, '__', nomemov)
            
                moveu(pasta, prazo, nomemov)
            
                # mov = threading.Thread(target=moveu, args=(pasta, prazo, nomemov,))
                # mov.start()
                # mov.join()
                
                # except Exception as e:
                #     with open(log_file_M, 'a') as log:
                #         log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                #         mostrar_mensagem_erro("Aviso: Erro generico ao processar pasta {pasta}: {e}\n")
                #     break
                    
                    
                
            if var_selecionar_todos_gplan.get() or var_novo_projeto.get() or var_importa.get() or var_importar_projeto.get() or var_abrir_parametro.get() or var_configurar_otimizacao.get() or var_imprimir_loop.get() or var_gerar_gvision.get() or var_abrir_producao.get():
                # Inicia o segundo processo após o primeiro terminar
                log_message(f"Processo Gplan Iniciado...{pasta}")
                if var_novo_projeto.get():
                    clicar("./img/abrir_gplan.png")
                    time.sleep(3)
                    if procurar("./img/val_chave.png"):
                        mostrar_mensagem_erro("ENFIA A CHAVE!")
                        raise Exception("Sem Chave :(")
                    else:
                        aguarde("./img/ta_aberto.png", timeout=15)

                # try:
                    # gplan = threading.Thread(target=processar_pastas_gplan, args=(pasta,))
                    # gplan.start()
                    # gplan.join()
                
                    
                processar_pastas_gplan(pasta)
                
                clicar("./img/abrir_gplan.png")     
                  
                # except ValueError as e:  # Captura especificamente o erro lançado na função clicar
                #     with open(log_file_G, 'a') as log:
                #         log.write(f'Erro ao processar pasta {pasta}: {e}\n')
                #         log_message(f'Erro ao processar pasta {pasta}: {e}\n')
                #         mostrar_mensagem_erro("Aviso: Erro ao processar pasta {pasta}: {e}\n")
                #     break
                # except Exception as e:
                #     with open(log_file_G, 'a') as log:
                #         log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                #         mostrar_mensagem_erro("Aviso: Erro generico ao processar pasta {pasta}: {e}\n")
                #     break
                
            if var_importa_n.get() or var_exportar_n.get() or var_limpar_lista_n.get() or var_relatorio_pdf_n.get() or gerar_pdf_html_n.get():
                if var_importa_n.get():
                    log_message(f"Processo Nesting Iniciado...{pasta}")
                    # if not manda_pra_frente('wsnesting'):

                    clicar("./img/abrir_nesting.png")
                    time.sleep(3)
                    aguarde('./img/btt_carregar_arq.png', timeout=15)
            
                #try:
                    # nesting = threading.Event()
                    # nest = threading.Thread(target=processo_nesting, args=(pasta,nesting))
                    # nest.start()
                    # nesting.wait()
                processo_nesting(pasta)
                
                clicar("./img/abrir_nesting.png")
                # except ValueError as e:  # Captura especificamente o erro lançado na função clicar
                #     with open(log_file_G, 'a') as log:
                #         log.write(f'Erro ao processar pasta {pasta}: {e}\n')
                #         log_message(f'Erro ao processar pasta {pasta}: {e}\n')
                #         mostrar_mensagem_erro("Aviso: Erro ao processar pasta {pasta}: {e}\n")
                #     break
                # except Exception as e:
                #     with open(log_file_G, 'a') as log:
                #         log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                #         mostrar_mensagem_erro("Aviso: Erro generico ao processar pasta {pasta}: {e}\n")
                #     break
                
            if remover_lista.get() or contar_chapa.get() or compress_vend.get() or gerar_json.get():
                vendedor = os.path.join(pasta, "VENDEDOR")
                cliente = obter_nome(pasta)
                
                if contar_chapa.get():
                    log_message("Contando Chapas...")
                    gerar_pdf_com_tabela(vendedor,pasta)
                
                if gerar_json.get():
                    log_message("Gerando json...")
                    gerar_aciete(pasta)          
                            
                if compress_vend.get():
                    log_message("Compactando VENDEDOR...")
                    compress_to_rar(vendedor, cliente)

                #if compress_vend.get():
                #    log_message("Arrastando Banana...")
                #    #arrasta_banana(pasta)

                if remover_lista.get():
                    log_message("removendopastadalista...")
                    pastas.remove(pasta)
                    atualizar_frame_pastas() 
            
    global check_pasta_count
    def check_pasta_count():
        # Se houver mais de uma pasta, desabilita os checkboxes relacionados ao Promob
        if len(pastas) > 1:
            var_selecionar_todos_promob.set(0)
            var_gplan.set(0)
            var_dinheirinho.set(0)
            checkbutton_selecionar_todos_promob.config(state='disabled')
            checkbutton_gplan.config(state='disabled')
            checkbutton_dinheirinho.config(state='disabled')
            var_abrir_producao.set(1)
            var_limpar_lista_n.set(1)
        else:
            # Reabilita os checkboxes se apenas uma pasta estiver selecionada
            checkbutton_selecionar_todos_promob.config(state='normal')
            checkbutton_gplan.config(state='normal')
            checkbutton_dinheirinho.config(state='normal')
            
    def on_close():
        janela.destroy()
        os._exit(0)
    
    global caminho_label, frame_pastas, pastas, text_log
    # Criar janela principal
    janela = TkinterDnD.Tk()    
    janela.title("Seleção de Pastas")
    janela.geometry("600x800")
    
    pastas=[]

    # Variáveis para checkboxes (definidas após criar a janela)
    global var_gplan, var_dinheirinho, var_producao, var_RPecas, var_NPecas
    var_selecionar_todos_promob = IntVar(value=0)
    var_gplan = IntVar(value=1)
    var_dinheirinho = IntVar(value=1)
    var_producao = IntVar(value=1)
    var_RPecas = IntVar(value=0)
    var_NPecas = IntVar(value=1)
    
    global var_moveu
    var_moveu = IntVar(value=1)
    
    global var_novo_projeto, var_importa, var_importar_projeto, var_abrir_parametro, var_configurar_otimizacao, var_imprimir_loop, var_gerar_gvision, var_abrir_producao
    var_selecionar_todos_gplan = IntVar(value=1)
    var_novo_projeto = IntVar(value=1)
    var_importa = IntVar(value=1)
    var_importar_projeto = IntVar(value=1)
    var_abrir_parametro = IntVar(value=1)
    var_configurar_otimizacao = IntVar(value=1)
    var_imprimir_loop = IntVar(value=1)
    var_gerar_gvision = IntVar(value=1)
    var_abrir_producao = IntVar(value=1)
    janela.protocol("WM_DELETE_WINDOW", on_close)

    global var_importa_n, var_exportar_n, var_limpar_lista_n, var_relatorio_pdf_n, gerar_pdf_html_n, var_selecionar_todos_nesting
    # Variáveis para checkboxes (definidas após criar a janela)
    var_selecionar_todos_nesting = IntVar(value=1)
    var_importa_n = IntVar(value=1)
    var_exportar_n= IntVar(value=1)
    var_limpar_lista_n = IntVar(value=1)
    var_relatorio_pdf_n = IntVar(value=1)
    gerar_pdf_html_n = IntVar(value=1)
    
    global remover_lista, contar_chapa, compress_vend, gerar_json
    contar_chapa = IntVar(value=0) # gera contagem de chapa em pdf separado "cumaru 15mm 1 - gplan 1 - nesting" ...
    gerar_json = IntVar(value=1)
    compress_vend = IntVar(value=1)
    remover_lista = IntVar(value=0)
    #copiar_pasta = IntVar(value=0)
    
    # Frame para exibir as pastas
    frame_pastas = tk.Frame(janela)
    frame_pastas.pack(pady=2, padx=10)

    # Label para exibir o caminho completo
    caminho_label = tk.Label(janela, text="", fg="blue")
    caminho_label.pack(pady=5)

    # Adicionar checkboxes em um Frame
    checkbox_frame = tk.Frame(janela)
    checkbox_frame.pack(pady=2)
    
    tk.Label(checkbox_frame, text="Promob process").grid(row=0, column=0, sticky='w')
    checkbutton_selecionar_todos_promob = tk.Checkbutton(
        checkbox_frame, text="Selecionar Todes", variable=var_selecionar_todos_promob,
        command=selecionar_todos_promob  # Conecta a função
    )
    checkbutton_selecionar_todos_promob.grid(row=0, column=1, sticky='w')

    checkbutton_dinheirinho = tk.Checkbutton(checkbox_frame, text="Processo Dinheirinho", variable=var_dinheirinho)
    checkbutton_dinheirinho.grid(row=0, column=3, sticky='w')

    checkbutton_gplan = tk.Checkbutton(checkbox_frame, text="Processo GPlan", variable=var_gplan)
    checkbutton_gplan.grid(row=1, column=0, sticky='w')
    
    tk.Checkbutton(checkbox_frame, text="Projeto Produção", variable=var_producao).grid(row=1, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Relatorio Pecas", variable=var_RPecas).grid(row=1, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Contar Pecas", variable=var_NPecas).grid(row=1, column=3, sticky='w')
    
    tk.Label(checkbox_frame, text="Chrome process").grid(row=2, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Moveu", variable=var_moveu, command=toggle_chrome_inputs).grid(row=2, column=1, sticky='w')
    # Campo de texto e data para o processo "Chrome"
    text_label = tk.Label(checkbox_frame, text="Vendedor:")
    text_label.grid(row=2, column=2, sticky='w', padx=1)
    entry_vendedor = tk.Entry(checkbox_frame, state='normal')
    entry_vendedor.grid(row=2, column=3, sticky='w', padx=1)
    
    text_label = tk.Label(checkbox_frame, text="Cliente:")
    text_label.grid(row=3, column=0, sticky='w', padx=1)
    entry_cliente = tk.Entry(checkbox_frame, state='normal')
    entry_cliente.grid(row=3, column=1, sticky='w', padx=1)

    date_label = tk.Label(checkbox_frame, text="Prazo:")
    date_label.grid(row=3, column=2, sticky='w', padx=10)
    date_prazo = DateEntry(checkbox_frame, width=17, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy', state='normal')
    date_prazo.grid(row=3, column=3, sticky='w', padx=1)
    
    tk.Label(checkbox_frame, text="GPlan process").grid(row=4, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Selecionar Todes", variable=var_selecionar_todos_gplan, command=selecionar_todos_gplan).grid(row=4, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Novo Projeto", variable=var_novo_projeto).grid(row=5, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Importa", variable=var_importa).grid(row=5, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Importar Projeto", variable=var_importar_projeto).grid(row=5, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Abrir Parâmetro", variable=var_abrir_parametro).grid(row=5, column=3, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Configurar Otimização", variable=var_configurar_otimizacao).grid(row=6, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Imprimir Loop", variable=var_imprimir_loop).grid(row=6, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Gerar GVision", variable=var_gerar_gvision).grid(row=6, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Abrir Produção", variable=var_abrir_producao).grid(row=6, column=3, sticky='w')
    
    # Checkboxes adicionados ao frame com layout grid
    tk.Label(checkbox_frame, text="Nesting process").grid(row=7, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Selecionar Todes", variable=var_selecionar_todos_nesting, command=selecionar_todos_nesting).grid(row=7, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Importa", variable=var_importa_n).grid(row=8, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Exportar", variable=var_exportar_n).grid(row=8, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Limpar Lista", variable=var_limpar_lista_n).grid(row=8, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Chapas PDF", variable=var_relatorio_pdf_n).grid(row=9, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Relatorio PDF", variable=gerar_pdf_html_n).grid(row=9, column=1, sticky='w')
    
    tk.Label(checkbox_frame, text="BA-NA-NAS").grid(row=10, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Contar Chaps", variable=contar_chapa).grid(row=11, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Gerar Json", variable=gerar_json).grid(row=11, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Comp Vendedor", variable=compress_vend).grid(row=11, column=2, sticky='w')
    # tk.Checkbutton(checkbox_frame, text="ArrastarBanana", variable=copiar_pasta).grid(row=11, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Remover Lista", variable=remover_lista).grid(row=11, column=3, sticky='w')

    # Frame para organizar os botões
    frame_botoes = tk.Frame(janela)
    frame_botoes.pack(pady=2)

    # Botão para seleção manual de pastas
    botao_selecionar = tk.Button(frame_botoes, text="Selecionar Pasta Manualmente", command=selecionar_pasta_manual)
    botao_selecionar.pack(side=tk.LEFT, padx=10)

    # Botão OK para finalizar
    botao_ok = tk.Button(frame_botoes, text="OK", command=ok)
    botao_ok.pack(side=tk.LEFT, padx=10)

    # Criação do widget Text
    text_log = tk.Text(janela, height=150, width=80)
    text_log.pack(pady=10, padx=10)

    # Configuração de drag-and-drop
    janela.drop_target_register(DND_FILES)
    janela.dnd_bind('<<Drop>>', on_drop)

    # Iniciar a interface
    janela.mainloop()
    
    
if __name__ == '__main__':  
    main()