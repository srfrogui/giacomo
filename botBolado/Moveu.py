import pyautogui as ag
import time
from PromobAuto import aguarde, procurar, clicar
import pyautogui as ag
import time
from PromobAuto import aguarde, procurar, clicar
import pyperclip
import threading
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
from tkcalendar import DateEntry

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
        
imagem = {
    "perfil": "./img/mov_perfil.png",
    "perfil2": "./img/mov_perfildois.png",
    "confirm": "./img/mov_confirm.png",
    "cut": "./img/mov_cut.png",
    "setinha": "./img/mov_aguardesetinha.png",
    "mais": "./img/mov_mais.png",
    "ref": "./img/mov_ref.png",
    "importar": "./img/mov_import.png",
    "mov_login": "./img/mov_login.png",
    "reftab": "./img/mov_refcarregatabela.png"
}

def mostrar_mensagem_erro(mensagem):
    # Show a Tkinter window to notify the user and wait for their response
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Atenção", f"{mensagem}. Clique OK para continuar.")
    root.destroy()

def extrair_nome(caminho):
    # Extrai o nome da última pasta do caminho
    nome_pasta = os.path.basename(caminho)
    return nome_pasta

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
    #ag.hotkey('ctrl', 'v')
    ag.press('enter') #dentro da pasta
    time.sleep(1)
    clicar('./img/salvar_salvar.png', ajusteX=-100)


arquivos_necessarios = ["PROJJE_FERRAGENS.CSV", "Projeto_producao.xls"]
def verificar_arquivos(pasta, arquivos):
    for arquivo in arquivos:
        if not os.path.isfile(os.path.join(pasta, arquivo)):
            print(f"Erro: Arquivo '{arquivo}' não encontrado na pasta '{pasta}'.")
            return False
    return True

def moveu(pasta, prazo, nomemov):
    if not verificar_arquivos(pasta, arquivos_necessarios):
        raise Exception(f"Certifique-se de que todos os {arquivos_necessarios} estão na pasta")
    ag.hotkey('win','r')
    ag.write('https://moveoecomobile.projje.com.br/#/pedidos')
    ag.press('enter')
    aguarde(imagem['mov_login'])
    clicar(imagem['mov_login'])
    
    aguarde(imagem["ref"])
    ag.press('tab')
    ag.press('enter')
    
    aguarde(imagem['importar'])
    clicar(imagem['importar'])
    
    time.sleep(1)
    ag.press('tab', 5)
    ag.press('enter')
    
    salvar(pasta, nome='"PROJJE_FERRAGENS.CSV" "Projeto_producao.xls"')
    aguarde(imagem['confirm'])
    
    clicar(imagem['ref'], ajusteY=160)
    
    time.sleep(3)
    clicar(imagem['mais']) #aperta no +     
    
    time.sleep(2)
    ag.press('tab', 3, interval=0.1) #aperta na caixinha
    ag.press('enter')

    ag.press('tab', 4)
    ag.write(prazo, interval=0.1)
    ag.press('tab')
    time.sleep(1)
    ag.write(nomemov)
    time.sleep(1)
    ag.press('tab')
    ag.press('enter')
    
    aguarde(imagem['cut'])
    time.sleep(1)
    clicar(imagem['cut'])
    salvar(pasta)
    
    time.sleep(2)
    ag.press('tab', 8) #vai para produzidos
    ag.press('enter')
    
    aguarde(imagem["setinha"])
    
    ag.press('tab', 5)
    ag.press('enter')
    
    ag.press('tab', 4)
    ag.press('enter')
    time.sleep(3)
    
    
    if procurar(imagem["perfil"]):
        ag.hotkey('ctrl', 'p')
        time.sleep(2)
        ag.press('enter')
        time.sleep(2)
        salvar(pasta+"\VENDEDOR", nome='Projje PCP - Requisições Ferragem')
        
    time.sleep(1)
    ag.hotkey('ctrl', 'w')

    ag.press('enter')
    ag.press('tab', 7)
    ag.press('enter')
    time.sleep(3)
    if procurar(imagem["perfil2"]):
        ag.hotkey('ctrl', 'p')
        time.sleep(2)
        ag.press('enter')
        salvar(pasta+"\VENDEDOR", nome='Projje PCP - Requisições Perfil')
    time.sleep(1)
    ag.hotkey('ctrl', 'w')    
    time.sleep(1)
    ag.hotkey('ctrl', 'w')
        

def log_message(message):
    if text_log:
        text_log.insert(tk.END, message + '\n')
        text_log.see(tk.END)

log_file_M = 'moveu.log'

def main():

    def toggle_chrome_inputs():
        if var_moveu.get():
            entry_vendedor.config(state='normal')
            entry_cliente.config(state='normal')
            date_prazo.config(state='normal')
        else:
            entry_vendedor.config(state='disabled')
            entry_cliente.config(state='disabled')
            date_prazo.config(state='disabled')   
    
    def ok():
        try:
            logging.info('Botão OK pressionado. Fechando a aplicação.')  # Fecha a janela e encerra o loop do tkinter
            if not pastas :
                messagebox.showinfo("Atenção", "Selecione uma pasta.")
                return
            log_message("Processo Moveu Iniciado...")
            with open(log_file_M, 'a') as log:
                log.write(f'Processando pasta: {pasta}\n')
            
            prazo = date_prazo.get_date().strftime("%d%m%Y")
            nomemov = entry_vendedor.get() + " - " + entry_cliente.get()
            print(prazo, '__', nomemov.upper())
            for pasta in pastas:
                threading.Thread(target=moveu, args=(pasta, prazo, nomemov,)).start()
                
            with open(log_file_M, 'a') as log:
                log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
                log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')
                
        except Exception as e:
            with open(log_file_M, 'a') as log:
                log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                mostrar_mensagem_erro("Aviso: Erro generico ao processar pasta {pasta}: {e}\n")
            
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
    
    # Frame para organizar os checkboxes horizontalmente
    checkbox_frame = tk.Frame(janela)
    checkbox_frame.pack(pady=3)
    
    global var_moveu
    var_moveu = IntVar(value=0)
    
    tk.Label(checkbox_frame, text="Chrome process").grid(row=2, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Moveu", variable=var_moveu, command=toggle_chrome_inputs).grid(row=2, column=1, sticky='w')
    # Campo de texto e data para o processo "Chrome"
    text_label = tk.Label(checkbox_frame, text="Vendedor:")
    text_label.grid(row=2, column=2, sticky='w', padx=1)
    entry_vendedor = tk.Entry(checkbox_frame)
    entry_vendedor.grid(row=2, column=3, sticky='w', padx=1)
    
    text_label = tk.Label(checkbox_frame, text="Cliente:")
    text_label.grid(row=3, column=0, sticky='w', padx=1)
    entry_cliente = tk.Entry(checkbox_frame)
    entry_cliente.grid(row=3, column=1, sticky='w', padx=1)

    date_label = tk.Label(checkbox_frame, text="Prazo:")
    date_label.grid(row=3, column=2, sticky='w', padx=10)
    date_prazo = DateEntry(checkbox_frame, width=17, background='darkblue', foreground='white', borderwidth=2)
    date_prazo.grid(row=3, column=3, sticky='w', padx=1)
    
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