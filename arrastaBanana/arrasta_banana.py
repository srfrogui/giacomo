import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys

def load_categories(filename):
    categorias = {}
    if os.path.exists(filename):
        with open(filename, 'r') as file:
            for line in file:
                if line.startswith('#'):
                    continue
                key, value = line.strip().split(',')
                categorias[key] = value
    else:
        with open(filename, 'w') as file:
            file.write("# Adicione aqui as categorias\n")
            file.write('_org,organizacao\n')  # Exemplo de categoria padrão
        log_message(f"Arquivo de categorias \"{filename}\" criado. Adicione suas categorias.")
    return categorias

def select_folder():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    folder = filedialog.askdirectory(title="Selecione uma pasta")
    return folder

def get_backup_category(nome_pasta, categorias):
    for key, value in categorias.items():
        if key in nome_pasta:
            return value
    return "SEMCATEGORIA"

def log_message(message):
    if text_log:
        text_log.insert(tk.END, message + '\n')
        text_log.see(tk.END)

def copy_with_progress(src, dst):
    total_files = sum([len(files) for r, d, files in os.walk(src)])
    copied_files = 0
    
    for dirpath, dirnames, filenames in os.walk(src):
        relative_path = os.path.relpath(dirpath, src)
        destination_path = os.path.join(dst, relative_path)
        os.makedirs(destination_path, exist_ok=True)

        for filename in filenames:
            source_file = os.path.join(dirpath, filename)
            shutil.copy2(source_file, destination_path)
            copied_files += 1
            
            progress = (copied_files / total_files) * 100
            progress_var.set(progress)
            progress_label.config(text=f'Copiando... {copied_files}/{total_files} arquivos')
            progress_label.update()

    progress_var.set(100)
    progress_label.config(text='Cópia concluída!')

def prepare_destination_giben(nome_pasta, folder, furacao_path):
    source_giben = os.path.join(folder, "Giben")
    destination_giben = os.path.join(furacao_path, nome_pasta, "Giben")

    if not os.path.exists(source_giben):
        log_message(f"A pasta \"Giben\" não foi encontrada em \"{folder}\". Backup cancelado.")
        return False

    os.makedirs(os.path.dirname(destination_giben), exist_ok=True)
    log_message("Copiando arquivos de furação...")
    copy_with_progress(source_giben, destination_giben)
    return True
    
def prepare_destination_nesting(nome_pasta, folder, corte_path):
    # Check if 'Nesting' folder exists within the 'folder'
    source_nesting = os.path.join(folder, "Nesting")
    
    # Proceed only if the 'Nesting' folder exists
    if os.path.exists(source_nesting) and os.path.isdir(source_nesting):
        destination_nesting = os.path.join(corte_path, nome_pasta, "Nesting")
        
        # Create the destination directory if it doesn't exist
        os.makedirs(os.path.dirname(destination_nesting), exist_ok=True)
        
        # Log message and copy the files with progress
        log_message("Copiando arquivos de Nesting...")
        copy_with_progress(source_nesting, destination_nesting)
    else:
        log_message("A pasta 'Nesting' não foi encontrada, pulando a cópia.")

def prepare_destination_img(nome_pasta, img_path, etiqueta_path): 
    nome_pasta = os.path.join(nome_pasta, "Gplan") #mefudecomatualizacao
    source_img = os.path.join(img_path, nome_pasta)
    destination_etiqueta = os.path.join(etiqueta_path, nome_pasta)

    if not os.path.exists(source_img):
        log_message("A pasta ETIQUETA LOCAL não foi encontrada. Backup cancelado.")
        return False

    os.makedirs(destination_etiqueta, exist_ok=True)
    log_message("Copiando Etiquetas...")
    copy_with_progress(source_img, destination_etiqueta)
    return True

def backup_folder(folder, nome_pasta, backup_path, categoria):
    final_backup_path = os.path.join(backup_path, categoria, nome_pasta)
    os.makedirs(final_backup_path, exist_ok=True)
    log_message("Iniciando Backup...")
    copy_with_progress(folder, final_backup_path)
    log_message("Backup Completo!")
    root.after(2000, close_window)  # Aguarda 2 segundos e fecha a janela

def aceitar_vendedor(folder, aceite_vendas):
    folder_vendedor = os.path.join(folder, "VENDEDOR")
    # Lista todos os arquivos da pasta
    for nome_arquivo in os.listdir(folder_vendedor):
        if nome_arquivo.endswith(".vendas"):
            caminho_origem = os.path.join(folder_vendedor, nome_arquivo)
            caminho_destino = os.path.join(aceite_vendas, nome_arquivo)

            log_message("Copiando .vendas...")
            shutil.copy2(caminho_origem, caminho_destino)
            return True

    log_message("O arquivo .vendas não foi encontrado.")
    # return False
    return True

def start_backup(folder, nome_pasta, backup_path, furacao_path, img_path, etiqueta_path, categorias, nesting_path, aceite_vendas):
    categoria = get_backup_category(nome_pasta, categorias)
    log_message(f"Nome da pasta: {nome_pasta}")
    log_message(f"Iniciando Backup... {categoria}")
    
    prepare_destination_nesting(nome_pasta, folder, nesting_path)
    
    if prepare_destination_giben(nome_pasta, folder, furacao_path):
        if prepare_destination_img(nome_pasta, img_path, etiqueta_path):
            if aceitar_vendedor(folder, aceite_vendas):
                backup_folder(folder, nome_pasta, backup_path, categoria)

def close_window():
    root.quit()  # Fecha a janela Tkinter
    os._exit(0)  # Encerra o programa completamente

def main():
    global progress_var, progress_label, text_log, root

    backup_path = r'Z:\BACKUP ALEXANDRE\PROJETOS PROMOB\PROMOB PROJJE\OMAR'
    img_path = r'C:\Giben\GvisionXPPROMOB\CNC\Media\Img'
    categorias_file = 'categoria.txt'
    furacao_path = r'Z:\PRODUCAO\FURAÇÃO PROMOB'
    etiqueta_path = r'Z:\PRODUCAO\ETIQUETA GVISION GPLAN PROMOB'
    nesting_path= r'Z:\PRODUCAO\CORTE G2 NESTING'
    aceite_vendas= r'Z:\VENDAS\ACEITE-VENDAS'
    
    root = tk.Tk()
    root.title("Backup de Pastas")
    root.geometry("400x300")

    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.pack(pady=20, padx=20, fill=tk.X)

    progress_label = tk.Label(root, text="")
    progress_label.pack(pady=5)

    text_log = tk.Text(root, height=10, width=50)
    text_log.pack(pady=10, padx=10)

    categorias = load_categories(categorias_file)

    # Verifica se um caminho foi passado como argumento
    if len(sys.argv) > 1:
        folder = sys.argv[1]  # Pega o primeiro argumento
    else:
        folder = select_folder()  # Se não, usa o seletor de pastas

    if not os.path.exists(folder):
        messagebox.showwarning("Aviso", "O caminho selecionado não existe.")
        return

    nome_pasta = os.path.basename(folder)

    root.protocol("WM_DELETE_WINDOW", close_window)

    start_backup(folder, nome_pasta, backup_path, furacao_path, img_path, etiqueta_path, categorias, nesting_path, aceite_vendas)

    root.mainloop()

if __name__ == "__main__":
    main()
