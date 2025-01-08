import tkinter as tk
from tkinter import IntVar, Tk
from tkinterdnd2 import TkinterDnD  # Para drag-and-drop, se necessário

# Função que será chamada quando a seleção do OptionMenu for alterada
def mudar_processamento(*args):
    escolha = var_processamento.get()
    if escolha == "Nesting":
        label_gplan_nesting.config(text="Nesting process")
        alternar_checkboxes(checkboxes_gplan, checkboxes_nesting, checkboxes_gplan_vars, checkboxes_nesting_vars)
    else:
        label_gplan_nesting.config(text="GPlan process")
        alternar_checkboxes(checkboxes_nesting, checkboxes_gplan, checkboxes_nesting_vars, checkboxes_gplan_vars)

# Criar janela principal
janela = TkinterDnD.Tk()
janela.title("Seleção de Pastas")
janela.geometry("600x400")

# Variáveis para checkboxes
var_selecionar_todos_gplan = IntVar(value=0)
var_novo_projeto = IntVar(value=0)
var_importa = IntVar(value=0)
var_importar_projeto = IntVar(value=0)
var_abrir_parametro = IntVar(value=0)
var_configurar_otimizacao = IntVar(value=0)
var_imprimir_loop = IntVar(value=0)
var_gerar_gvision = IntVar(value=0)
var_abrir_producao = IntVar(value=0)

# Variáveis para checkboxes de Nesting
var_selecionar_todos_nesting = IntVar(value=0)
var_novo_projeto_nesting = IntVar(value=0)
var_importa_nesting = IntVar(value=0)
var_importar_projeto_nesting = IntVar(value=0)
var_abrir_parametro_nesting = IntVar(value=0)
var_configurar_otimizacao_nesting = IntVar(value=0)
var_imprimir_loop_nesting = IntVar(value=0)
var_gerar_gvision_nesting = IntVar(value=0)
var_abrir_producao_nesting = IntVar(value=0)

# Variável para seleção do tipo de processo (GPlan ou Nesting)
var_processamento = tk.StringVar(value="GPlan")  # GPlan é a opção inicial
var_processamento.trace_add("write", mudar_processamento)  # Vincula a função de mudança

# Frame para exibir as pastas
frame_pastas = tk.Frame(janela)
frame_pastas.pack(pady=20, padx=10)

# Label para exibir o caminho completo
caminho_label = tk.Label(janela, text="", fg="blue")
caminho_label.pack(pady=5)

# Adicionar checkboxes em um Frame
checkbox_frame = tk.Frame(janela)
checkbox_frame.pack(pady=15)

# Label para o Menu de Seleção
label_selecao = tk.Label(checkbox_frame, text="Selecione o tipo de processo:")
label_selecao.grid(row=0, column=0, sticky='w')

# Menu de seleção (OptionMenu) para GPlan ou Nesting
optionmenu = tk.OptionMenu(checkbox_frame, var_processamento, "GPlan", "Nesting")
optionmenu.grid(row=0, column=1, sticky='w')

# Label para alternar entre GPlan e Nesting
label_gplan_nesting = tk.Label(checkbox_frame, text="GPlan process")
label_gplan_nesting.grid(row=1, column=0, sticky='w')

# Checkboxes de GPlan
checkboxes_gplan = [
    tk.Checkbutton(checkbox_frame, text="Selecionar Todes", variable=var_selecionar_todos_gplan),
    tk.Checkbutton(checkbox_frame, text="Novo Projeto", variable=var_novo_projeto),
    tk.Checkbutton(checkbox_frame, text="Importa", variable=var_importa),
    tk.Checkbutton(checkbox_frame, text="Importar Projeto", variable=var_importar_projeto),
    tk.Checkbutton(checkbox_frame, text="Abrir Parâmetro", variable=var_abrir_parametro),
    tk.Checkbutton(checkbox_frame, text="Configurar Otimização", variable=var_configurar_otimizacao),
    tk.Checkbutton(checkbox_frame, text="Imprimir Loop", variable=var_imprimir_loop),
    tk.Checkbutton(checkbox_frame, text="Gerar GVision", variable=var_gerar_gvision),
    tk.Checkbutton(checkbox_frame, text="Abrir Produção", variable=var_abrir_producao)
]

# Checkboxes de Nesting
checkboxes_nesting = [
    tk.Checkbutton(checkbox_frame, text="Selecionar Todes", variable=var_selecionar_todos_nesting),
    tk.Checkbutton(checkbox_frame, text="Novo Pau", variable=var_novo_projeto_nesting),
    tk.Checkbutton(checkbox_frame, text="Importa nesting", variable=var_importa_nesting),
    tk.Checkbutton(checkbox_frame, text="Importar nesjeto", variable=var_importar_projeto_nesting),
    tk.Checkbutton(checkbox_frame, text="Abrir  nesting Parâmetro", variable=var_abrir_parametro_nesting),
    tk.Checkbutton(checkbox_frame, text="Configurar O nesting timização", variable=var_configurar_otimizacao_nesting),
    tk.Checkbutton(checkbox_frame, text="Imprimir Loop", variable=var_imprimir_loop_nesting),
    tk.Checkbutton(checkbox_frame, text="Gerar GVision", variable=var_gerar_gvision_nesting),
    tk.Checkbutton(checkbox_frame, text="Abrir Produção", variable=var_abrir_producao_nesting)
]

# Lista de variáveis associadas aos checkboxes
checkboxes_gplan_vars = [
    var_selecionar_todos_gplan,
    var_novo_projeto,
    var_importa,
    var_importar_projeto,
    var_abrir_parametro,
    var_configurar_otimizacao,
    var_imprimir_loop,
    var_gerar_gvision,
    var_abrir_producao
]

checkboxes_nesting_vars = [
    var_selecionar_todos_nesting,
    var_novo_projeto_nesting,
    var_importa_nesting,
    var_importar_projeto_nesting,
    var_abrir_parametro_nesting,
    var_configurar_otimizacao_nesting,
    var_imprimir_loop_nesting,
    var_gerar_gvision_nesting,
    var_abrir_producao_nesting
]

def alternar_checkboxes(ocultar, mostrar, ocultar_vars, mostrar_vars):
    """Esconde os checkboxes de 'ocultar' e mostra os checkboxes de 'mostrar' em duas linhas."""
    
    # Esconder checkboxes de um processo
    for checkbox in ocultar:
        checkbox.grid_forget()
    # Desmarcar variáveis associadas aos checkboxes
    for var in ocultar_vars:
        var.set(0)
    
    # Organizar os checkboxes nas duas linhas
    for i, checkbox in enumerate(mostrar):
        # Calcular linha e coluna para distribuição em duas linhas
        row = (i // 5)+1  # A linha será 1 para os primeiros 5 itens, 2 para os próximos
        col = i % 5         # A coluna será de 0 a 4 (5 colunas por linha)
        checkbox.grid(row=row, column=col, sticky='w')  # Organiza na linha e coluna calculadas

# Iniciar com o processo GPlan visível
alternar_checkboxes(checkboxes_nesting, checkboxes_gplan, checkboxes_nesting_vars, checkboxes_gplan_vars)

# Iniciar a interface
janela.mainloop()
