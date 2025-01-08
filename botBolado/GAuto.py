import os
import shutil
import time
import glob
import tkinter as tk
import pyautogui as ag
from tkinter import filedialog, messagebox, IntVar, Frame, Button, Label
from tkinterdnd2 import TkinterDnD, DND_FILES
import re
import subprocess
import pytesseract
import threading

#PREPARARO ======================================================================================================
def extrair_nome(caminho):
    # Extrai o nome da última pasta do caminho
    nome_pasta = os.path.basename(caminho)
    return nome_pasta
    
def verificar_arquivos(pasta):
    # Corrigindo a utilização de f-string para garantir que {pasta} seja expandido corretamente
    caminho_cut = os.path.join(pasta, 'planoCorte_Moveo_Ecomobile_OP_*.xls')
    arquivos_cut = glob.glob(caminho_cut)  # Busca pelos arquivos

    if not arquivos_cut:
        print(f"Nenhum arquivo encontrado para: {caminho_cut}")
        raise ValueError(f"Nenhum arquivo encontrado para: {caminho_cut}")
    
    # Caminho para o arquivo Gplan
    caminho_gplan = os.path.join(pasta, 'Gplan', 'Projeto_lista_de_paineis.xls')
    
    if not os.path.isfile(caminho_gplan):
        print(f"Nenhum arquivo encontrado para: {caminho_gplan}")
        raise ValueError(f"Nenhum arquivo encontrado para: {caminho_gplan}")
    print('Tem arquivos planoCorte e Gplan!')
    
    return caminho_gplan, arquivos_cut[0]

def seleciona_tudo(caminho):
    ag.keyDown('shift')
    time.sleep(0.1)
    ag.press('tab')
    time.sleep(0.1)
    ag.keyUp('shift')
    time.sleep(0.1) 
    ag.press('tab')
    time.sleep(0.1)
    ag.write(caminho)
    time.sleep(0.1)

def mostrar_mensagem_erro(mensagem):
    # Show a Tkinter window to notify the user and wait for their response
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Atenção", f"{mensagem}. Clique OK para continuar.")
    root.destroy()

def vaievolta_imprimir(num):
    ag.keyDown('shift')
    ag.press('tab', presses=num)  # Pressiona 'Tab' duas vezes para voltar
    ag.keyUp('shift')
    ag.press('enter')  # Pressiona 'Enter' para confirmar
    time.sleep(1)

def garantir_pasta(caminho):
    # Verifica se o diretório 'VENDEDOR' existe e, se não, cria-o
    pasta_vendedor = os.path.join(caminho, 'VENDEDOR')
    if not os.path.exists(pasta_vendedor):
        os.makedirs(pasta_vendedor)
        print(f"Pasta '{pasta_vendedor}' criada.")
    else:
        print(f"Pasta '{pasta_vendedor}' já existe.")

#### A LIMPO

def aguarde(imagem, confianca=0.95, timeout=50, intervalo=1, inverter=False):
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

def procurar(imagem, confianca=0.95, limite=0.8):
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

def procurar_colorido(imagem, confianca=0.95, limite=0.8):
    print(f"Procurando imagem... {imagem} - {os.path.exists(imagem)}")
    #print(f"Caminho relativo: {os.path.relpath(imagem)}")
    
    while confianca >= limite:
        try:
            localizacao = ag.locateCenterOnScreen(imagem, confidence=confianca, grayscale=False)
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

def extrair_texto(image_reference_path, pasta, correction_x=0, correction_y=0, capture_width=200, capture_height=400):
    def create_overlay(region_x, region_y, width, height):
        print("Criando sobreposição (overlay) na tela...")  # Indica que a sobreposição está sendo criada
        # Cria uma janela transparente
        overlay = tk.Toplevel()
        overlay.attributes('-alpha', 0.3)  # Define a transparência (30%)
        overlay.attributes('-topmost', True)  # Sempre no topo
        overlay.overrideredirect(True)  # Remove as bordas da janela
        
        # Define a geometria (posição e tamanho) da janela
        overlay.geometry(f"{width}x{height}+{region_x}+{region_y}")
        
        # Cria um frame laranja ao redor da janela
        frame = tk.Frame(overlay, bg="orange", bd=5)
        frame.pack_propagate(0)
        frame.pack(fill=tk.BOTH, expand=True)

        # Fecha a janela após 4 segundos
        overlay.after(5000, overlay.destroy)
        # Executa `update_idletasks()` para atualizar a janela sem entrar no loop principal
        root = overlay.master
        root.update_idletasks()
        print("Sobreposição removida.")  # Indica que a sobreposição foi removida
        
    max_attempts = 2  # Definimos o número máximo de tentativas
    attempt = 0
    confirmacao = False
    
    while attempt < max_attempts:
        root = tk.Tk()
        root.withdraw()
        
        print(f"Iniciando o processo de captura e extração de texto para a imagem: {image_reference_path}")
        
        # Verifica se a imagem de referência existe
        if not os.path.exists(image_reference_path):
            print(f"Erro: {image_reference_path} não encontrado.")
            return "skip"

        print(f"Procurando o ponto de referência na tela baseado na imagem {image_reference_path}...")
        
        # Localiza o ponto de referência na tela baseado na imagem
        reference_point = ag.locateOnScreen(image_reference_path, confidence=0.6)

        if reference_point:
            x, y, w, h = reference_point  # Pega as coordenadas do ponto de referência
            print(f"Ponto de referência encontrado em: ({x}, {y})")
        else:
            print("Erro: Ponto de referência não encontrado")
            messagebox.showerror("Erro", "Ponto de referência não encontrado.")
            return "skip"

        # Aplica a correção nas coordenadas (x, y)
        corrected_x = x + correction_x
        corrected_y = y + correction_y

        print(f"Correção aplicada: x={corrected_x}, y={corrected_y}")

        # Coordenadas da área a ser capturada com o ajuste de largura e altura
        region_x = corrected_x + w // 2 - capture_width // 2
        region_y = corrected_y + h // 2 - capture_height // 2

        print(f"Área de captura definida nas coordenadas: x={region_x}, y={region_y}, largura={capture_width}, altura={capture_height}")

        print("Capturando a imagem da área definida...")

        # Converte os valores da região para inteiros
        region = (int(region_x), int(region_y), int(capture_width), int(capture_height))

        # Captura a área da tela ao redor do ponto de referência
        captured_image = ag.screenshot(region=region)
        
        # Cria a sobreposição na tela para mostrar o retângulo laranja
        #create_overlay(region_x, region_y, capture_width, capture_height)

        # Salva a imagem capturada para visualização, se necessário
        captured_image_path = pasta + "\paineis.png"
        captured_image.save(captured_image_path)
        print(f"Imagem capturada salva em: {captured_image_path}")

        print("Extraindo o texto da imagem capturada usando Tesseract...")
        # Extrai o texto da imagem capturada usando o Tesseract
        text = pytesseract.image_to_string(captured_image)

        print("Texto extraído da imagem:")
        print(text)

        # Mostra o texto extraído em uma mensagem de confirmação
        #confirmacao = messagebox.askyesno("Confirmação", f"O texto extraído foi:\n\n{text}\n\nIsso está correto?")
        print("Texto confirmado pelo usuário.")
        return text
        
        if confirmacao:
            print("Texto confirmado pelo usuário.")
            return text
        else:
            attempt += 1  # Incrementa o número de tentativas
            print(f"Tentativa {attempt}/{max_attempts}: Texto não confirmado. O usuário solicitou ajuste.")
            messagebox.showinfo("Sugestão", "Verifique a imagem de referência ou ajuste a área de captura.")
        
        root.destroy()

    # Após 2 tentativas falhas, retorna "skip"
    log_message("Número máximo de tentativas atingido. Retornando 'skip'.")
    return "skip"
    



#PROCESSO GPLAN ======================================================================================================
# Função para abrir o GPlan
def abrir_gplan():
    def is_gplan_running():
        # Executa o comando tasklist para verificar se o GPlan está rodando
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq GPlan.exe'], capture_output=True, text=True)
        return 'GPlan.exe' in result.stdout
    
    if is_gplan_running():
        print('GPlan já está aberto.')
    else:
        print('Abrindo GPlan...')
        subprocess.Popen(r'%USERPROFILE%\..\..\Giben\GPlan\GPlan.exe')
        time.sleep(11)
        print('GPlan foi aberto.')
    
def novo_projeto():
    print('Criando Novo Projeto...')
    clicar('./img/papel_branco.png')
    aguarde('./img/import_validacao.png')
    print('Projeto Criado')
    
def importa():
    clicar('./img/arquivo.png')
    aguarde('./img/import_exel.png')
    clicar('./img/import_exel.png')
    aguarde('./img/import_referencia.png')

# Função para importar o projeto
def importar_projeto(pasta, gplano, cuti):
    nome=extrair_nome(pasta)
    
    print(f'Importando projeto para a pasta {pasta}')
    
    # Simular cliques e automações aqui para preencher os campos
    clicar('./img/import_referencia.png', 60, 5) #painel
    seleciona_tudo(gplano)
    clicar('./img/import_referencia.png', 60, 50) #cut
    seleciona_tudo(cuti)
    clicar('./img/import_referencia.png', 60, 95) #input
    seleciona_tudo(nome)
    clicar('./img/import_referencia.png', 60, 120) #check
    clicar('./img/import_referencia.png', 580, 150) #Fim
    
    #espera a janela de importacao fechar para ver se carregou
    aguarde('./img/import_referencia.png', timeout=1000, inverter=True)
    time.sleep(5) #esperar um tempo depois da janela fechar (dlay pra aparecer o carregamento)
    
    if procurar('./img/import_carregamento.png'):
        aguarde('./img/import_carregamento.png', inverter=True)
        print('Carregar Sumiu')

    aguarde('./img/import_val.png', timeout=1000, inverter=True)
    
    if procurar('./img/import_validacao.png') is None:
        print("Validacao e Carregamento não encontrado. Refazendo operação com apenas o input.")
        # Refazendo a operação com apenas clicar no input e inserir o nome
        importa()
        clicar('./img/import_referencia.png', 60, 95) #input
        seleciona_tudo(nome)
        clicar('./img/import_referencia.png', 580, 150) #Fim
        
    aguarde('./img/import_validacao.png', timeout=1000)
    print("Importação bem-sucedida")          

# Função para abrir a produção
def abrir_parametro():
    # Simular clique na engrenagem
    print("Abrindo parametro...")
    clicar('./img/parametro.png')
    time.sleep(1)
    
# Função para configurar a otimização
def configurar_optimizacao():
    print("Configurando otimização...")
    clicar('./img/exit.png', -85) # SetinhaGVISION
    time.sleep(1)
    clicar('./img/exit.png', -85, 70) # GVISION
    clicar('./img/exit.png', -600, 100) # XZ
    clicar('./img/exit.png', -200, 170) # Ambos
    clicar('./img/exit.png', -655, 340 + 10) # 5, 5, 5, 5
    for i in range(4):
        ag.press('5')
        ag.press('tab')
    clicar('./img/embaralhador.png')
    print('Optimizacao Configurada')
    aguarde('./img/valida_embaralhamento.png', timeout=500, intervalo=3)

def conferir_impressora():
    aguarde('./img/aguarde_pdf.png') 
    while True:
        if procurar('./img/imprimir_PDF.png', confianca=1, limite=0.98) is None:
            ag.press('down')
        else:
            print('IMpressora PDF selecionado!')
            ag.press('enter')
            break
    
def imprimir_loop(caminho):
            
    def text_paineis(texto):
        # Lista para armazenar os painéis
        paineis = []
        
        # Regex para encontrar as informações no formato desejado
        # O padrão agora aceita materiais que têm espaços e captura o formato de espessura corretamente
        padrao = r'(\d+); Material ([\w\s]+(?:_COMPOSTA)?); Espessura ([\d,]+);?(?: Cor [A-Z]+)?'
        
        # Encontrar todas as correspondências
        matches = re.findall(padrao, texto)
        
        # Iterar sobre as correspondências e adicionar à lista
        for match in matches:
            numero, material, espessura = match
            painel = {
                'numero': int(numero),
                'material': material.strip(),  # Remove espaços em branco
                'espessura': float(espessura.replace(',', '.'))  # Converte espessura para float
            }
            
            paineis.append(painel)
            
        return paineis
    
    
    #tira os ,00 da espessura    
    texto = extrair_texto('./img/peck_fdt.png', caminho, correction_x=227, correction_y=305, capture_height=670, capture_width=240)
    if texto == 'skip':
        # Abre uma caixa de mensagem aguardando operação manual
        messagebox.showinfo("Ação Necessária", "Processo manual necessário. O processo foi interrompido.")
        print("Processo manual necessário. O processo foi interrompido.")
    else:
        paineis = text_paineis(texto)
        for painel in paineis:
            print(painel)
        garantir_pasta(caminho)
        caminho_doleite = caminho + r'\VENDEDOR'
        clicar('./img/peck_fdt.png', 95, -270) #clicar peck_fdt
        print(f"Quantidade de painéis: {len(paineis)}")
        
        for painel in paineis:
            espessura = int(painel['espessura']) if painel['espessura'].is_integer() else painel['espessura'] 
            print(f"Painel {painel['numero']}: Material {painel['material']}, Espessura {espessura}")
            
            ag.press('down')
            time.sleep(0.3)
            if procurar_colorido('./img/erro.png'):
                with open(log_file_G, 'a') as log:
                    log.write(f'DEU ERRO AE ERRO ERRADO DO ERRO VERMELHO!!! \n')
                    log_message(f'DEU ERRO AE ERRO ERRADO DO ERRO VERMELHO!!! \n')
                mostrar_mensagem_erro("Aviso: DEU ERRO AE ERRO ERRADO DO ERRO VERMELHO!!! \n")
                
            clicar('./img/exit.png', -300) #clicar imprimir esquemas
            aguarde('./img/imprimir_loop.png') 
            clicar('./img/imprimir_loop.png', 150) #clicar imprimir
            conferir_impressora() #seleciona PDF
            aguarde('./img/busca_ref.png')
            #add nome
            ag.write(f"zMDF - {painel['numero']} - Mat {painel['material']} - {espessura}MM")
            #add caminho
            clicar('./img/busca_ref.png')
            vaievolta_imprimir(2)
            ag.write(caminho_doleite)
            ag.press('enter') #aplica o caminho
            time.sleep(1)
            clicar('./img/salvarr.png')
            ag.hotkey('alt', 'f4')
            time.sleep(1)

        print('Impressao concluida...')
    
def gerar_gvision(pasta):
    nome = extrair_nome(pasta)
    print(f'Gerando GVision e copiando para a pasta {pasta}')
    print(f'{nome}.fdt ---> {pasta}')
    
    clicar('./img/exit.png', -100)
    aguarde('./img/cortepromob.png')
    ag.hotkey('alt', 'f4')
    shutil.copy(f'X:\\CORTE PROMOB\\{nome}.fdt', pasta)
    print(f'{nome}.fdt ---> {pasta}')
    
# Função para abrir a produção
def abrir_producao():
    # Simular clique na engrenagem
    print("Abrindo produção...")
    clicar('./img/exit.png', -40)
    clicar('./img/peck_fdt.png') #clicar producao
    aguarde('./img/ta_aberto.png') #aguarda carregamento

log_file_G = 'gplano.log'

def processar_pastas_gplan(pasta):

    # Abrir o GPlan uma vez
    #abrir_gplan() ## quando abre o gplan por aqui por algum motivo ele n consegue gerar_gvision
    gplano, cuti = verificar_arquivos(pasta)
    with open(log_file_G, 'a') as log:
        log.write(f'Processando pasta: {pasta}\n')
        
    # Etapas do processamento
    if var_novo_projeto.get():
        log_message("Iniciando NovoProjeto...")
        novo_projeto()
    if var_importa.get():
        log_message("Iniciando Importacao 1/2...")
        importa()
    if var_importar_projeto.get():
        log_message("Importando Projeto 2/2...")
        importar_projeto(pasta, gplano, cuti)
    if var_abrir_parametro.get():
        abrir_parametro()
    if var_configurar_otimizacao.get():
        log_message("Configurando optimizacao...")
        configurar_optimizacao()
    if var_imprimir_loop.get():
        log_message("Gerando PDF's...")
        imprimir_loop(pasta)
    if var_gerar_gvision.get():
        log_message("Gerando GVision...")
        gerar_gvision(pasta)
    if var_abrir_producao.get():
        log_message("Retornando Loop...")
        abrir_producao()
    
    with open(log_file_G, 'a') as log:
        log.write(f'Processamento da {extrair_nome(pasta)}: concluido \n')
        log_message(f'Processamento da {extrair_nome(pasta)}: concluido \n')

def log_message(message):
    if text_log:
        text_log.insert(tk.END, message + '\n')
        text_log.see(tk.END)

def main():
    def adicionar_pasta_interface(pasta):
        # Extrair o nome da pasta
        nome_pasta = pasta.split('/')[-1]
        pasta_label = tk.Label(frame_pastas, text=nome_pasta, font=("Arial", 12), relief=tk.RAISED)

        # Evento para mostrar o caminho completo ao passar o mouse sobre o nome
        pasta_label.bind("<Enter>", lambda event, p=pasta: caminho_label.config(text=p))
        pasta_label.pack(pady=2)

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
            
    def ok():
        if not pastas:
            messagebox.showinfo("Atenção", "Selecione uma pasta.")
            return
        for pasta in pastas:       
            try:
                threading.Thread(target=processar_pastas_gplan, args=(pasta,)).start()
                log_message("Processo Iniciado...")
                    
            except ValueError as e:  # Captura especificamente o erro lançado na função clicar
                with open(log_file_G, 'a') as log:
                    log.write(f'Erro ao processar pasta {pasta}: {e}\n')
                    log_message(f'Erro ao processar pasta {pasta}: {e}\n')
                    mostrar_mensagem_erro("Aviso: Erro ao processar pasta {pasta}: {e}\n")
                break
            except Exception as e:
                with open(log_file_G, 'a') as log:
                    log.write(f'Erro generico ao processar pasta {pasta}: {e}\n')
                    mostrar_mensagem_erro("Aviso: Erro generico ao processar pasta {pasta}: {e}\n")
                break

        
    def on_close(): 
        janela.destroy()
        os._exit(0)
    
    global caminho_label, frame_pastas, pastas, text_log
    pastas = []
    # Criar janela principal
    global janela
    janela = TkinterDnD.Tk()
    janela.title("Seleção de Pastas")
    janela.geometry("600x400")

    global var_novo_projeto, var_importa, var_importar_projeto, var_abrir_parametro, var_configurar_otimizacao, var_imprimir_loop, var_gerar_gvision, var_abrir_producao
    # Variáveis para checkboxes (definidas após criar a janela)
    var_novo_projeto = IntVar(value=1)
    var_importa = IntVar(value=1)
    var_importar_projeto = IntVar(value=1)
    var_abrir_parametro = IntVar(value=1)
    var_configurar_otimizacao = IntVar(value=1)
    var_imprimir_loop = IntVar(value=1)
    var_gerar_gvision = IntVar(value=1)
    var_abrir_producao = IntVar(value=0)

    janela.protocol("WM_DELETE_WINDOW", on_close)

    # Frame para exibir as pastas
    frame_pastas = tk.Frame(janela)
    frame_pastas.pack(pady=20, padx=10)

    # Label para exibir o caminho completo
    caminho_label = tk.Label(janela, text="", fg="blue")
    caminho_label.pack(pady=5)

    # Adicionar checkboxes em um Frame
    checkbox_frame = tk.Frame(janela)
    checkbox_frame.pack(pady=10)

    tk.Checkbutton(checkbox_frame, text="Novo Projeto", variable=var_novo_projeto).grid(row=0, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Importa", variable=var_importa).grid(row=0, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Importar Projeto", variable=var_importar_projeto).grid(row=0, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Abrir Parâmetro", variable=var_abrir_parametro).grid(row=0, column=3, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Configurar Otimização", variable=var_configurar_otimizacao).grid(row=1, column=0, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Imprimir Loop", variable=var_imprimir_loop).grid(row=1, column=1, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Gerar GVision", variable=var_gerar_gvision).grid(row=1, column=2, sticky='w')
    tk.Checkbutton(checkbox_frame, text="Abrir Produção", variable=var_abrir_producao).grid(row=1, column=3, sticky='w')

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


# Início do script
if __name__ == '__main__':  
    main()
