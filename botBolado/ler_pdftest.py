import PyPDF2

def ler_pdf(caminho_arquivo):
    # Abre o arquivo PDF no modo leitura binária
    with open(caminho_arquivo, 'rb') as arquivo_pdf:
        # Cria um objeto PDF reader
        leitor = PyPDF2.PdfReader(arquivo_pdf)
        
        # Inicializa uma variável para armazenar o conteúdo
        conteudo = ""
        
        # Itera sobre todas as páginas do PDF
        for pagina in range(len(leitor.pages)):
            # Obtém o conteúdo de cada página
            pagina_pdf = leitor.pages[pagina]
            conteudo += pagina_pdf.extract_text()  # Extrai o texto da página

    return conteudo



import pdfplumber

def ler_pdfa(caminho_arquivo):
    # Abre o arquivo PDF usando pdfplumber
    with pdfplumber.open(caminho_arquivo) as pdf:
        conteudo = ""
        
        # Itera sobre todas as páginas
        for pagina in pdf.pages:
            # Extrai o texto da página
            conteudo += pagina.extract_text()
    
    return conteudo
# Caminho do arquivo PDF
caminho_arquivo = 'C:\Giben\GvisionXPPROMOB\CNC\Media\Img\PROJJE_DEC SHOWROOM DEC LOJA\VENDEDOR\ListagemCompleta.pdf'

# Chama a função para ler o PDF e exibe o conteúdo
texto_pdf = ler_pdf(caminho_arquivo)
print(texto_pdf)
print("")
print("")
texto_1 = ler_pdfa(caminho_arquivo)
print(texto_1)