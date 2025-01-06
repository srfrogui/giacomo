# Gerador de PDF de Projeto Produção

Este projeto é um aplicativo Python que gera um arquivo PDF a partir de dados filtrados de um arquivo Excel. O aplicativo permite que os usuários selecionem uma pasta de projeto, escolham uma coluna para filtrar os dados e especifiquem os valores de filtro. Os dados filtrados são então utilizados para criar um PDF que inclui informações relevantes e imagens associadas.

## Funcionalidades

- Seleção de pasta do projeto.
- Filtragem de dados com base em colunas específicas.
- Geração de um arquivo PDF com os dados filtrados.
- Inclusão de imagens associadas aos dados no PDF.

## Dependências

Este aplicativo requer as seguintes bibliotecas Python:

- `reportlab`: Para a geração de PDFs.
- `PIL` (Pillow): Para manipulação de imagens.
- `tkinter`: Para a interface gráfica do usuário.
- `pandas`: Para manipulação de dados em tabelas.
- `pywin32`: Para interação com o Excel.

Você pode instalar as dependências necessárias usando o seguinte comando:

```bash
pip install reportlab Pillow pandas pywin32
```

## Como Usar

# Opção 1: Executar o script diretamente

 - Execute o script GERARELATORIO_FALTANTES.py.
 - Na interface gráfica, clique em "Selecionar" para escolher a pasta do projeto que contém os arquivos Excel.
 - Selecione a coluna que deseja usar para filtrar os dados.
 - Digite os valores de filtro, separados por vírgula.
 - Clique em "Gerar PDF" para criar o arquivo PDF com os dados filtrados.
   
# Opção 2: Utilizar o release pré-compilado

 - Acesse a página de **[releases](https://github.com/giacomo/releases)** do projeto e baixe o arquivo *GeradorRelatorioFaltanteV1.0.zip*.
 - Extraia o conteúdo do arquivo e execute o programa *GERARELATORIO_FALTANTES.exe*.
 - Siga as mesmas etapas mencionadas acima para gerar o PDF.

## Estrutura do Código

O código principal está contido na classe `GeradorPDFApp`, que gerencia a interface do usuário e a lógica para gerar o PDF. As principais funções incluem:

- `selecionar_pasta()`: Permite ao usuário selecionar uma pasta de projeto.
- `gerar_pdf()`: Lê os dados do Excel, aplica os filtros e gera o PDF.
- `criar_pdf()`: Cria o PDF a partir dos dados filtrados.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir um problema ou enviar um pull request.

## Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo `LICENSE` para mais detalhes.
