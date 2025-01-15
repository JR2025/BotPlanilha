Projeto: Extração de Dados de PDFs

Este projeto realiza a extração de dados específicos de arquivos PDF e gera uma planilha com as informações extraídas. Foi desenvolvido para facilitar o processamento de guias de imposto fiscal.

Estrutura do Projeto

front.py: Interface gráfica criada com Tkinter para permitir a seleção da pasta de PDFs e iniciar o processo.

extrairDados.py: Responsável pela extração de dados dos PDFs, incluindo CNPJ, NF-e, Total a Recolher e Linha Digitável.

criaPlanilha.py: Gera uma planilha Excel com os dados extraídos.

caminho_pasta.txt: Arquivo temporário que armazena o caminho da pasta selecionada pelo usuário.

pdfs: Diretório para armazenar os PDFs a serem processados.

planilhas: Pasta onde as planilhas geradas serão salvas.

Como Usar

Requisitos

Certifique-se de ter o seguinte instalado no seu sistema:

Python 3.8 ou superior

Bibliotecas: tkinter, pandas, openpyxl, PyPDF2

Instale as dependências com o seguinte comando:

pip install pandas openpyxl PyPDF2

Execução

Abra o terminal e navegue até o diretório do projeto.

Execute o arquivo front.py com o comando:

python front.py

Use a interface para:

Selecionar a pasta que contém os PDFs.

Clicar no botão "Inserir" para iniciar o processo de extração.

Clicar no botão "Limpar" para resetar o caminho da pasta.

As planilhas serão salvas no diretório planilhas com os dados processados.
