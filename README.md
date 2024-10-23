# AUTOGEN CERTIFICATE ULTIMATE V1.9

## Descrição

Este é um script Python que automatiza a geração de Certificados de Transporte com base em arquivos de Bipagem. Ele converte códigos EAN13 para nomes de produtos (códigos NX) e gera um documento PDF formatado pronto para ser enviado ao cliente ou arquivado.

O principal objetivo é reduzir o tempo e aumentar a eficiência no processo de geração de minutas de transporte, eliminando a necessidade de preenchimento manual.

## Funcionalidades

- **Conversão Automática**: Converte códigos EAN13 presentes nos arquivos de Bipagem para os códigos NX correspondentes.
- **Geração de Documentos**: Cria documentos Word baseados em um template e os converte em PDF.
- **Interface Gráfica**: Possui uma interface amigável usando Tkinter, permitindo arrastar e soltar arquivos.
- **Armazenamento Automatizado**: Salva automaticamente os PDFs gerados no diretório especificado para arquivamento ou envio.

## Pré-requisitos

Certifique-se de ter instalado os seguintes pacotes Python:

- Python 3.x
- `tkinter`
- `tkinterdnd2`
- `python-docx`
- `pywin32`

Para instalar os pacotes necessários, execute:

`bash`
`pip install tkinterdnd2 python-docx pywin32`

O script utiliza um arquivo Word como template localizado em:
`G:\Meu Drive\CONTROLLER\MINUTA\MODELO_MINUTA_TRANSPORTE_AUTOMATICA.docx`

Se necessário, ajuste o caminho do template no código para refletir a localização correta em seu sistema.

## Interface Gráfica:

- Nome do Cliente: Insira o nome do cliente para o qual a minuta será gerada.
- Nº Doc.Saída ou NF: Insira o número do pedido ou nota fiscal.
- Arraste os arquivos: Arraste e solte os arquivos de bipagem na área designada da interface.
- Gerar Minuta: Clique no botão "Gerar Minuta" para iniciar o processo.

## Processo Automático:

- Os arquivos serão convertidos, e os códigos EAN13 serão substituídos pelos códigos NX correspondentes.
- Um documento Word será gerado e preenchido com as informações fornecidas.
- O documento será convertido em PDF e salvo no diretório do cliente especificado.
- Uma mensagem de sucesso será exibida com o caminho do arquivo PDF gerado.

## Redefinir:

- Para iniciar um novo processo, clique no botão "Redefinir" para limpar os campos e remover os arquivos temporários.

## Estrutura do Projeto

- main.py: Arquivo principal contendo o código do script.
- BIPAGEM CONVERTIDA: Diretório gerado automaticamente para armazenar os arquivos de bipagem convertidos.

## Dependências
- tkinter: Biblioteca padrão para interfaces gráficas em Python.
- tkinterdnd2: Extensão para permitir arrastar e soltar na interface Tkinter.
- python-docx: Biblioteca para criar e atualizar arquivos Word.
- pywin32: Biblioteca para integração com APIs do Windows, necessária para converter Word em PDF.

## Personalização
Dicionário de Substituição:

O dicionário substituicoes_ean_para_nx contém as correspondências entre códigos EAN13 e códigos NX. Você pode atualizar este dicionário conforme necessário para refletir os códigos de produtos específicos da sua empresa.

Caminhos de Arquivos:
Ajuste os caminhos para o template do Word e o diretório base onde as pastas dos clientes estão localizadas, caso sejam diferentes em seu ambiente.

