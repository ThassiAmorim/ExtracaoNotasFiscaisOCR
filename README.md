# Projeto de Extração e Processamento de Dados em PDFs

## Visão Geral
Este projeto utiliza Python para extrair imagens de romaneios digitalizados em PDFs, aplicar OCR (reconhecimento óptico de caracteres) do Google Cloud Vision para extrair texto das imagens e, em seguida, manipular e salvar os dados em uma planilha Excel. É ideal para processar documentos como notas fiscais e gerar relatórios consolidados.

## Funcionalidades
- **Extração de texto e imagens de PDFs** usando `PyMuPDF (fitz)` e `PIL`.
- **Aplicação de OCR** nas imagens extraídas dos PDFs utilizando a API do Google Cloud Vision.
- **Processamento e validação de dados extraídos**, como valores monetários e nomes de municípios e estabelecimentos.
- **Comparação de similaridade** para encontrar a melhor correspondência entre os dados extraídos e um arquivo de gabarito.
- **Atualização de uma planilha Excel** com os dados extraídos e processados.
- **Geração de arquivos de PDF** com páginas que não puderam ser processadas corretamente.

## Pré-requisitos

1. **Python 3.8+**
2. **Bibliotecas Python** (instale usando `pip install <biblioteca>`)
   - `PyMuPDF (fitz)`
   - `google-cloud-vision`
   - `Pillow`
   - `pandas`
   - `openpyxl`
   - `shutil`
   - `difflib`
3. **Configuração da API Google Cloud Vision**
   - Crie um projeto no [Google Cloud Platform](https://cloud.google.com/vision).
   - Habilite a API Google Cloud Vision.
   - Baixe e salve o arquivo de credenciais como `intropythoncloudvision.json`.
   - Defina a variável de ambiente `GOOGLE_APPLICATION_CREDENTIALS` para apontar para esse arquivo:
     ```bash
     export GOOGLE_APPLICATION_CREDENTIALS="caminho/para/intropythoncloudvision.json"
     ```

## Estrutura do Projeto

- **pdf_to_text(pdf_path)**: Extrai o texto de cada página de um arquivo PDF (Para o PDF gerado no sistema de lançamentos das notas fiscais).
- **extract_image_from_pdf(pdf_path)**: Extrai imagens de um PDF e corta para remover carimbos e anotações.
- **image_to_text(image_bytes)**: Aplica OCR às imagens extraídas usando a API Google Cloud Vision.
- **convert_to_float(value)**: Converte strings de valores monetários para o formato numérico float.
- **scrapping_data(text)**: Realiza extração de dados (município, estabelecimento e valor) do texto processado.
- **most_similar(targets, candidates)**: Retorna o item mais semelhante entre duas listas de strings usando `difflib`.
- **save_romaneios(pdf, page_num)**: Salva uma página de PDF que não foi processada corretamente.
- **extract_romaneios(pdf_path, new_path)**: Processa romaneios, encontrando municípios, estabelecimentos e valores e atualiza a planilha Excel.
- **extract_gerenciais(gerenciais_path, excel_path)**: Extrai dados de consultas gerenciais e os salva na planilha.
- **drop_empty_lines(excel_path)**: Remove linhas vazias do Excel.
- **create_excel(romaneios_path, gerenciais_path, excel_path, new_path)**: Executa o fluxo completo de extração e atualização de dados.

## Como Usar

1. Coloque seus arquivos PDFs (`romaneios.pdf` e `ConsultasGerenciais.pdf`) e seu arquivo Excel (`gabarito.xlsx`) no diretório do projeto.
2. Execute o script:
   ```bash
   python main.py
   ```
3. O arquivo `resultados_atualizados.xlsx` será gerado, contendo os dados extraídos e formatados.

## Notas
- O script tenta combinar os municípios e estabelecimentos com base na similaridade das strings. Caso a correspondência seja inferior a um limite definido, a página do PDF será salva em um diretório separado para revisão manual.
- Certifique-se de configurar corretamente as credenciais da API Google Cloud Vision para que o OCR funcione corretamente.
- Os romaneios usados estão no formato:
  ![romaneio](https://github.com/user-attachments/assets/d11828a5-df6d-4fe5-88aa-f06e94ea21ff)
