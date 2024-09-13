import os
import re
import fitz  # PyMuPDF
import pandas as pd
import shutil
from google.cloud import vision
from google.cloud.vision_v1 import types
from PIL import Image
import io
import difflib


os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "intropythoncloudvision.json"

def pdf_to_text(pdf_path):
    """
    Extrai o texto de todas as páginas de um arquivo PDF.

    Args:
        pdf_path (str): O caminho para o arquivo PDF.

    Returns:
        str: Todo o texto extraído do PDF.
    """
    text = ""
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    
    pdf_document.close()
    
    return text


def extract_image_from_pdf(pdf_path):
    """
    Extrai todas as imagens das páginas de um arquivo PDF, corta parte das imagens e as retorna.

    Args:
        pdf_path (str): O caminho para o arquivo PDF.

    Returns:
        list: Uma lista de tuplas contendo as imagens cortadas em bytes e o número da página correspondente.
        fitz.Document: O objeto PDF aberto.
    """
    pdf = fitz.open(pdf_path)
    images_list = []
    
    for i in range(len(pdf)):
        page = pdf.load_page(i)
        images = page.get_images(full=True)
        
        if not images:
            continue
        
        for img_index in range(len(images)):
            image = images[img_index]
            xref = image[0]
            base_image = pdf.extract_image(xref)
            image_bytes = base_image.get("image")

            if image_bytes is None:
                continue
            
            try:
                image = Image.open(io.BytesIO(image_bytes))
                width, height = image.size
                new_height = int(height * 0.75)
                
                # cortar a imagem para remover carimbos e anotações à mão
                cropped_image = image.crop((0, 0, width, new_height))
                
                # converter a imagem cortada de volta para bytes
                img_byte_arr = io.BytesIO()
                cropped_image.save(img_byte_arr, format=image.format)
                cropped_image_bytes = img_byte_arr.getvalue()
                
                images_list.append((cropped_image_bytes, i))
            
            except Exception as e:
                print(f"Erro ao processar a imagem da página {i}, índice {img_index}: {e}")
                continue
    
    return images_list, pdf


def image_to_text(image_bytes):
    """
    Realiza OCR em uma imagem usando a API Google Cloud Vision e retorna o texto.

    Args:
        image_bytes (bytes): A imagem em bytes para a qual o OCR será aplicado.

    Returns:
        str: O texto extraído da imagem. Retorna None se nenhum texto for encontrado.
    """
    client = vision.ImageAnnotatorClient()
    image = types.Image(content=image_bytes)
    image_context = vision.ImageContext()
    response = client.document_text_detection(image=image, image_context=image_context)
    full_text = response.text_annotations
    
    return full_text[0].description if full_text else None


def convert_to_float(value):
    """
    Converte um valor monetário no formato string para um valor float.

    Args:
        value (str): O valor monetário como string (com pontos e vírgulas).

    Returns:
        float: O valor convertido para float.
    """
    exp = 1 if (value[-2] == '.') or (value[-2] == ',') else 2

    value = re.sub(r'[^\d.,]', '', value)  # remover R$
    value = value.replace('.', '').replace(',', '')  # remover , e .

    return float(value) / 10**exp


def scrapping_data(text):
    """
    Extrai os municípios, estabelecimentos e valores monetários de um texto.

    Args:
        text (str): O texto de onde os dados serão extraídos.

    Returns:
        tuple: Três listas contendo municípios, estabelecimentos e o valor monetário (float) mais alto encontrado.
    """
    municipios = []
    estabelecimentos = []
    valor = None

    if text is None:
        return 0, 0, 0

    regex_municipio = re.compile(r"munic[ií]pio:\s*(.*)", re.IGNORECASE)
    finds_municipios = regex_municipio.findall(text)

    for m in finds_municipios:
        municipios.append(m.strip())

    regex_estabelecimento = re.compile(r"estabelecimento:\s*(.*)", re.IGNORECASE)
    finds_estabelecimentos = regex_estabelecimento.findall(text)

    for e in finds_estabelecimentos:
        estabelecimentos.append(e.strip())

    valores = re.findall(r"(?:^|\s)(?:R\$?\s?)?((?:\d{1,3}[.,]\d{3}[.,]\d{1,2})|(?:\d{1,3}[.,]\d{1,2})|(?:\d{1,3}[.,]\d{1,2}))\b", text)
    
    if valores:
        valores_float = [convert_to_float(valor) for valor in valores]
        valor = max(valores_float)

    return municipios, estabelecimentos, valor 


def most_similar_2(targets, candidate):
    """
    Encontra o item mais semelhante ao candidato em uma lista de alvos.

    Args:
        targets (list): Uma lista de strings para comparar.
        candidate (str): A string candidata a ser comparada.

    Returns:
        str: O item mais semelhante ao candidato.
    """
    high_score = 0
    most_similar = None

    for target in targets:
        if isinstance(target, str) and isinstance(candidate, str):
            score = difflib.SequenceMatcher(None, target, candidate).ratio()
            if score > high_score:
                high_score = score
                most_similar = target
    return most_similar


def most_similar(targets, candidates):
    """
    Encontra o item mais semelhante entre várias combinações de alvos e candidatos.

    Args:
        targets (list): Uma lista de strings de comparação.
        candidates (list): Uma lista de strings candidatas.

    Returns:
        tuple: O item mais semelhante e a pontuação de similaridade.
    """
    high_score = 0
    most_similar = None

    for candidate in candidates:
        for target in targets:
            if isinstance(target, str) and isinstance(candidate, str):
                score = difflib.SequenceMatcher(None, target, candidate).ratio()
                if score > high_score:
                    high_score = score
                    most_similar = target
                    if high_score > 0.9:
                        return most_similar, high_score

    return most_similar, high_score


def save_romaneios(pdf, page_num):
    """
    Salva uma página específica do PDF em um diretório separado para revisão.

    Args:
        pdf (fitz.Document): O objeto PDF aberto.
        page_num (int): O número da página a ser salva.
    """
    not_read_dir = "notas_nao_lidas"
    if not os.path.exists(not_read_dir):
        os.makedirs(not_read_dir)

    pdf_writer = fitz.open()
    pdf_writer.insert_pdf(pdf, from_page=page_num, to_page=page_num)
    pdf_writer.save(os.path.join(not_read_dir, f"page_{page_num + 1}.pdf"))
    pdf_writer.close()
    print(f"Página {page_num + 1} movida para {not_read_dir}")


def formato_brasileiro(valor):
    """
    Formata um valor float no formato monetário brasileiro.

    Args:
        valor (float): O valor a ser formatado.

    Returns:
        str: O valor formatado como string.
    """
    if isinstance(valor, float):
        return f'{valor:,.2f}'.replace(',', '').replace('.', ',')
    return valor


def extract_romaneios(pdf_path, new_path):
    """
    Extrai dados de romaneios de um PDF, realiza comparações de similaridade com um gabarito e atualiza uma planilha Excel.

    Args:
        pdf_path (str): O caminho para o PDF contendo os romaneios.
        new_path (str): O caminho para a planilha Excel a ser atualizada.
    """
    images_romaneios, pdf = extract_image_from_pdf(pdf_path)
    df_romaneios = pd.read_excel(new_path)

    df_romaneios['Municipio'] = df_romaneios['Municipio'].astype(str).str.upper()
    df_romaneios['Estabelecimento'] = df_romaneios['Estabelecimento'].astype(str).str.upper()
    municipios_gabarito = df_romaneios['Municipio'].unique().tolist()

    for image_bytes, page_num in images_romaneios:
        text = image_to_text(image_bytes)
        municipios, estabelecimentos, valor = scrapping_data(text)

        if municipios and estabelecimentos and valor:
            municipios_upper = [m.upper() for m in municipios]
            estabelecimentos_upper = [e.upper() for e in estabelecimentos]

            i = 0
            while i < len(municipios_upper):
                most_similar_mun, _ = most_similar(municipios_gabarito, municipios_upper)
                df_municipio_target = df_romaneios[df_romaneios['Municipio'] == most_similar_mun]
                most_similar_estab, high_score_estab = most_similar(df_municipio_target['Estabelecimento'].tolist(), estabelecimentos_upper)
                if high_score_estab >= 0.65:
                    break
                municipios_upper.pop(0)
                i += 1

            matching_rows = df_romaneios[
                (df_romaneios['Municipio'].str.strip() == most_similar_mun.strip()) &
                (df_romaneios['Estabelecimento'].str.strip() == most_similar_estab.strip())
            ]

            if not matching_rows.empty:
                row_index = matching_rows.index[0]
                df_romaneios.at[row_index, 'Valor Total'] = (
                    df_romaneios.at[row_index, 'Valor Total'] + valor
                )

        else:
            save_romaneios(pdf, page_num)

    df_romaneios.to_excel(new_path, index=False)


def extract_gerenciais(gerenciais_path, excel_path):
    """
    Extrai dados de um PDF de Consultas Gerenciais e os salva em uma planilha Excel.

    Args:
        gerenciais_path (str): O caminho para o PDF de Consultas Gerenciais.
        excel_path (str): O caminho para a planilha Excel a ser atualizada.
    """
    text = pdf_to_text(gerenciais_path)
    municipios, estabelecimentos, valor = scrapping_data(text)

    df = pd.read_excel(excel_path)

    for municipio, estabelecimento in zip(municipios, estabelecimentos):
        df = df.append({
            'Municipio': municipio,
            'Estabelecimento': estabelecimento,
            'Valor Total': valor
        }, ignore_index=True)

    df.to_excel(excel_path, index=False)


def drop_empty_lines(excel_path):
    """
    Remove linhas vazias de uma planilha Excel.

    Args:
        excel_path (str): O caminho para a planilha Excel a ser limpa.
    """
    df = pd.read_excel(excel_path)
    df.dropna(how='all', inplace=True)
    df.to_excel(excel_path, index=False)


def create_excel(romaneios_path, gerenciais_path, excel_path, new_path):
    """
    Executa o fluxo completo de extração de dados de PDFs e atualização da planilha Excel.

    Args:
        romaneios_path (str): O caminho para o PDF de romaneios.
        gerenciais_path (str): O caminho para o PDF de Consultas Gerenciais.
        excel_path (str): O caminho para a planilha Excel original.
        new_path (str): O caminho para a nova planilha Excel a ser gerada.
    """
    shutil.copy(excel_path, new_path)
    extract_romaneios(romaneios_path, new_path)
    extract_gerenciais(gerenciais_path, new_path)
    drop_empty_lines(new_path)


if __name__ == '__main__':
    romaneios_path = "romaneios.pdf" # pdf com os romaneios de entrada
    gabarito_excel_path = "gabarito.xlsx" # tabela com a primeira coluna de municipios e a segunda coluna com o nome dos estabelecimentos 
    save_path = "resultados_atualizados.xlsx" # saida atualizada com os dados já extraidos
    gerenciais_path = "ConsultasGerenciais.pdf" # consultas do sistema web 
    create_excel(romaneios_path, gerenciais_path, gabarito_excel_path, save_path)


   
