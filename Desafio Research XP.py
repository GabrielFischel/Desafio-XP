#Gabriel Fischel 28/07/2024
#=============================================================================#
# IMPORTANTE 

# FAÇA O DOWNLOAD DA PASTA ZIP ENVIADA POR EMAIL E EXTRAIA TODOS OS ARQUIVOS PARA A PASTA DE DOWNLOAD

# ANTES DE RODAR O CÓDIGO, EXECUTAR OS SEGUINTES COMANDOS NO CONSOLE (INFERIOR A DIREITA NO SPYDER):
    
   # pip install python-docx reportlab


# APÓS RODAR O PROGRAMA, FAVOR CHECAR A PASTA DE DOWNLOADS DO COMPUTADOR E PROCURAR PELO ARQUIVO COM O NOME:
    # Guia de Investimentos Trend

# OBRIGADO
#=============================================================================#


#Importando as bibliotecas-----------------------------------------------------
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import utils
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import os
from pathlib import Path

#Definindo a função para ler o Word--------------------------------------------
def read_word_document(file_path):
    document = Document(file_path)
    content = []
    for para in document.paragraphs:
        style = para.style.name
        text = para.text
        content.append((style, text))
    return content

#Definindo a função que irá criar o cabeçalho----------------------------------
def draw_header(c, width, height, margin):
    c.setFont("Roboto-Light", 12)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(margin, height - margin + 20, "Alocação e Fundos | XP Research")
    c.drawRightString(width - margin, height - margin + 20, "28/07/2024")

#Definindo a função que irá criaro rodapé--------------------------------------
def draw_footer(c, width, margin, logo_path):
    # Carregando a imagem do logo
    logo = utils.ImageReader(logo_path)
    logo_width, logo_height = logo.getSize()
    aspect_ratio = logo_height / logo_width
    
    # Define o tamanho desejado do logo
    desired_logo_width = 25  # Ajuste esse valor para ajustar o tamanho
    desired_logo_height = desired_logo_width * aspect_ratio
    
    # Desenhando o logo
    c.drawImage(logo_path, margin, margin - desired_logo_height, width=desired_logo_width, height=desired_logo_height, mask='auto')
    
    # Colocando o texto do lado do logo
    c.setFont("Roboto-Light", 12)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(margin + desired_logo_width + 5, margin - desired_logo_height / 2, "| Gabriel Fischel")

#Definindo a função que vai criar o PDF----------------------------------------
def create_pdf(content, output_path, logo_path):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    margin = 50  # Definindo a margem do documento
    line_height = 14  # Altura da linha
    
    # Registrando as fontes que serão utilizadas
    pdfmetrics.registerFont(TTFont('Roboto', 'Roboto-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('Roboto-Light', 'Roboto-Light.ttf'))

    # Escolhendo PDF meta-data
    c.setTitle("Guia de Investimentos Trend")
    
    y_position = height - margin

    page_number = 1
    for style, line in content:
        if page_number == 1 or style == "Heading 2":
            if page_number > 1:
                c.showPage()  # Termine a pagina atual e comece uma nova
                y_position = height - margin  # Resetando y_position para a página nova
            draw_header(c, width, height, margin)  # Desenhando os cabeçalhos
            draw_footer(c, width, margin, logo_path)  # Desenhando os rodapés
            page_number += 1
            
        # Definindo a fonte e tamanho que será utilizado dependendo do tipo de texto
        if style.startswith("Heading"):
            if style == "Heading 1":
                font_name = "Roboto-Light"
                font_size = 16
                font_color = (255/255, 188/255, 0/255)
            elif style == "Heading 2":
                font_name = "Roboto"
                font_size = 14
                font_color = (0/255, 0/255, 0/255)
            elif style == "Heading 3":
                font_name = "Roboto-Light"
                font_size = 12
                font_color = (50/255, 52/255, 54/255)
            else:
                font_name = "Roboto-Light"
                font_size = 10
                font_color = (0, 0, 0)
            
            c.setFont(font_name, font_size)
            c.setFillColorRGB(*font_color)
            y_position -= 20
        else:
            c.setFont("Roboto-Light", 10)
            c.setFillColorRGB(0, 0, 0)

        # Dividindo linhas longas em várias linhas
        text_lines = utils.simpleSplit(line, c._fontname, c._fontsize, width - 2*margin)
        for text_line in text_lines:
            if y_position < margin:
                c.showPage()
                y_position = height - margin
                draw_header(c, width, height, margin)
                draw_footer(c, width, margin, logo_path)
                c.setFont("Roboto-Light", 10 if not style.startswith("Heading") else font_name)
                c.setFontSize(font_size if style.startswith("Heading") else 10)
                c.setFillColorRGB(*font_color if style.startswith("Heading") else (0, 0, 0))
            
            c.drawString(margin, y_position, text_line)
            y_position -= line_height

        if style.startswith("Heading"):
            y_position -= 10

    c.save()

#Definindo a principal função que vai executar o código------------------------
def main():
    word_file = "Guia de Investimentos (lista de Trends1).docx"
    logo_path = "XP Logo.png"  # Path to the logo image file
    
    # Determinando o diretório de download do usuário
    home_dir = Path.home()
    downloads_dir = home_dir / "Downloads"
    pdf_output = downloads_dir / "Guia_de_Investimentos_Trend.pdf"
    
    content = read_word_document(word_file)
    create_pdf(content, str(pdf_output), logo_path)
    print(f"PDF criado com sucesso: {pdf_output}")
    print("CHECAR A PASTA DE DOWNLOAD DO COMPUTADOR")
if __name__ == "__main__":
    main()
