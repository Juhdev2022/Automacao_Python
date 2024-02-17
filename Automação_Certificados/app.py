
                # Quero ver  apossibilidade de criar um programa usando o Python para automatizar enviando
                # os dados da planilha para preencher os campos mutávels no certificado padrão.
                # Tipo nome do curso, nome participante, tipo de participação , data o início data do término,
                # carga horária, data de emissão do certificado e as assinaturas do Gestor Geral, do Coordenador


# # Pegar os dados da planilha
import openpyxl
from PIL import Image, ImageDraw, ImageFont 
import os

# Obtém o caminho do diretório atual
current_directory = os.path.dirname(os.path.abspath('Automação_Certificados\planilha_alunos.xlsx'))

# Combina o caminho do diretório atual com o nome do arquivo
file_path = os.path.join(current_directory, 'planilha_alunos.xlsx')

workbook_alunos = openpyxl.load_workbook(file_path)

# Abrir a planilha
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # Cada célula que contém a info que precisamos
    nome_curso = linha[0].value  # nome do curso
    nome_participante = linha[1].value 
    tipo_participacao = linha[2].value
    carga_horaria = linha[5].value

    data_inicio = linha[3].value
    data_final = linha[4].value
    
    data_emissao = linha[6].value

    # Transferir para a imagem do certificado
    fonte_nome = ImageFont.truetype('./tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./tahoma.ttf',55)

    image_path = os.path.join(current_directory, 'certificado_padrao.jpg')
    image = Image.open(image_path)
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1020, 827), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435, 1065), tipo_participacao, fill='black', font=fonte_data)
    desenhar.text((1480, 1182), str(carga_horaria), fill='black', font=fonte_geral)

    desenhar.text((750, 1770),data_inicio, fill='black', font=fonte_data)
    desenhar.text((750, 1930),data_final, fill='black', font=fonte_data)

    desenhar.text((2220, 1930),data_emissao, fill='black', font=fonte_data)

    image.save(f'./{indice} {nome_participante} certificado.png')

