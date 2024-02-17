import openpyxl
from PIL import Image, ImageDraw, ImageFont


# Abrindo a planinha no python
planilha_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
pagina_planilha = planilha_alunos['Planilha1']

for indice, linha in enumerate(pagina_planilha.iter_rows(min_row=2)):
    # Pegando os dados da planilha
    curso_planilha = linha[0].value # Coluna do nome do curso
    participante_planilha = linha[1].value # Coluna do nome do aluno
    data_ini = linha[2].value # Coluna data inicio
    data_ini_str = data_ini.strftime("%Y-%m-%d") # Conversão date para string
    data_fim = linha[3].value # Coluna data final
    data_fim_str = data_fim.strftime("%Y-%m-%d")
    data_emi = linha[4].value # Coluna data emissão
    data_emi_str = data_emi.strftime("%Y-%m-%d")

    #
    font_geral = ImageFont.truetype('./Rasa-Regular.ttf', 48)
    font_nome = ImageFont.truetype('./Calligrapher.ttf', 150)

    #
    image = Image.open('./certificado.png')
    desenhar = ImageDraw.Draw(image)

    # desenhar.text = criar texto na imagem 
    # (400,900) = coordenada do texto, primeiro valor Horizontal e segundo Vertical
    # variavel da planilha para obter o texto
    # fill='' para escolher a cor do texto
    # font= para escolher a fonte do texto
    desenhar.text((700,500),participante_planilha,fill='blue',font=font_nome)
    desenhar.text((970,700),curso_planilha,fill='blue',font=font_geral)

    # como a biblioteca só lê texto precisamos converter os dados para string (linha do código 14)
    desenhar.text((450,759),data_ini_str,fill='black',font=font_geral)
    desenhar.text((705,759),data_fim_str,fill='black',font=font_geral)
    desenhar.text((970,886),data_emi_str,fill='black',font=font_geral)



    image.save(f'./certificados/{indice} {participante_planilha}certificado.png')

