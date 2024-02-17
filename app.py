import openpyxl
from PIL import Image, ImageDraw, ImageFont


# Lendo a planinha no python
planilha_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
# Escolhendo a página que os dados estão na planilha
pagina_planilha = planilha_alunos['Planilha1']

# Criado um indice para cada id de aluno da planilha por linha
# Como a primeira linha da planilha é um cabeçalho, min_row=2 é para começar a ler a partir da segunda linha, porém linha 2 = indice 0
for indice, linha in enumerate(pagina_planilha.iter_rows(min_row=2)):

    # Pegando os dados da planilha
    curso_planilha = linha[0].value # Coluna do nome do curso
    participante_planilha = linha[1].value # Coluna do nome do aluno
    data_ini = linha[2].value # Coluna data inicio
    data_ini_str = data_ini.strftime("%Y-%m-%d") # Conversão da data inicial para string
    data_fim = linha[3].value # Coluna data final
    data_fim_str = data_fim.strftime("%Y-%m-%d") # Conversão da data final para string
    data_emi = linha[4].value # Coluna data emissão
    data_emi_str = data_emi.strftime("%Y-%m-%d") # Conversão da data emissao para string

    # Duas fontes adicionadas para esse projeto
    # font Rasa-Regular para todos textos com um tamanho de 48
    font_geral = ImageFont.truetype('./Rasa-Regular.ttf', 48)
    # font Calligrapher só para o nome do aluno com um tamanho de 150
    font_nome = ImageFont.truetype('./Calligrapher.ttf', 150)

    # Pegando a imagem do certificado sem informção na pasta
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

    # local do save dos certificados com indice "id" e nome do participante
    image.save(f'./certificados/{indice} {participante_planilha}certificado.png')

