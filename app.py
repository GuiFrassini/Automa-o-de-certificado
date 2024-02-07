import openpyxl #Faz leitura do excel
from PIL import Image,ImageDraw,ImageFont

#workbook é igual a planilha em ingles então por isso atribui esse nome
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']


#neste trecho estamos passando linha a linha com iter rows
#min_row inicia a busca apartir da linha 2
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    #acessar cada celula que precisamos
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_de_participacao = linha[2].value
    data_inicio = linha[3].value
    data_fim = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    #Transferir os dados da planilha para a imagem

    font_nome = ImageFont.truetype('./Fontes/tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./Fontes/tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./Fontes/tahoma.ttf',55)

    image = Image.open('certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1020,827),nome_participante,fill='black',font=font_nome)
    desenhar.text((1060,950),nome_curso,fill='black',font=fonte_geral)
    desenhar.text((1435,1065),tipo_de_participacao,fill='black',font=fonte_geral)
    desenhar.text((1480,1182),str(carga_horaria),fill='black',font=fonte_geral)
    
    desenhar.text((750,1770),data_inicio,fill='black',font=fonte_data)
    desenhar.text((750,1930),data_fim,fill='black',font=fonte_data)

    desenhar.text((2220,1930),data_emissao,fill='black',font=fonte_data)

    image.save(f'./{indice}{nome_participante} certificado.png') #cria o nome que vai ser salvo a imagem e aonde vai ser salvo 

