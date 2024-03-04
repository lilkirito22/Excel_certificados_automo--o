"""
1- pegar dados da planilha

2- transferir dados para imagem do certificado
"""


#pegar dados da planilha

import openpyxl
from PIL import Image, ImageDraw, ImageFont

#abrindo a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
  # cada celula que contem a info que precisamos
  nome_curso = linha[0].value # nome do curso
  nome_aluno = linha[1].value # nome do aluno
  tipo_participacao = linha[2].value # tipo de participação
  data_inicio = linha[3].value # data de inicio
  data_final = linha[4].value # data final
  carga_horaria = linha[5].value # carga horaria
  data_emissao = linha[6].value # data de emissão
  


  #transferir dados para imagem do certificado
  #definindo a fonte
  fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
  fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
  fonte_data = ImageFont.truetype('./tahoma.ttf', 55)


  #abrir a imagem
  image = Image.open('./certificado_padrao.jpg')
  desenhar = ImageDraw.Draw(image)
  desenhar.text((1015, 830), nome_aluno, fill='black', font=fonte_nome,)
  desenhar.text((1074, 956), nome_curso, fill='black', font=fonte_geral,)
  desenhar.text((1444, 1069), tipo_participacao, fill='black', font=fonte_geral,)
  desenhar.text((1503, 1190), str(carga_horaria),fill='black', font=fonte_geral,)
  desenhar.text((735, 1785), str(data_inicio), fill='black', font=fonte_data,)
  desenhar.text((725, 1934), str(data_final), fill='black', font=fonte_data,)
  desenhar.text((2205, 1930), str(data_emissao), fill='black', font=fonte_data,)
  image.save(f'./{indice} {nome_aluno}.png')



""" 

cordenadas :
Nome = 1015,887
Curso = 1074,1007
Participação = 1444, 1121
carga_horaria = 1503,1238
inicio: 739,1826
final:728,1971
emissao = 2230,1965


"""