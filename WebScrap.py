import buscador
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl import load_workbook


#função que tranforma o dicionário em uma tabela do excel
def Tabela_pesquisa(procura):
   
   #Onde está definido o site de busca e retorna os intens da pesquisa
   df_pesquisa = buscador.Pesquisa('https://www.mercadolivre.com.br/', procura) 

   #Cria o arquivo do excel vazio com uma aba. 
   arquivo = Workbook()
   aba = arquivo.active
   aba.title = "Base de dados"

   #Retorna os nomes do cabeçalho do dicionário
   nomes_colunas = list(df_pesquisa.keys())

   #Iteração com os intens do dicionário na tabela do excel
   for num_coluna, nome_coluna in enumerate(nomes_colunas, start=1):
       celula = aba.cell(row=1, column=num_coluna)
       celula.value = nome_coluna
       celula.alignment = Alignment(horizontal='center')
       lista_valores = df_pesquisa[nome_coluna]
       for num_linha, valor, in enumerate(lista_valores, start=2):
           celula = aba.cell(row=num_linha, column=num_coluna)
           celula.value = valor
           celula.alignment = Alignment(horizontal='center')
           if num_coluna == 5:
            celula.hyperlink = valor
            celula.alignment = Alignment(horizontal="left")

   #Salva a tabela do excel na pasta correta com o nome do item pesquisado e data da pesquisa
   caminho_pasta = 'Projeto-WebScrap\\Pesquisados\\'
   data = (str(datetime.now()))
   nome = (procura + '_' + data[:10] +'.xlsx')
   arquivo.save(caminho_pasta + nome)



print('O que deseja pesquisar: ')
pesquisado = input()

Tabela_pesquisa(pesquisado)