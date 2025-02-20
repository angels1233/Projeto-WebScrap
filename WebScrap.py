import pandas as pd
from datetime import datetime
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl import load_workbook


def Pesquisa(url, pesquisa):

    preco_old_list = []
    preco_new_list = []
    desconto_list = []
    link_list = []
    nome_list = []

    service = Service()
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service = service, options = options)

    driver.get(url)
    driver.find_element(By.ID,'cb1-edit').send_keys(pesquisa)
    time.sleep(1)
    driver.find_element(By.CLASS_NAME,'nav-search-btn').click()
    
    itens = driver.find_elements(By.CLASS_NAME,'ui-search-result__wrapper')

    for iten in itens: 
        link = iten.find_element(By.TAG_NAME,'a').get_attribute('href')
        nome_aux = iten.find_element(By.CLASS_NAME,'poly-component__title-wrapper')
        nome = nome_aux.find_element(By.TAG_NAME,'a').text
        preco_new_aux = iten.find_element(By.CLASS_NAME,'poly-price__current')    
        preco_new = preco_new_aux.find_element(By.CLASS_NAME,'andes-money-amount__fraction').text

        try:
            cents_new = preco_new_aux.find_element(By.CLASS_NAME,'andes-money-amount__cents').text
            preco_new = (preco_new +',' + cents_new)
        
        except:
             preco_new = (preco_new + ',' + '00')

    
        try:
            iten.find_element(By.CLASS_NAME,'andes-money-amount--previous')
            desconto = preco_new_aux.find_element(By.CLASS_NAME,'andes-money-amount__discount').text[:3]
            preco_old_aux = iten.find_element(By.CLASS_NAME,'andes-money-amount--previous')
            preco_old = preco_old_aux.find_element(By.CLASS_NAME,'andes-money-amount__fraction').text
            
            try:
                cents_old = preco_old_aux.find_element(By.CLASS_NAME,'andes-money-amount__cents').text
                preco_old = (preco_old + ',' + cents_old)
            except:
                preco_old = (preco_old + ',' + '00')

        except:
            preco_old = ''
            desconto = ''

        preco_old_list.append(preco_old)
        preco_new_list.append(preco_new)
        desconto_list.append(desconto)
        link_list.append(link)
        nome_list.append(nome)

        # print('preço full: '+ preco_old + '  Desconto: ' + desconto + '  preço com desconto: ' + preco_new)
 
    dictIten = {'nome': nome_list,
                'preco': preco_new_list,
                'desconto': desconto_list,
                'preco_anterior': preco_old_list,
                'link': link_list}
    
    # return pd.DataFrame.from_dict(dictIten)
    return dictIten
    time.sleep(10)

def Tabela_pesquisa(procura):
   
   df_pesquisa = Pesquisa('https://www.mercadolivre.com.br/', procura) 

   arquivo = Workbook()
   aba = arquivo.active
   aba.title = "Base de dados"

   nomes_colunas = list(df_pesquisa.keys())

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


   
   data = (str(datetime.now()))
   nome = (procura + '_' + data[:10] +'.xlsx')
   arquivo.save(nome)
#    print(df_pesquisa)
   
#    df_pesquisa.to_excel(nome)



print('O que deseja pesquisar: ')
pesquisado = input()

Tabela_pesquisa(pesquisado)