import time
# import os
# import sys
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
# import json
import pyodbc


conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\marcos.gomes\NEOTASS PUBLICIDADE E PRODUCOES LT\Neotass - Documentos\Neotass\aDATA\BASE\BD_ADATA.accdb;')
# abrindo conexão com access 

cursor = conn.cursor()
cursor.execute('select PESQUISA, id_cnpj from cns_Enderecos_empresas where result_lat_long is null') # fazendo select para abrir uma tabela

tabela = cursor.fetchall()

url = "https://developers-dot-devsite-v2-prod.appspot.com/maps/documentation/utils/geocoder#q%3D"
print(url)
driver = webdriver.Chrome(executable_path='./chromedriver.exe') #chromedriver.exe
driver.get(url)
time.sleep(5)

for endereco in tabela:    
    element = driver.find_element_by_id("query-input")
    element.clear()
    element.send_keys(endereco.PESQUISA) 
    driver.find_element_by_id("geocode-button").click() #clicando no botão para pesquisqar
    time.sleep(2)

    
    id_cnpj1 = endereco.id_cnpj
    element = driver.find_element_by_id("result-0") 
    element2 = driver.find_element_by_id("result-0") 
    element = element.find_element_by_class_name("result-location")
    element2 = element2.find_element_by_class_name("result-formatted-address")
    lat_long =  element.text
    endereco_econtrado = element2.text
    endereco_econtrado = endereco_econtrado.replace("'","¹")

    lat_long = lat_long.replace(" (type: ROOFTOP)","") #substituindo palavras
    lat_long = lat_long.replace("Location: ","")  #substituindo palavras
    lat_long = lat_long.replace(" (type: GEOMETRIC_CENTER)","")  #substituindo palavras
    lat_long = lat_long.replace(" (type: RANGE_INTERPOLATED)","")  #substituindo palavras
    endereco = endereco.PESQUISA

    print (endereco)     
    print(lat_long, " - ", endereco_econtrado)
    
    sql = """\
        UPDATE tb_Dcnpj
        SET result_lat_long ='{lat_long}',
        endereco_encontrado ='{endereco_econtrado}'
        WHERE id_cnpj ={id_cnpj1}
        """.format(lat_long=lat_long,id_cnpj1=id_cnpj1,endereco_econtrado=endereco_econtrado) #montando sql para inserir lat e long na base

    cursor.execute(sql) #executando sql
    conn.commit()
    print('-----')
    print('-----')


driver.quit()
cursor.quit()
conn.quit()