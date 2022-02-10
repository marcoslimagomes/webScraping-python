import time
import pandas as pd
from bs4 import BeautifulSoup  as soup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import json
import pyodbc



conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\marcos.gomes\NEOTASS PUBLICIDADE E PRODUCOES LT\Neotass - Documentos\Neotass\aDATA\BASE\BD_ADATA.accdb;')
cursor = conn.cursor()
cursor.execute("select id_cnpj, result_lat_long from cns_Enderecos_empresas ORDER BY id_cnpj ASC") 

tabela1 = cursor.fetchall()
tabela2 = tabela1

#https://www.google.com.br/maps/dir/-23.4851224,-46.5917735/-23.5662942,-46.6468642/@-23.5143625,-46.6476861,13.02z/data=!4m2!4m1!3e3

driver = webdriver.Chrome(executable_path='./chromedriver.exe')
for RESULTADO1 in tabela1:       
    lat_long1 = RESULTADO1.result_lat_long
    id_end1 = RESULTADO1.id_cnpj    
    if id_end1 >= 274:
        for RESULTADO2 in tabela2:
            if RESULTADO1.id_cnpj != RESULTADO2.id_cnpj:
                lat_long2 = RESULTADO2.result_lat_long
                id_end2 = RESULTADO2.id_cnpj 
                link = """\
                https://www.google.com.br/maps/dir/{lat_long1}/{lat_long2}/@-23.6065417,-46.7777652,10.96z/data=!4m2!4m1!3e3
                    """.format(lat_long1=lat_long1,lat_long2=lat_long2)

                link = link.replace(" ","")
                driver.get(link)
                time.sleep(1)
                try:
                    page = soup(driver.page_source, 'html.parser')   
                    
                    element = driver.find_element_by_class_name("xB1mrd-T3iPGc-iSfDt-n5AaSd") ####### xB1mrd-T3iPGc-iSfDt-duration gm2-subtitle-alt-1
                    tempo = element.text #xB1mrd-T3iPGc-iSfDt-n5AaSd #xB1mrd-T3iPGc-iSfDt-duration

                    #element = page.findAll('div',{'jstcache':'314'})
                    #tempos = page.findAll('div',{'jstcache':'314'})
                    #tempo = element[1]
                    print(tempo)
                
                    #print(tempos[:50])
                   
                except ValueError:                    
                    tempo = " "               
                

                print(tempo , "          id:" , id_end1 , " id:" , id_end2)
                sql = """\
                    INSERT INTO tb_Fdistancia_cnpj(id_cnpj1, id_cnpj2, tempo, link_google)
                    VALUES ('{id_end1}','{id_end2}','{tempo}','{link_google}')
                    """.format(id_end1=id_end1,id_end2=id_end2,tempo=tempo,link_google=link)

                cursor.execute(sql)
                conn.commit()
    print('------- novo endereço comp ---')
        



cursor.quit()
conn.quit()


def coletar_info():
    print('atenção')