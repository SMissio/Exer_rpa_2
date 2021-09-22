#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from selenium import webdriver as opsele
from selenium.webdriver.common.keys import Keys
import pyautogui as tempopc
import xlsxwriter
import os
meuNavegador = opsele.Chrome()
meuNavegador.get("https://www.google.com/")
tempopc.sleep(6)
meuNavegador.find_element_by_name("q").send_keys("Dolar Hoje")
tempopc.sleep(4)
meuNavegador.find_element_by_name("q").send_keys(Keys.RETURN)
tempopc.sleep(4)
valorDolar = meuNavegador.find_elements_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

#_________________________________________________________
tempopc.sleep(4)
meuNavegador.find_element_by_name("q").send_keys("")
tempopc.sleep(4)
tempopc.press('tab')
tempopc.sleep(4)
tempopc.press('enter')
tempopc.sleep(3)
meuNavegador.find_element_by_name("q").send_keys("Euro Hoje")
tempopc.sleep(4)
meuNavegador.find_element_by_name("q").send_keys(Keys.RETURN)
tempopc.sleep(3)
valorEuro = meuNavegador.find_elements_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text
tempopc.sleep(4)
#_____________________________

nomeArquivo = 'C:\\Users\\User1\\Desktop\\RPA1\\Dolar e Euro Google.xlsx'
plancriada = xlsxwriter.Workbook(nomeArquivo)
sheet = plancriada.add_worksheet()
tempopc.sleep(4)

sheet.write("A1","Dolar")
sheet.write("B1","Euro")
sheet.write("A2",valorDolar)
sheet.write("B2",valorEuro)

plancriada.close()
os.startfile(nomeArquivo)

print ("Dolar e Euro Extraido com sucesso!")

