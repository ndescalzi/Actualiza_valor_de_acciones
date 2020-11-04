import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import datetime

#Te lee el archivo excel. 
archivo_excel ="Inversiones.xlsx"
excel = pd.read_excel(archivo_excel)
#Variable donde se van a guardar los valores de los titulos actualizados
Valor_actualizado =[]
#Actualiza los valores mediante Selenium
for Titulo in excel['Titulo']:
	#Navegacion en internet
	driver = webdriver.Chrome("C:\\Drivers_Selenium\\Chrome\\chromedriver.exe")
	driver.get("https://www.invertironline.com/Mercado/Cotizaciones")
	#Busca el titulo correspondiente de la iteración. 
	driver.find_element_by_id("header-busqueda").send_keys(Titulo)
	driver.find_element_by_id("header-busqueda").send_keys(Keys.ENTER)
	time.sleep(3)
	Valor = driver.find_element_by_xpath('//*[@id="IdTitulo"]/span[2]').text
	driver.close()
	Valor_actualizado.append(Valor)
#Actualizamos el Data Frame con la fecha del dia de ejecucion del programa y los nuevos valores
excel['fecha'] = datetime.date.today()
excel['Valor actual'] = Valor_actualizado
# Configuramos Pandas y cargamos el archivo correspondiente (en este caso se llama archivo.xlsx)                       
book = load_workbook(archivo_excel)
writer = pd.ExcelWriter(archivo_excel, engine='openpyxl') 
writer.book = book
# Guardamos el df en el excel en el lugar apropiado.
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
excel['Valor actual'].to_excel(writer, book.worksheets[1].title, startcol = 1,  index = False) 
excel['fecha'].to_excel(writer, book.worksheets[1].title, startcol = 2,  index = False) 
writer.save()