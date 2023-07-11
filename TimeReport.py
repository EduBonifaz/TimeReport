from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from datetime import datetime
import pandas as pd
import json
import time
import locale
import os
import pyautogui

with open('config.json') as f:
    config = json.load(f)

URLTimeReport = config["URLTimeReport"]
URLSheet = config["URLSheet"]
FolderDrive = config["FolderDrive"]
CarpetaDescarga = config["CarpetaDescarga"]
RutaUserData = config["RutaUserData"]
ProfileBBVA = config["ProfileBBVA"]
ProfileBluetab = config["ProfileBluetab"]

locale.setlocale(locale.LC_TIME, 'es_ES')
fecha_actual = datetime.now()
day = fecha_actual.strftime('%d')
month = fecha_actual.strftime('%m')
year = fecha_actual.strftime('%Y')
month_str = datetime.now().strftime('%B').capitalize()

path = f'{CarpetaDescarga}/{day}-{month}.png'

options = webdriver.ChromeOptions()
options.add_argument(f'--user-data-dir={RutaUserData}')
options.add_argument(f'--profile-directory={ProfileBBVA}')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.maximize_window()
driver.get('https://www.google.com')
driver.get(f'{URLTimeReport}')

WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH,'//*/img[@class="task-info infoTablet iconInfo"]/following-sibling::div')))
Proyecto = driver.find_elements(By.XPATH,r'//div[@class="task-title full-width"]')
Tarea = driver.find_elements(By.XPATH,r'//img[@class="task-info infoDesktop iconInfo"]/following-sibling::div[@class="full-width task-description"]')

Columnas = ["Proyecto","Tarea","Horas","Minutos"]
ReportDf = pd.DataFrame(columns=Columnas)
for i in range(len(Proyecto)):
	ReportDf.loc[i] = [Proyecto[i].text,Tarea[i].text,0,0]

if os.path.isfile("./Report.xlsx"):
	ReportTareasDf = pd.read_excel('./Report.xlsx', sheet_name='Tareas', dtype = 'object', usecols="A:D").dropna()
	HorasDf = pd.read_excel('./Report.xlsx', sheet_name='Tareas', dtype = 'object', usecols="F").dropna()
	if ReportDf[['Proyecto', 'Tarea']].equals(ReportTareasDf[['Proyecto', 'Tarea']]):
		if HorasDf.loc[0][0] == "8:00 Horas":
			Botones = driver.find_elements(By.XPATH,r'//*/li[@class="task task-grid__task"]/div')
			Minutos = driver.find_elements(By.XPATH,r'//*/input[contains(@id,"input-minutes")]')
			Horas = driver.find_elements(By.XPATH,r'//*/input[contains(@id,"input-hours")]')
			Aceptar = driver.find_elements(By.XPATH,r'//*/button[contains(@class,"accept")]')
			Collapse = driver.find_elements(By.XPATH,r'//*/li[@class="task task-grid__task"]/div/div[3]')

			for i in range(len(Botones)):
				Botones[i].click()
				WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH,'//*/div[@class="task-collapse"]')))
				Horas[i].send_keys(Keys.BACKSPACE,Keys.BACKSPACE,ReportTareasDf.iloc[i]["Horas"])
				Minutos[i].send_keys(Keys.BACKSPACE,Keys.BACKSPACE,ReportTareasDf.iloc[i]["Minutos"])
				Aceptar[i].click()
				Salir = 1
				while Salir:
					try:
						WebDriverWait(driver, 0.5).until(expected_conditions.element_to_be_clickable((By.XPATH,'//*/div[@class="task-collapse"]')))
					except:
						Salir = 0

			screenshot = driver.find_element(By.XPATH,r'//*/section[@class="tasks-bg pr-0 main-view__task-grid"]').screenshot_as_png

			with open(path, 'wb') as file:
			    file.write(screenshot)
			driver.quit()
			options = webdriver.ChromeOptions()
			options.add_argument(f'--user-data-dir={RutaUserData}')
			options.add_argument(f'--profile-directory={ProfileBluetab}')
			driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
			actions = ActionChains(driver)
			driver.get('https://www.google.com')
			driver.get(f'{FolderDrive}')
			WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH,f'//*/c-wiz[div/div/div/div/div/div[@data-tooltip="Google Drive Folder: {year}"]]')))
			Year = driver.find_element(By.XPATH,f'//*/c-wiz[div/div/div/div/div/div[@data-tooltip="Google Drive Folder: {year}"]]')
			Year.click()
			time.sleep(1)
			actions.double_click(Year).perform()

			WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH,f'//*/c-wiz[div/div/div/div/div/div[@data-tooltip="Google Drive Folder: {month}-{month_str}"]]')))
			Year = driver.find_element(By.XPATH,f'//*/c-wiz[div/div/div/div/div/div[@data-tooltip="Google Drive Folder: {month}-{month_str}"]]')
			Year.click()
			time.sleep(1)
			actions.double_click(Year).perform()
			time.sleep(1)
			WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH,r'//button[@guidedhelpid="new_menu_button"]')))
			driver.find_element(By.XPATH,r'//button[@guidedhelpid="new_menu_button"]').click()
			WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH,r'//div[div/span/div[@data-tooltip="File upload"]]')))
			driver.find_element(By.XPATH,r'//div[div/span/div[@data-tooltip="File upload"]]').click()
			WebDriverWait(driver, 20).until(expected_conditions.presence_of_element_located((By.XPATH,r'//*/input[@type="file"]')))
			time.sleep(1)
			pyautogui.hotkey('alt', 'f4')
			driver.find_element(By.XPATH,r'//*/input[@type="file"]').send_keys(path)
			WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH,f'//*/c-wiz[div/div/div/div/div/div[@data-tooltip="Image: {day}-{month}.png"]]')))
			imagen = driver.find_element(By.XPATH,f'//*/div[div/div/div/div/div[@data-tooltip="Image: {day}-{month}.png"]]').get_attribute("data-id")
			driver.get(f'https://drive.google.com/file/d/{imagen}')
			driver.execute_script(f"window.open('{URLSheet}', '_blank');")
			time.sleep(1000)
		else:
			print(f'Horas declaradas: {HorasDf.loc[0][0]}, Favor de llenar las horas y que sumen 8:00 Horas')
			print(f'Editar el Archivo: {os.getcwd()}\Report.xlsx')
	else:
		print('Se actualizó el Archivo "Report.xlsx", favor de llenar las horas y que sumen 8:00 Horas')
		print(f'Editar: {os.getcwd()}\Report.xlsx')
		driver.quit()
		writer = pd.ExcelWriter("./Report.xlsx", engine="xlsxwriter")
		ReportDf.to_excel(writer,sheet_name='Tareas', index=False)
		workbook = writer.book
		worksheet = writer.sheets["Tareas"]
		worksheet.set_column(0, 0, 40)
		worksheet.set_column(1, 1, 45)
		worksheet.set_column(2, 3, 9)
		formato = workbook.add_format({'bold': True, 'align': 'center'})
		worksheet.set_column(5, 5, 13,formato)
		worksheet.write('F1', 'Horas Totales')
		worksheet.write_formula('F2', '=QUOTIENT(SUM(D:D),60)+SUM(C:C)&":"&TEXT(MOD(SUM(D:D),60),"00")&" Horas"')
		writer.close()
else:
	print('No se encontró el Archivo "Report.xlsx", favor de llenar las horas y que sumen 8:00 Horas')
	print(f'Editar: {os.getcwd()}\Report.xlsx')
	driver.quit()
	writer = pd.ExcelWriter("./Report.xlsx", engine="xlsxwriter")
	ReportDf.to_excel(writer,sheet_name='Tareas', index=False)
	workbook = writer.book
	worksheet = writer.sheets["Tareas"]
	worksheet.set_column(0, 0, 40)
	worksheet.set_column(1, 1, 45)
	worksheet.set_column(2, 3, 9)
	formato = workbook.add_format({'bold': True, 'align': 'center'})
	worksheet.set_column(5, 5, 13,formato)
	worksheet.write('F1', 'Horas Totales')
	worksheet.write_formula('F2', '=QUOTIENT(SUM(D:D),60)+SUM(C:C)&":"&TEXT(MOD(SUM(D:D),60),"00")&" Horas"')
	writer.close()

print()
print('Finalizó el Programa')