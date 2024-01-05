import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
import zipfile
import requests
import pandas as pd

# Ruta al controlador de Chrome (reemplaza 'path/to/chromedriver' con la ubicación real)
chrome_driver_path = 'chromedriver.exe'
chrome_service = ChromeService(chrome_driver_path)

# Inicializar el navegador Chrome
driver = webdriver.Chrome(service=chrome_service)

try:
    # Navegar a la página
    driver.get("https://cde.ucr.cjis.gov/LATEST/webapp/#/pages/home")

    # Esperar a que la página cargue completamente
    time.sleep(5)

    # Localizar el enlace y hacer clic
    driver.find_element(By.LINK_TEXT, "Documents & Downloads").click()

    # Esperar a que la página cargue
    time.sleep(5)

    # Localizar la sección "National Incident-Based Reporting System (NIBRS) Tables"
    seccion_especifica = driver.find_element(By.ID, "dwnnibrs-card")

    # Esperar a que la página cargue
    time.sleep(5)

    # Elegir la tabla: Victims, año: 2022, location: Florida
    seccion_especifica.find_element(By.ID, "table-selector").click()
    seccion_especifica.find_element(
        By.XPATH, "//option[contains(text(),'Victims')]").click()
    seccion_especifica.find_element(By.ID, "year-selector").click()
    seccion_especifica.find_element(
        By.XPATH, "//option[contains(text(),'2022')]").click()
    seccion_especifica.find_element(By.ID, "location-selector").click()
    seccion_especifica.find_element(
        By.XPATH, "//option[contains(text(),'Florida')]").click()

    # Esperar a que la página cargue
    time.sleep(2)

    # Descargar el archivo
    driver.find_element(By.ID, "download-button").click()

    # Esperar a que se complete la descarga (ajusta el tiempo según sea necesario)
    time.sleep(10)

    # Extraer el archivo zip
    with zipfile.ZipFile("Victims_Age_by_Offense_Category_2022.zip", "r") as zip_ref:
        zip_ref.extractall()

    # Leer el archivo Excel con Pandas
    df = pd.read_excel("Victims_Age_by_Offense_Category_2022.xlsx")

    # Filtrar por la categoría "Crimes Against Property"
    df_filtered = df[df['Category'] == 'Crimes Against Property']

    # Generar CSV sin totales, footer ni index
    df_filtered.to_csv("output.csv", index=False, header=True)

    print("Prueba completada. El archivo CSV se ha generado correctamente.")

finally:
    # Cerrar el navegador al finalizar
    driver.quit()
