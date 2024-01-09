import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
import zipfile
import pandas as pd

# Ruta al controlador de Chrome (reemplaza 'path/to/chromedriver' con la ubicación real)
chrome_driver_path = 'chromedriver.exe'
chrome_service = ChromeService(chrome_driver_path)

path = os.path.dirname(os.path.abspath(__file__))
prefs = {"download.default_directory": path}
options = Options()
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=chrome_service, options=options)

try:
    # Navegar a la página
    driver.get("https://cde.ucr.cjis.gov/LATEST/webapp/#/pages/home")
    #
    # Esperar a que la página cargue completamente
    time.sleep(5)
    #
    # Localizar el enlace y hacer clic
    driver.find_element(By.LINK_TEXT, "Documents & Downloads").click()
    #
    # Esperar a que la página cargue
    time.sleep(5)
    #
    # Localizar la sección "National Incident-Based Reporting System (NIBRS) Tables"
    seccion_especifica = driver.find_element(By.ID, "nibrsPubTablesDownloads")
    #
    # Esperar a que la página cargue
    time.sleep(5)
    #
    # Elegir la tabla: Victims, año: 2022, location: Florida class="select-button top"
    seccion_especifica.find_element(
    By.ID, "dwnnibrs-download-select").click()
    time.sleep(1)
    selectVictims = driver.find_element(By.TAG_NAME, "nb-option-list")
    selectVictims.find_element(
    By.XPATH, "//nb-option[text()=' Victims ']").click()
    seccion_especifica.find_element(By.ID, "dwnnibrscol-year-select").click()
    time.sleep(1)
    seccion_especifica.find_element(
    By.XPATH, "//nb-option[contains(text(),'2022')]").click()
    seccion_especifica.find_element(By.ID, "dwnnibrsloc-select").click()
    time.sleep(1)
    seccion_especifica.find_element(
    By.XPATH, "//nb-option[contains(text(),'Florida')]").click()
    #
    # Esperar a que la página cargue
    time.sleep(2)
    #
    # Descargar el archivo
    driver.find_element(By.ID, "nibrs-download-button").click()
    #
    # Esperar a que se complete la descarga (ajusta el tiempo según sea necesario)
    time.sleep(5)
    driver.quit()

    with zipfile.ZipFile("victims.zip") as z:
        with z.open("Victims_Age_by_Offense_Category_2022.xlsx") as f:
            df = pd.read_excel(f, header=4, skipfooter=1)
            print(df)    # print the first 5 rows

    # Filtrar por la categoría "Crimes Against Property"
    indice_crimes_against_property = df[df['Unnamed: 0']
                                        == "Crimes Against Property"].index[0]
    # Generar CSV sin totales, footer ni index
    print(indice_crimes_against_property)
    df = df.iloc[indice_crimes_against_property:]
    print(df)
    df.to_csv(
        "Crimes_Against_Property.csv", index=False, header=True, sep=';', encoding='utf-8')

    print("Prueba completada. El archivo CSV se ha generado correctamente.")

finally:
    # Cerrar el navegador al finalizar
    driver.quit()
