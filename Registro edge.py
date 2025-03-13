from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Configurar WebDriver para Microsoft Edge
driver = webdriver.Edge()
driver.maximize_window()
wait = WebDriverWait(driver, 15)

# URL de inicio de sesión
url = "https://caprendizaje.sena.edu.co/sgva/SGVA_Diseno/pag/login.aspx"
driver.get(url)

# Iniciar sesión
def login():
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))).send_keys("900729437")
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))).send_keys("719D62B4")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']"))).click()
        print("✅ Inicio de sesión exitoso.")
    except Exception as e:
        print("❌ Error al iniciar sesión:", e)
        driver.quit()
        exit()

login()

# Ir a solicitudes
solicitudes_url = "https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/"
driver.get(solicitudes_url)
time.sleep(3)

# Buscar aprendices
buscar_aprendices_url = "https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/BuscarAprendices"
driver.get(buscar_aprendices_url)
time.sleep(3)

# Ver aprendices para la última solicitud creada
try:
    ver_aprendices_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='solicitud_aprendices']")))
    driver.execute_script("arguments[0].scrollIntoView();", ver_aprendices_btn)
    driver.execute_script("arguments[0].click();", ver_aprendices_btn)
    print("✅ Ver aprendices.")
    time.sleep(2)
except Exception as e:
    print(f"❌ Error al intentar ver aprendices: {e}")
    driver.quit()
    exit()

# Extraer datos de aprendices
aprendices_data = []
try:
    filas = driver.find_elements(By.XPATH, "//tbody[@id='filasTabla']/tr")
    for fila in filas:
        columnas = fila.find_elements(By.TAG_NAME, "td")
        datos_aprendiz = [columna.text for columna in columnas]
        aprendices_data.append(datos_aprendiz)

        # Hacer clic en 'Detalle' para ver más datos del aprendiz
        detalle_btn = fila.find_element(By.XPATH, ".//button[@id='solicitud_aprendiz']")
        detalle_btn.click()
        time.sleep(2)

        # Extraer número de teléfono del aprendiz
        telefono = wait.until(EC.presence_of_element_located((By.XPATH, "//label[@id='lbl_apr_telefono']"))).text
        datos_aprendiz.append(telefono)
        print(f"✅ Datos del aprendiz {datos_aprendiz[3]} copiados.")

    # Guardar datos en Excel
    df = pd.DataFrame(aprendices_data, columns=["Acción", "Tipo de Documento", "Documento", "Nombre", "Ciudad", "Fecha Inicio", "Fecha Fin", "Fecha Máxima", "Teléfono"])
    df.to_excel("aprendices.xlsx", index=False)
    print("✅ Datos de aprendices guardados en Excel.")
except Exception as e:
    print(f"❌ Error al extraer datos de aprendices: {e}")
finally:
    driver.quit()