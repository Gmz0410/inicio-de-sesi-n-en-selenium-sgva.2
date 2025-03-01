from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Configurar WebDriver
driver = webdriver.Chrome()
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

# Ir a la página de creación de solicitud
solicitudes_url = "https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/Crear"
driver.get(solicitudes_url)
time.sleep(3)

# Función para llenar inputs
def fill_input(xpath, value, field_name):
    try:
        field = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].click();", field)
        field.clear()
        field.send_keys(value)
        print(f"✅ {field_name} llenado: {value}")
    except Exception as e:
        print(f"⚠️ No se pudo llenar {field_name}: {e}")

# Seleccionar departamento y ciudad
try:
    wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id, 'departamento')]")))
    departamento_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'departamento')]"))
    departamentos = [option.get_attribute("value") for option in departamento_select.options if option.get_attribute("value") != "0"]

    for departamento in departamentos:
        departamento_select.select_by_value(departamento)
        print(f"🔹 Departamento seleccionado: {departamento}")
        time.sleep(2)

        ciudad_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'ciudad')]"))
        ciudades = [option.get_attribute("value") for option in ciudad_select.options if option.get_attribute("value") != "0"]

        for ciudad in ciudades:
            ciudad_select.select_by_value(ciudad)
            print(f"   ↳ Ciudad seleccionada: {ciudad}")
            
            # Llenar datos de la solicitud
            fill_input("//input[@id='txt_direccion']", "Bogota calle 15 - 23", "Dirección")
            fill_input("//input[@id='txt_contacto']", "Nordicol", "Persona de Contacto")  # FIXED
            fill_input("//input[@id='txt_cedula']", "126477323", "Cédula")
            fill_input("//input[@id='txt_telefono']", "3221234555", "Teléfono")
            fill_input("//input[@id='txt_email']", "Nordicol@gmail.com", "Correo Electrónico")
            time.sleep(2)

            # Buscar y seleccionar especialidades
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='btn_especialidad_buscar']"))).click()  # FIXED
            time.sleep(2)

            especialidades = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
            for especialidad in especialidades:
                especialidad.click()
                time.sleep(1)

            driver.find_element(By.XPATH, "//button[contains(text(), 'Seleccionar Especialidad')]").click()

            # Seleccionar competencia
            competencia_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'competencia')]"))
            competencia_options = competencia_select.options
            if len(competencia_options) > 1:
                competencia_select.select_by_index(1)
            
            # Crear solicitud
            driver.find_element(By.XPATH, "//button[contains(text(), 'Crear Solicitud')]").click()
            time.sleep(2)

            # Ir a aplicaciones y ver aprendices
            driver.find_element(By.XPATH, "//button[contains(text(), 'Ir a aplicaciones')]").click()
            time.sleep(2)
            driver.find_element(By.XPATH, "//button[contains(text(), 'Ver Aprendices')]").click()
            time.sleep(2)
            
            # Extraer datos de aprendices
            aprendices_data = []
            filas = driver.find_elements(By.XPATH, "//table[@id='tabla_aprendices']//tr")
            for fila in filas[1:]:
                columnas = fila.find_elements(By.TAG_NAME, "td")
                datos_aprendiz = [columna.text for columna in columnas]
                aprendices_data.append(datos_aprendiz)

            # Guardar datos en Excel
            df = pd.DataFrame(aprendices_data, columns=["Nombre", "Documento", "Programa", "Estado"])
            df.to_excel("aprendices.xlsx", index=False)
            print("✅ Datos de aprendices guardados en Excel.")

except Exception as e:
    print("❌ Error en el proceso:", e)
finally:
    driver.quit()
