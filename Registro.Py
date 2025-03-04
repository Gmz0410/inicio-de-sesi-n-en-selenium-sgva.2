from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time

# Configurar WebDriver
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 30)  # Aumentamos el tiempo de espera a 30 segundos

# URL de inicio de sesión
url = "https://caprendizaje.sena.edu.co/sgva/SGVA_Diseno/pag/login.aspx"
driver.get(url)

# Iniciar sesión
def login():
    try:
        print("Intentando ingresar usuario...")
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))).send_keys("900729437")
        print("Usuario ingresado correctamente.")

        print("Intentando ingresar contraseña...")
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))).send_keys("719D62B4")
        print("Contraseña ingresada correctamente.")

        print("Intentando hacer clic en el botón de inicio de sesión...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']"))).click()
        print("✅ Inicio de sesión exitoso.")
    except TimeoutException:
        print("❌ Tiempo de espera agotado durante el inicio de sesión.")
        driver.save_screenshot("login_error.png")
        driver.quit()
        exit()
    except NoSuchElementException:
        print("❌ Elemento no encontrado durante el inicio de sesión.")
        driver.save_screenshot("login_error.png")
        driver.quit()
        exit()
    except Exception as e:
        print(f"❌ Error inesperado durante el inicio de sesión: {e}")
        driver.save_screenshot("login_error.png")
        driver.quit()
        exit()

login()

# Ir a la página de creación de solicitud
solicitudes_url = "https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/Crear"
driver.get(solicitudes_url)

# Función para llenar inputs
def fill_input(xpath, value, field_name):
    try:
        print(f"Intentando llenar {field_name}...")
        field = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].click();", field)  # Hacer clic usando JavaScript
        field.clear()
        field.send_keys(value)
        print(f"✅ {field_name} llenado: {value}")
    except TimeoutException:
        print(f"⚠️ Tiempo de espera agotado al llenar {field_name}.")
        driver.save_screenshot(f"fill_input_error_{field_name}.png")
    except NoSuchElementException:
        print(f"⚠️ Elemento no encontrado al llenar {field_name}.")
        driver.save_screenshot(f"fill_input_error_{field_name}.png")
    except Exception as e:
        print(f"⚠️ Error inesperado al llenar {field_name}: {e}")
        driver.save_screenshot(f"fill_input_error_{field_name}.png")

# Seleccionar departamento y ciudad
try:
    print("Intentando seleccionar departamento...")
    departamento_select = Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id, 'departamento')]"))))
    departamentos = [option.get_attribute("value") for option in departamento_select.options if option.get_attribute("value") != "0"]

    for departamento in departamentos:
        departamento_select.select_by_value(departamento)
        print(f"🔹 Departamento seleccionado: {departamento}")

        print("Intentando seleccionar ciudad...")
        ciudad_select = Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id, 'ciudad')]"))))
        ciudades = [option.get_attribute("value") for option in ciudad_select.options if option.get_attribute("value") != "0"]

        for ciudad in ciudades:
            ciudad_select.select_by_value(ciudad)
            print(f"   ↳ Ciudad seleccionada: {ciudad}")

            # Llenar datos de la solicitud
            fill_input("//input[@id='txt_direccion']", "Bogota calle 15 - 23", "Dirección")
            fill_input("//input[@id='txt_contacto']", "Nordicol", "Persona de Contacto")
            fill_input("//input[@id='txt_cedula']", "126477323", "Cédula")
            fill_input("//input[@id='txt_telefono']", "3221234555", "Teléfono")
            fill_input("//input[@id='txt_email']", "Nordicol@gmail.com", "Correo Electrónico")

            # Buscar y seleccionar especialidades
            print("Intentando buscar el botón de especialidad...")
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='btn_especialidad_buscar']"))).click()
            print("Botón de especialidad encontrado y clickeado.")

            # Seleccionar una especialidad
            print("Intentando buscar las casillas de verificación...")
            checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='checkbox']")))
            if checkboxes:
                checkboxes[0].click()
                print("✅ Especialidad marcada correctamente.")

                # Copiar las competencias de la especialidad
                print("Intentando copiar las competencias...")
                competencias_xpath = "//div[contains(@class, 'competencias')]"
                competencias_element = wait.until(EC.presence_of_element_located((By.XPATH, competencias_xpath)))
                competencias_texto = competencias_element.text
                print("📋 Competencias copiadas:", competencias_texto)

                # Seleccionar especialidad
                print("Intentando seleccionar la especialidad...")
                seleccionar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Seleccionar Especialidad')]")))
                driver.execute_script("arguments[0].click();", seleccionar_btn)
                print("✅ Especialidad seleccionada correctamente.")
            else:
                print("⚠️ No hay especialidades disponibles.")

            # Seleccionar competencia
            print("Intentando seleccionar competencia...")
            competencia_select = Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id, 'competencia')]"))))
            competencia_options = competencia_select.options
            if len(competencia_options) > 1:
                competencia_select.select_by_index(1)

            # Crear solicitud
            print("Intentando crear la solicitud...")
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Crear Solicitud')]"))).click()
            print("✅ Solicitud creada correctamente.")

            # Ir a aplicaciones y ver aprendices
            print("Intentando ir a aplicaciones...")
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Ir a aplicaciones')]"))).click()
            print("✅ Navegación a aplicaciones exitosa.")

            print("Intentando ver aprendices...")
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Ver Aprendices')]"))).click()
            print("✅ Navegación a aprendices exitosa.")

            # Extraer datos de aprendices
            print("Intentando extraer datos de aprendices...")
            aprendices_data = []
            filas = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='tabla_aprendices']//tr")))
            for fila in filas[1:]:  # Ignorar la primera fila (encabezados)
                columnas = fila.find_elements(By.TAG_NAME, "td")
                datos_aprendiz = [columna.text for columna in columnas]
                aprendices_data.append(datos_aprendiz)

            # Guardar datos en Excel
            print("Intentando guardar datos en Excel...")
            df = pd.DataFrame(aprendices_data, columns=["Nombre", "Documento", "Programa", "Estado"])
            df.to_excel("aprendices.xlsx", index=False)
            print("✅ Datos de aprendices guardados en Excel.")

except TimeoutException:
    print("❌ Tiempo de espera agotado durante el proceso.")
    driver.save_screenshot("process_error.png")
except NoSuchElementException:
    print("❌ Elemento no encontrado durante el proceso.")
    driver.save_screenshot("process_error.png")
except Exception as e:
    print(f"❌ Error inesperado durante el proceso: {e}")
    driver.save_screenshot("process_error.png")
finally:
    driver.quit()  # Cerrar el navegador al finalizar