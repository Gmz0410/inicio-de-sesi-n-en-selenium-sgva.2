from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os

# Configurar WebDriver para Microsoft Edge
driver = webdriver.Edge()
driver.maximize_window()
wait = WebDriverWait(driver, 15)

# Archivo de texto para almacenar las combinaciones de departamentos y ciudades ya utilizadas
used_locations_file = "used_locations.txt"

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

# Leer ubicaciones ya utilizadas desde el archivo
def leer_ubicaciones_utilizadas():
    if os.path.exists(used_locations_file):
        with open(used_locations_file, "r") as file:
            ubicaciones = file.read().splitlines()
    else:
        ubicaciones = []
    return set(ubicaciones)

# Guardar nueva ubicación en el archivo
def guardar_ubicacion(departamento, ciudad):
    with open(used_locations_file, "a") as file:
        file.write(f"{departamento} - {ciudad}\n")

ubicaciones_utilizadas = leer_ubicaciones_utilizadas()

# Navegar a la página de creación de solicitud
solicitudes_url = "https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/Crear"
driver.get(solicitudes_url)
time.sleep(3)

# Función para llenar inputs
def fill_input(xpath, value, field_name):
    try:
        field = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView();", field)
        field.clear()
        field.send_keys(value)
        print(f"✅ {field_name} llenado: {value}")
    except Exception as e:
        print(f"⚠️ No se pudo llenar {field_name}: {e}")

# Función para seleccionar una ciudad que no se haya utilizado previamente
def seleccionar_ciudad_no_utilizada(departamento, ciudades):
    for ciudad in ciudades:
        if f"{departamento} - {ciudad}" not in ubicaciones_utilizadas:
            return ciudad
    return None

# Seleccionar departamento y ciudad que no se hayan utilizado previamente
def seleccionar_ubicacion_no_utilizada():
    try:
        wait.until(EC.presence_of_element_located((By.ID, "sel_departamento")))
        departamento_select = Select(driver.find_element(By.ID, "sel_departamento"))
        departamentos = [option.get_attribute("value") for option in departamento_select.options if option.get_attribute("value") != "0"]

        for departamento in departamentos:
            departamento_select.select_by_value(departamento)
            print(f"🔹 Departamento seleccionado: {departamento}")
            time.sleep(2)

            ciudad_select = Select(driver.find_element(By.ID, "sel_ciudad"))
            ciudades = [option.get_attribute("value") for option in ciudad_select.options if option.get_attribute("value") != "0"]

            ciudad = seleccionar_ciudad_no_utilizada(departamento, ciudades)
            if ciudad:
                ciudad_select.select_by_value(ciudad)
                print(f"   ↳ Ciudad seleccionada: {ciudad}")
                return departamento, ciudad

        print("⚠️ No se encontró una combinación de departamento y ciudad no utilizada.")
        return None, None
    except Exception as e:
        print(f"⚠️ Error al seleccionar ubicación: {e}")
        return None, None

# Crear solicitud
def crear_solicitud():
    while True:
        departamento, ciudad = seleccionar_ubicacion_no_utilizada()

        if not departamento or not ciudad:
            print("⚠️ No se pudo encontrar una combinación de departamento y ciudad no utilizada.")
            break

        # Guardar la ubicación utilizada
        guardar_ubicacion(departamento, ciudad)
        
        # Llenar datos de la solicitud
        fill_input("//input[@id='txt_direccion']", "Bogota calle 15 - 23", "Dirección")
        fill_input("//input[@id='txt_contacto']", "Nordicol", "Persona de Contacto")
        fill_input("//input[@id='txt_cedula']", "126477323", "Cédula")
        fill_input("//input[@id='txt_telefono']", "3221234555", "Teléfono")
        fill_input("//input[@id='txt_email']", "Nordicol@gmail.com", "Correo Electrónico")
        time.sleep(2)

        # Buscar y seleccionar especialidades
        buscar_especialidad_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='btn_especialidad_buscar']")))
        driver.execute_script("arguments[0].scrollIntoView();", buscar_especialidad_btn)
        driver.execute_script("arguments[0].click();", buscar_especialidad_btn)
        time.sleep(2)

        # Seleccionar especialidad en el menú desplegable
        especialidad_select = Select(wait.until(EC.presence_of_element_located((By.ID, "sel_especialidad"))))
        especialidad_options = especialidad_select.options
        if len(especialidad_options) > 1:
            especialidad_select.select_by_index(1)
            print(f"✅ Especialidad seleccionada: {especialidad_options[1].text}")
        else:
            print("⚠️ No hay especialidades disponibles para seleccionar.")
            continue

        # Asegurarse de que las competencias se cargan completamente
        time.sleep(3)
        competencias_div = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='div_especialidad_competencias']")))
        competencias = competencias_div.find_elements(By.XPATH, ".//label[@class='empresaLabelTexto']")
        if competencias:
            competencia_texto = competencias[0].text
            print(f"✅ Competencia laboral copiada: {competencia_texto}")

            # Hacer clic en el botón 'Seleccionar especialidad'
            seleccionar_especialidad_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_seleccionar_especialidad")))
            driver.execute_script("arguments[0].scrollIntoView();", seleccionar_especialidad_btn)
            time.sleep(1)  # Añadir un pequeño retraso para asegurarse de que el botón está completamente cargado
            driver.execute_script("arguments[0].click();", seleccionar_especialidad_btn)
            print("✅ Botón 'Seleccionar especialidad' clicado.")

            # Llenar el campo 'Perfil del aspirante'
            fill_input("//textarea[@id='txta_perfil']", "Tecnico prueba", "Perfil del aspirante")

            # Pegar la competencia copiada en el campo 'Funciones a desarrollar'
            fill_input("//textarea[@id='txta_funciones']", competencia_texto, "Funciones a desarrollar")

            # Llenar el campo 'Cantidad de aprendices'
            fill_input("//input[@id='txt_cantidad_aprendices']", "3", "Cantidad de aprendices")

            # Seleccionar la fecha de cierre
            fecha_cierre_input = wait.until(EC.element_to_be_clickable((By.ID, "txt_fecha_cierre")))
            driver.execute_script("arguments[0].scrollIntoView();", fecha_cierre_input)
            fecha_cierre_input.click()
            time.sleep(1)

            # Seleccionar el primer día disponible en el calendario
            try:
                fecha_disponible = driver.find_element(By.XPATH, "//td[@data-handler='selectDay']/a")
                fecha_disponible.click()
                print("✅ Fecha de cierre seleccionada.")
            except Exception as e:
                print(f"⚠️ No se pudo seleccionar la fecha de cierre: {e}")

            # Hacer clic en 'Crear solicitud de aprendices'
            crear_solicitud_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_crear_solicitud")))
            driver.execute_script("arguments[0].scrollIntoView();", crear_solicitud_btn)
            driver.execute_script("arguments[0].click();", crear_solicitud_btn)
            print("✅ Solicitud de aprendices creada.")
            time.sleep(2)

            # Manejar el mensaje de solicitud ya existente
            try:
                modal_respuesta = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@id='modal_respuesta']")))
                if "Ya existe una solicitud abierta" in modal_respuesta.text:
                    cerrar_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_respuesta_cerrar")))
                    driver.execute_script("arguments[0].click();", cerrar_btn)
                    print("⚠️ Ya existe una solicitud abierta. Pasando a la siguiente ciudad.")
                    continue
            except:
                pass

            # Ir a aplicaciones y ver aprendices
            ir_aplicaciones_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Ir a aplicaciones')]")))
            driver.execute_script("arguments[0].scrollIntoView();", ir_aplicaciones_btn)
            driver.execute_script("arguments[0].click();", ir_aplicaciones_btn)
            print("✅ Ir a aplicaciones.")
            time.sleep(2)
            
            ver_aprendices_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Ver Aprendices')]")))
            driver.execute_script("arguments[0].scrollIntoView();", ver_aprendices_btn)
            driver.execute_script("arguments[0].click();", ver_aprendices_btn)
            print("✅ Ver Aprendices.")
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
        else:
            print("⚠️ No hay competencias disponibles para seleccionar.")
            continue

crear_solicitud()
driver.quit()