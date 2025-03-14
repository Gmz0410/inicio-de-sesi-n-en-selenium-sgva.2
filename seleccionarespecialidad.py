from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import json
import os
# configuracion 
Max_solicitudes = 5    # Numero de solicitudes a crear
Tiempo_espera_maximo = 120 # Tiempo de espera maximo en segundos

# Función para cargar combinaciones ya utilizadas
def load_used_combinations():
    if os.path.exists("used_combinations.json"):
        try:
            with open("used_combinations.json", "r") as file:
                return json.load(file)
        except:
            return []
    else:
        return []

# Función para guardar combinaciones utilizadas
def save_used_combination(departamento, ciudad, especialidad):
    combinations = load_used_combinations()
    new_combination = {
        "departamento": departamento,
        "ciudad": ciudad,
        "especialidad": especialidad
    }
    if new_combination not in combinations:
        combinations.append(new_combination)
        with open("used_combinations.json", "w") as file:
            json.dump(combinations, file, indent=4)
        print(f"✅ Combinación guardada: {departamento}-{ciudad}-{especialidad}")

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

# Función para verificar si existe un modal de error
def check_for_error_modal():
    try:
        # Esperar un breve momento para ver si aparece el modal
        error_message = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//span[@id='modal_respuesta' and contains(text(), 'Ya existe una solicitud')]"))
        )
        print("⚠️ Error: Ya existe una solicitud abierta con esta combinación")
        
        # Cerrar el modal
        close_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@id='btn_respuesta_cerrar']"))
        )
        driver.execute_script("arguments[0].click();", close_button)
        print("✅ Modal de error cerrado")
        return True
    except:
        return False

# Función para preparar el formulario para una nueva especialidad
def prepare_form(departamento_value, ciudad_value):
    try:
        # Ir a la página de creación de solicitud
        solicitudes_url = "https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/Crear"
        driver.get(solicitudes_url)
        time.sleep(3)
        
        # Seleccionar departamento
        wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id, 'departamento')]")))
        departamento_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'departamento')]"))
        departamento_select.select_by_value(departamento_value)
        print(f"🔹 Departamento seleccionado: {departamento_value}")
        time.sleep(2)
        
        # Seleccionar ciudad
        ciudad_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'ciudad')]"))
        ciudad_select.select_by_value(ciudad_value)
        print(f"   ↳ Ciudad seleccionada: {ciudad_value}")
        
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
        
        return True
    except Exception as e:
        print(f"❌ Error al preparar el formulario: {e}")
        return False

# Función para conseguir todas las especialidades disponibles
def get_available_specialties():
    try:
        especialidad_select = Select(wait.until(EC.presence_of_element_located((By.ID, "sel_especialidad"))))
        especialidades = [(i, option.get_attribute("value"), option.text.strip()) for i, option in enumerate(especialidad_select.options) if i > 0]
        return especialidades
    except Exception as e:
        print(f"❌ Error al obtener especialidades: {e}")
        return []

# Función para procesar una especialidad específica
def process_specialty(departamento_value, ciudad_value, idx, especialidad_value, especialidad_text, used_combinations):
    try:
        # Verificar si esta combinación ya fue utilizada
        combinacion_actual = {
            "departamento": departamento_value, 
            "ciudad": ciudad_value, 
            "especialidad": especialidad_value
        }
        
        if combinacion_actual in used_combinations:
            print(f"⚠️ Combinación ya utilizada: {departamento_value}-{ciudad_value}-{especialidad_text}")
            return False
        
        # Seleccionar especialidad
        especialidad_select = Select(wait.until(EC.presence_of_element_located((By.ID, "sel_especialidad"))))
        especialidad_select.select_by_index(idx)
        print(f"✅ Especialidad seleccionada: {especialidad_text}")
        
        # Esperar a que las competencias se carguen
        time.sleep(3)
        
        competencias_div = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='div_especialidad_competencias']")))
        competencias = competencias_div.find_elements(By.XPATH, ".//label[@class='empresaLabelTexto']")
        
        if not competencias:
            print("⚠️ No hay competencias disponibles para esta especialidad.")
            save_used_combination(departamento_value, ciudad_value, especialidad_value)
            return False
            
        competencia_texto = competencias[0].text
        print(f"✅ Competencia laboral copiada: {competencia_texto}")
        
        # Hacer clic en el botón 'Seleccionar especialidad'
        seleccionar_especialidad_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_seleccionar_especialidad")))
        driver.execute_script("arguments[0].scrollIntoView();", seleccionar_especialidad_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", seleccionar_especialidad_btn)
        print("✅ Botón 'Seleccionar especialidad' clicado.")
        
        # Llenar los campos restantes
        fill_input("//textarea[@id='txta_perfil']", "Tecnico prueba", "Perfil del aspirante")
        fill_input("//textarea[@id='txta_funciones']", competencia_texto, "Funciones a desarrollar")
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
        print("✅ Solicitud de aprendices enviada.")
        time.sleep(2)
        
        # Verificar si hay un modal de error
        if check_for_error_modal():
            # Guardar esta combinación como utilizada
            save_used_combination(departamento_value, ciudad_value, especialidad_value)
            return False
        
        # Si no hay error, procesar los resultados
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
        
        # Guardar datos en Excel con nombre de departamento y ciudad
        excel_filename = f"aprendices_{departamento_value}_{ciudad_value}_{especialidad_value}.xlsx"
        df = pd.DataFrame(aprendices_data, columns=["Nombre", "Documento", "Programa", "Estado"])
        df.to_excel(excel_filename, index=False)
        print(f"✅ Datos de aprendices guardados en {excel_filename}.")
        
        # Guardar esta combinación como utilizada
        save_used_combination(departamento_value, ciudad_value, especialidad_value)
        return True
        
    except Exception as e:
        print(f"❌ Error al procesar la especialidad {especialidad_text}: {e}")
        save_used_combination(departamento_value, ciudad_value, especialidad_value)
        return False

# Función principal para recorrer departamentos y ciudades
def main():
    try:
        # Obtener la lista de departamentos
        driver.get("https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/Crear")
        time.sleep(3)
        
        departamento_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'departamento')]"))
        departamentos = [(option.get_attribute("value"), option.text.strip()) for option in departamento_select.options if option.get_attribute("value") != "0"]
        
        # Cargar combinaciones ya utilizadas
        used_combinations = load_used_combinations()
        
        for dep_value, dep_text in departamentos:
            # Seleccionar departamento
            departamento_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'departamento')]"))
            departamento_select.select_by_value(dep_value)
            print(f"\n🔹 DEPARTAMENTO: {dep_text} ({dep_value})")
            time.sleep(2)
            
            # Obtener ciudades para este departamento
            ciudad_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'ciudad')]"))
            ciudades = [(option.get_attribute("value"), option.text.strip()) for option in ciudad_select.options if option.get_attribute("value") != "0"]
            
            for ciudad_value, ciudad_text in ciudades:
                print(f"\n   ↳ CIUDAD: {ciudad_text} ({ciudad_value})")
                
                # Preparar el formulario para esta combinación departamento-ciudad
                if not prepare_form(dep_value, ciudad_value):
                    print(f"⚠️ No se pudo preparar el formulario para {dep_text} - {ciudad_text}")
                    continue
                
                # Obtener todas las especialidades disponibles para esta ciudad
                especialidades = get_available_specialties()
                if not especialidades:
                    print(f"⚠️ No hay especialidades disponibles para {dep_text} - {ciudad_text}")
                    continue
                
                print(f"ℹ️ Se encontraron {len(especialidades)} especialidades para {ciudad_text}")
                
                # Para cada especialidad, intentar crear una solicitud
                especialidades_procesadas = False
                
                for idx, especialidad_value, especialidad_text in especialidades:
                    print(f"\n      ↳ ESPECIALIDAD: {especialidad_text} ({especialidad_value})")
                    
                    # Procesar esta especialidad
                    result = process_specialty(dep_value, ciudad_value, idx, especialidad_value, especialidad_text, used_combinations)
                    
                    if result:
                        especialidades_procesadas = True
                        print(f"✅ Solicitud creada exitosamente para {especialidad_text}")
                    
                    # Recargar combinaciones usadas después de cada intento
                    used_combinations = load_used_combinations()
                    
                    # Preparar el formulario para la siguiente especialidad
                    if idx < len(especialidades) - 1:  # Si no es la última especialidad
                        if not prepare_form(dep_value, ciudad_value):
                            print(f"⚠️ No se pudo preparar el formulario para la siguiente especialidad")
                            break
                
                if not especialidades_procesadas:
                    print(f"⚠️ No se pudieron procesar especialidades para {dep_text} - {ciudad_text}")
                
                # Volver a la página de creación para la siguiente ciudad
                driver.get("https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/Crear")
                time.sleep(3)
    
    except Exception as e:
        print(f"❌ Error general: {e}")
    finally:
        print("\n✅ Proceso finalizado.")
        driver.quit()

# Ejecutar el script
if __name__ == "__main__":
    main()  