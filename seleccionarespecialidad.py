from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import json
import os
import sys
import io

# Configurar la codificación de salida para evitar problemas con emojis
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Definir símbolos seguros para usar en la consola de Windows
CHECK = "[OK]" 
ERROR = "[ERROR]"
WARNING = "[AVISO]"
INFO = "[INFO]"
ARROW = "-->"

# configuracion 
Max_solicitudes = 1    # Numero de solicitudes a crear (cambiado a 1)
Tiempo_espera_maximo = 120 # Tiempo de espera maximo en segundos
Max_combinaciones_por_ciudad = 4  # Número máximo de combinaciones a guardar automáticamente por ciudad

# Contador global para combinaciones procesadas por ciudad
combinaciones_procesadas_por_ciudad = {}
# Variable global para rastrear si ya se creó una solicitud exitosa
solicitud_exitosa_creada = False

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
def save_used_combination(departamento, ciudad, especialidad, centro_value=None, centro_text=None, forzar_guardado=False):
    combinations = load_used_combinations()
    new_combination = {
        "departamento": departamento,
        "ciudad": ciudad,
        "especialidad": especialidad
    }
    
    # Añadir información del centro si está disponible
    if centro_value and centro_text:
        new_combination["centro_value"] = centro_value
        new_combination["centro_text"] = centro_text
        
    if new_combination not in combinations:
        combinations.append(new_combination)
        with open("used_combinations.json", "w") as file:
            json.dump(combinations, file, indent=4)
        
        centro_info = f" - Centro: {centro_text}" if centro_text else ""
        if forzar_guardado:
            print(f"{INFO} Combinación guardada automáticamente (dentro del límite): {departamento}-{ciudad}-{especialidad}{centro_info}")
        else:
            print(f"{CHECK} Combinación guardada: {departamento}-{ciudad}-{especialidad}{centro_info}")
    elif forzar_guardado:
        print(f"{INFO} Combinación ya existente (no se guarda de nuevo): {departamento}-{ciudad}-{especialidad}")

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
        print(f"{CHECK} Inicio de sesión exitoso.")
    except Exception as e:
        print(f"{ERROR} Error al iniciar sesión:", e)
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
        print(f"{CHECK} {field_name} llenado: {value}")
    except Exception as e:
        print(f"{WARNING} No se pudo llenar {field_name}: {e}")

# Función para verificar si existe un modal de error
def check_for_error_modal():
    try:
        # Esperar un breve momento para ver si aparece el modal
        error_message = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//span[@id='modal_respuesta' and contains(text(), 'Ya existe una solicitud')]"))
        )
        print(f"{WARNING} Error: Ya existe una solicitud abierta con esta combinación")
        
        # Cerrar el modal
        close_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@id='btn_respuesta_cerrar']"))
        )
        driver.execute_script("arguments[0].click();", close_button)
        print(f"{CHECK} Modal de error cerrado")
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
        print(f"{INFO} DEPARTAMENTO: {departamento_value}")
        time.sleep(2)
        
        # Seleccionar ciudad
        ciudad_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'ciudad')]"))
        ciudad_select.select_by_value(ciudad_value)
        print(f"{ARROW} CIUDAD: {ciudad_value}")
        
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
        print(f"{WARNING} Error al preparar el formulario: {e}")
        return False

# Función para conseguir todas las especialidades disponibles
def get_available_specialties():
    try:
        especialidad_select = Select(wait.until(EC.presence_of_element_located((By.ID, "sel_especialidad"))))
        especialidades = [(i, option.get_attribute("value"), option.text.strip()) for i, option in enumerate(especialidad_select.options) if i > 0]
        return especialidades
    except Exception as e:
        print(f"{WARNING} Error al obtener especialidades: {e}")
        return []

# Función para verificar y actualizar el contador de combinaciones por ciudad
def verificar_limite_combinaciones(departamento_value, ciudad_value, especialidad_value, especialidad_text, centro_value=None, centro_text=None):
    # Crear clave única para esta ciudad
    ciudad_key = f"{departamento_value}_{ciudad_value}"
    
    # Inicializar contador si no existe
    if ciudad_key not in combinaciones_procesadas_por_ciudad:
        combinaciones_procesadas_por_ciudad[ciudad_key] = 0
    
    # Incrementar contador
    combinaciones_procesadas_por_ciudad[ciudad_key] += 1
    
    # Verificar si estamos dentro del límite
    if combinaciones_procesadas_por_ciudad[ciudad_key] <= Max_combinaciones_por_ciudad:
        # Si estamos dentro del límite, guardar esta combinación automáticamente
        save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text, forzar_guardado=True)
        
        # Si es exactamente el límite, informar
        if combinaciones_procesadas_por_ciudad[ciudad_key] == Max_combinaciones_por_ciudad:
            print(f"{INFO} Se alcanzó el límite de {Max_combinaciones_por_ciudad} combinaciones automáticas para {ciudad_key}")
            
        return True  # Dentro del límite
    
    return False  # Fuera del límite

# Función para manejar el modal de selección de centro
def handle_centro_modal():
    try:
        # Verificar si aparece el modal de selección de centro
        modal_message = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Por favor seleccione un centro de formación')]"))
        )
        print(f"{WARNING} Modal detectado: 'Por favor seleccione un centro de formación'")
        
        # Cerrar el modal
        cerrar_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Cerrar')]"))
        )
        driver.execute_script("arguments[0].click();", cerrar_btn)
        print(f"{CHECK} Modal cerrado correctamente")
        return True
    except:
        return False

# Función para procesar centros de formación
def procesar_centros_formacion(departamento_value, ciudad_value, idx, especialidad_value, especialidad_text, used_combinations):
    global solicitud_exitosa_creada
    
    # Verificar si ya se creó una solicitud exitosa
    if solicitud_exitosa_creada:
        print(f"{INFO} Ya se creó una solicitud exitosa. Omitiendo esta especialidad.")
        return False
        
    try:
        # Seleccionar especialidad primero
        especialidad_select = Select(wait.until(EC.presence_of_element_located((By.ID, "sel_especialidad"))))
        especialidad_select.select_by_index(idx)
        print(f"{CHECK} Especialidad seleccionada: {especialidad_text}")
        
        # Esperar a que las competencias se carguen
        time.sleep(3)
        
        # Verificar si existe el selector de centros de formación
        try:
            print(f"{INFO} Buscando selector de centros de formación...")
            
            # Esperar explícitamente a que aparezca el selector de centros (si existe)
            centro_select = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "sel_centro"))
            )
            
            # Verificar si el selector está visible
            if not centro_select.is_displayed():
                print(f"{WARNING} El selector de centros existe pero no está visible.")
                
                # Intentar hacer clic en el botón 'Seleccionar especialidad' para ver si aparece el modal
                try:
                    seleccionar_especialidad_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_seleccionar_especialidad")))
                    driver.execute_script("arguments[0].scrollIntoView();", seleccionar_especialidad_btn)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", seleccionar_especialidad_btn)
                    print(f"{CHECK} Botón 'Seleccionar especialidad' clicado para verificar si se requiere centro.")
                    
                    # Verificar si aparece el modal de selección de centro
                    if handle_centro_modal():
                        print(f"{WARNING} Se requiere seleccionar un centro de formación pero el selector no está visible.")
                        save_used_combination(departamento_value, ciudad_value, especialidad_value)
                        return False
                except Exception as e:
                    print(f"{WARNING} Error al verificar si se requiere centro: {e}")
                
                # Si no apareció el modal, procesar normalmente
                verificar_limite_combinaciones(departamento_value, ciudad_value, especialidad_value, especialidad_text)
                result = process_specialty(departamento_value, ciudad_value, idx, especialidad_value, especialidad_text, used_combinations, None, None)
                save_used_combination(departamento_value, ciudad_value, especialidad_value)
                return result
            
            print(f"{CHECK} Selector de centros encontrado y visible.")
            
            # Convertir a objeto Select para trabajar con él
            centro_select_obj = Select(centro_select)
            
            # Obtener todos los centros disponibles (excluyendo la opción "Seleccione un centro")
            centros = [(option.get_attribute("value"), option.text.strip()) 
                      for option in centro_select_obj.options 
                      if option.get_attribute("value") != "0"]
            
            if not centros:
                print(f"{WARNING} No hay centros de formación disponibles para esta especialidad.")
                
                # Intentar hacer clic en el botón 'Seleccionar especialidad' para ver si aparece el modal
                try:
                    seleccionar_especialidad_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_seleccionar_especialidad")))
                    driver.execute_script("arguments[0].scrollIntoView();", seleccionar_especialidad_btn)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", seleccionar_especialidad_btn)
                    print(f"{CHECK} Botón 'Seleccionar especialidad' clicado para verificar si se requiere centro.")
                    
                    # Verificar si aparece el modal de selección de centro
                    if handle_centro_modal():
                        print(f"{WARNING} Se requiere seleccionar un centro de formación pero no hay centros disponibles.")
                        save_used_combination(departamento_value, ciudad_value, especialidad_value)
                        return False
                except Exception as e:
                    print(f"{WARNING} Error al verificar si se requiere centro: {e}")
                
                # Si no apareció el modal, procesar normalmente
                verificar_limite_combinaciones(departamento_value, ciudad_value, especialidad_value, especialidad_text)
                result = process_specialty(departamento_value, ciudad_value, idx, especialidad_value, especialidad_text, used_combinations, None, None)
                save_used_combination(departamento_value, ciudad_value, especialidad_value)
                return result
            
            print(f"{INFO} Se encontraron {len(centros)} centros de formación para esta especialidad:")
            for i, (c_value, c_text) in enumerate(centros):
                print(f"   {i+1}. {c_text} ({c_value})")
            
            # Procesar cada centro de formación
            for centro_value, centro_text in centros:
                # Verificar si ya se creó una solicitud exitosa
                if solicitud_exitosa_creada:
                    print(f"{INFO} Ya se creó una solicitud exitosa. Omitiendo los centros restantes.")
                    break
                    
                print(f"\n{ARROW} Procesando centro: {centro_text} ({centro_value})")
                
                # Verificar límite de combinaciones antes de procesar este centro (solo para guardar automáticamente)
                verificar_limite_combinaciones(departamento_value, ciudad_value, especialidad_value, especialidad_text, centro_value, centro_text)
                
                # Seleccionar este centro
                try:
                    # Asegurarse de que el selector sigue siendo accesible
                    centro_select_obj = Select(driver.find_element(By.ID, "sel_centro"))
                    centro_select_obj.select_by_value(centro_value)
                    print(f"{CHECK} Centro de formación seleccionado: {centro_text}")
                    time.sleep(2)  # Dar más tiempo para que se actualice la interfaz
                    
                    # Verificar que se haya seleccionado correctamente
                    selected_option = centro_select_obj.first_selected_option
                    selected_value = selected_option.get_attribute("value")
                    selected_text = selected_option.text.strip()
                    print(f"{INFO} Centro seleccionado actualmente: {selected_text} ({selected_value})")
                    
                    if selected_value != centro_value:
                        print(f"{WARNING} No se pudo seleccionar correctamente el centro de formación.")
                        continue
                except Exception as e:
                    print(f"{WARNING} Error al seleccionar centro de formación: {e}")
                    continue
                
                # Verificar si hay competencias disponibles
                try:
                    competencias_div = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='div_especialidad_competencias']")))
                    
                    # Buscar todas las etiquetas que contienen texto de competencias
                    competencias_labels = competencias_div.find_elements(By.XPATH, ".//label[@class='empresaLabelTexto']")
                    
                    if not competencias_labels:
                        # Intentar con otro selector si el primero no funciona
                        competencias_items = competencias_div.find_elements(By.XPATH, ".//*[contains(@class, 'competencia')]")
                        if not competencias_items:
                            print(f"{WARNING} No hay competencias disponibles para esta especialidad con este centro.")
                            save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
                            continue
                    
                    # Extraer el texto de las competencias para verificar si tienen contenido
                    textos_competencias = []
                    for label in competencias_labels:
                        texto = label.text.strip()
                        if texto:
                            textos_competencias.append(texto)
                    
                    if textos_competencias:
                        print(f"{CHECK} Se encontraron {len(textos_competencias)} competencias para este centro.")
                        print(f"{INFO} Primera competencia: {textos_competencias[0][:50]}...")
                    else:
                        print(f"{WARNING} Las etiquetas de competencias están vacías.")
                        # Intentar con otro selector
                        competencias_items = competencias_div.find_elements(By.XPATH, ".//*[contains(@class, 'competencia')]")
                        if competencias_items:
                            for item in competencias_items:
                                texto = item.text.strip()
                                if texto:
                                    textos_competencias.append(texto)
                            
                            if textos_competencias:
                                print(f"{CHECK} Se encontraron {len(textos_competencias)} competencias (selector alternativo).")
                                print(f"{INFO} Primera competencia: {textos_competencias[0][:50]}...")
                            else:
                                print(f"{WARNING} No se encontró texto en las competencias alternativas.")
                                save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
                                continue
                        else:
                            print(f"{WARNING} No hay competencias disponibles para esta especialidad con este centro.")
                            save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
                            continue
                except Exception as e:
                    print(f"{WARNING} Error al verificar competencias: {e}")
                    save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
                    continue
                
                # Hacer clic en el botón 'Seleccionar especialidad'
                try:
                    seleccionar_especialidad_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_seleccionar_especialidad")))
                    driver.execute_script("arguments[0].scrollIntoView();", seleccionar_especialidad_btn)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", seleccionar_especialidad_btn)
                    print(f"{CHECK} Botón 'Seleccionar especialidad' clicado.")
                    
                    # Verificar si aparece el modal de selección de centro
                    if handle_centro_modal():
                        print(f"{WARNING} Se requiere seleccionar un centro de formación. Intentando seleccionar nuevamente.")
                        # Intentar seleccionar el centro nuevamente
                        try:
                            centro_select_obj = Select(driver.find_element(By.ID, "sel_centro"))
                            centro_select_obj.select_by_value(centro_value)
                            print(f"{CHECK} Centro de formación seleccionado nuevamente: {centro_text}")
                            time.sleep(2)
                            
                            # Hacer clic en el botón 'Seleccionar especialidad' otra vez
                            seleccionar_especialidad_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_seleccionar_especialidad")))
                            driver.execute_script("arguments[0].scrollIntoView();", seleccionar_especialidad_btn)
                            time.sleep(1)
                            driver.execute_script("arguments[0].click();", seleccionar_especialidad_btn)
                            print(f"{CHECK} Botón 'Seleccionar especialidad' clicado nuevamente.")
                            
                            # Verificar si aparece el modal otra vez
                            if handle_centro_modal():
                                print(f"{WARNING} No se pudo seleccionar el centro correctamente después de varios intentos.")
                                save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
                                continue
                        except Exception as e:
                            print(f"{WARNING} Error al intentar seleccionar el centro nuevamente: {e}")
                            save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
                            continue
                except Exception as e:
                    print(f"{WARNING} Error al hacer clic en 'Seleccionar especialidad': {e}")
                    save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
                    continue
                
                # Procesar esta combinación con el centro seleccionado
                result = process_specialty(
                    departamento_value, 
                    ciudad_value, 
                    idx, 
                    especialidad_value, 
                    especialidad_text, 
                    used_combinations,
                    centro_value,
                    centro_text
                )
                
                if result:
                    print(f"{CHECK} Solicitud creada exitosamente para centro: {centro_text}")
                    # Si ya se creó una solicitud exitosa, terminar
                    if solicitud_exitosa_creada:
                        return True
                
                # Si no es el último centro y no se ha creado una solicitud exitosa, volver a preparar el formulario
                if centro_value != centros[-1][0] and not solicitud_exitosa_creada:
                    if not prepare_form(departamento_value, ciudad_value):
                        print(f"{WARNING} No se pudo preparar el formulario para el siguiente centro")
                        break
                        
                    # Volver a seleccionar la especialidad
                    try:
                        especialidad_select = Select(wait.until(EC.presence_of_element_located((By.ID, "sel_especialidad"))))
                        especialidad_select.select_by_index(idx)
                        print(f"{CHECK} Especialidad seleccionada nuevamente: {especialidad_text}")
                        time.sleep(3)
                    except Exception as e:
                        print(f"{WARNING} Error al volver a seleccionar la especialidad: {e}")
                        break
            
            return True
            
        except Exception as e:
            # Si no hay selector de centros, intentar procesar sin centro
            print(f"{WARNING} No se detectó selector de centros: {e}")
            
            # Intentar hacer clic en el botón 'Seleccionar especialidad' para ver si aparece el modal
            try:
                seleccionar_especialidad_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_seleccionar_especialidad")))
                driver.execute_script("arguments[0].scrollIntoView();", seleccionar_especialidad_btn)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", seleccionar_especialidad_btn)
                print(f"{CHECK} Botón 'Seleccionar especialidad' clicado para verificar si se requiere centro.")
                
                # Verificar si aparece el modal de selección de centro
                if handle_centro_modal():
                    print(f"{WARNING} Se requiere seleccionar un centro de formación pero no se pudo detectar el selector.")
                    save_used_combination(departamento_value, ciudad_value, especialidad_value)
                    return False
            except Exception as e:
                print(f"{WARNING} Error al verificar si se requiere centro: {e}")
            
            # Si no apareció el modal, procesar normalmente
            # Verificar límite de combinaciones antes de procesar (solo para guardar automáticamente)
            verificar_limite_combinaciones(departamento_value, ciudad_value, especialidad_value, especialidad_text)
            # Procesar la especialidad sin centro
            result = process_specialty(departamento_value, ciudad_value, idx, especialidad_value, especialidad_text, used_combinations, None, None)
            # Guardar esta combinación como utilizada incluso si hay error
            save_used_combination(departamento_value, ciudad_value, especialidad_value)
            return result
            
    except Exception as e:
        print(f"{ERROR} Error al procesar centros de formación: {e}")
        # Guardar la combinación incluso si hay error
        verificar_limite_combinaciones(departamento_value, ciudad_value, especialidad_value, especialidad_text)
        save_used_combination(departamento_value, ciudad_value, especialidad_value)
        return False

# Función para procesar una especialidad específica
def process_specialty(departamento_value, ciudad_value, idx, especialidad_value, especialidad_text, used_combinations, centro_value=None, centro_text=None):
    global solicitud_exitosa_creada
    
    try:
        # Verificar si ya se creó una solicitud exitosa
        if solicitud_exitosa_creada:
            print(f"{INFO} Ya se creó una solicitud exitosa. Omitiendo esta especialidad.")
            return False
            
        # Verificar si esta combinación ya fue utilizada
        combinacion_actual = {
            "departamento": departamento_value, 
            "ciudad": ciudad_value, 
            "especialidad": especialidad_value
        }
        
        # Añadir información del centro si está disponible
        if centro_value and centro_text:
            combinacion_actual["centro_value"] = centro_value
            combinacion_actual["centro_text"] = centro_text
        
        # Verificar si estamos dentro del límite de combinaciones para esta ciudad
        # Esto solo es para guardar automáticamente las primeras combinaciones
        verificar_limite_combinaciones(departamento_value, ciudad_value, especialidad_value, especialidad_text, centro_value, centro_text)
        
        if combinacion_actual in used_combinations:
            centro_info = f" - Centro: {centro_text}" if centro_text else ""
            print(f"{WARNING} Combinación ya utilizada: {departamento_value}-{ciudad_value}-{especialidad_text}{centro_info}")
            return False
        
        # Obtener las competencias laborales de manera más robusta
        competencia_texto = "Competencias técnicas relacionadas con la especialidad"  # Valor por defecto
        try:
            # Esperar a que el div de competencias esté presente
            competencias_div = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='div_especialidad_competencias']")))
            
            # Buscar todas las etiquetas que contienen texto de competencias
            competencias_labels = competencias_div.find_elements(By.XPATH, ".//label[@class='empresaLabelTexto']")
            
            if competencias_labels:
                # Extraer el texto de todas las competencias y unirlas
                textos_competencias = []
                for label in competencias_labels:
                    texto = label.text.strip()
                    if texto:
                        textos_competencias.append(texto)
                
                if textos_competencias:
                    competencia_texto = " | ".join(textos_competencias)
                    print(f"{CHECK} Competencias laborales encontradas: {len(textos_competencias)}")
                    print(f"{INFO} Primera competencia: {textos_competencias[0][:50]}...")
                else:
                    print(f"{WARNING} Las etiquetas de competencias están vacías, usando texto genérico.")
            else:
                # Intentar con otro selector si el primero no funciona
                competencias_items = competencias_div.find_elements(By.XPATH, ".//*[contains(@class, 'competencia')]")
                if competencias_items:
                    textos_competencias = []
                    for item in competencias_items:
                        texto = item.text.strip()
                        if texto:
                            textos_competencias.append(texto)
                    
                    if textos_competencias:
                        competencia_texto = " | ".join(textos_competencias)
                        print(f"{CHECK} Competencias laborales encontradas (selector alternativo): {len(textos_competencias)}")
                        print(f"{INFO} Primera competencia: {textos_competencias[0][:50]}...")
                    else:
                        print(f"{WARNING} No se encontró texto en las competencias alternativas, usando texto genérico.")
                else:
                    print(f"{WARNING} No se encontraron competencias, usando texto genérico.")
        except Exception as e:
            print(f"{WARNING} Error al obtener competencias: {e}")
        
        # Verificar que el texto de competencia no esté vacío
        if not competencia_texto or competencia_texto.isspace():
            competencia_texto = "Competencias técnicas relacionadas con la especialidad"
            print(f"{WARNING} Texto de competencias vacío, usando texto genérico.")
        
        print(f"{CHECK} Competencia laboral a utilizar: {competencia_texto[:100]}...")
        
        # Llenar los campos del formulario
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
            print(f"{CHECK} Fecha de cierre seleccionada.")
        except Exception as e:
            print(f"{WARNING} No se pudo seleccionar la fecha de cierre: {e}")
            # Guardar esta combinación como utilizada incluso si hay error
            save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
            return False
        
        # Hacer clic en 'Crear solicitud de aprendices'
        crear_solicitud_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_crear_solicitud")))
        driver.execute_script("arguments[0].scrollIntoView();", crear_solicitud_btn)
        driver.execute_script("arguments[0].click();", crear_solicitud_btn)
        print(f"{CHECK} Solicitud de aprendices enviada.")
        time.sleep(2)
        
        # Verificar si hay un modal de error
        if check_for_error_modal():
            # Guardar esta combinación como utilizada
            save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
            return False
        
        # Si no hay error, procesar los resultados
        # Ir a aplicaciones y ver aprendices
        ir_aplicaciones_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Ir a aplicaciones')]")))
        driver.execute_script("arguments[0].scrollIntoView();", ir_aplicaciones_btn)
        driver.execute_script("arguments[0].click();", ir_aplicaciones_btn)
        print(f"{CHECK} Ir a aplicaciones.")
        time.sleep(2)
        
        ver_aprendices_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Ver Aprendices')]")))
        driver.execute_script("arguments[0].scrollIntoView();", ver_aprendices_btn)
        driver.execute_script("arguments[0].click();", ver_aprendices_btn)
        print(f"{CHECK} Ver Aprendices.")
        time.sleep(2)
        
        # Extraer datos de aprendices
        aprendices_data = []
        filas = driver.find_elements(By.XPATH, "//table[@id='tabla_aprendices']//tr")
        for fila in filas[1:]:
            columnas = fila.find_elements(By.TAG_NAME, "td")
            datos_aprendiz = [columna.text for columna in columnas]
            aprendices_data.append(datos_aprendiz)
        
        # Guardar datos en Excel con nombre de departamento, ciudad y centro (si está disponible)
        centro_suffix = f"_{centro_value}" if centro_value else ""
        excel_filename = f"aprendices_{departamento_value}_{ciudad_value}_{especialidad_value}{centro_suffix}.xlsx"
        df = pd.DataFrame(aprendices_data, columns=["Nombre", "Documento", "Programa", "Estado"])
        df.to_excel(excel_filename, index=False)
        print(f"{CHECK} Datos de aprendices guardados en {excel_filename}.")
        
        # Marcar que ya se creó una solicitud exitosa
        solicitud_exitosa_creada = True
        print(f"{CHECK} ¡SOLICITUD EXITOSA CREADA! Se terminará el proceso después de esta solicitud.")
        
        # Guardar esta combinación como utilizada
        save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
        return True
        
    except Exception as e:
        print(f"{WARNING} Error al procesar la especialidad {especialidad_text}: {e}")
        # Guardar esta combinación como utilizada incluso si hay error
        save_used_combination(departamento_value, ciudad_value, especialidad_value, centro_value, centro_text)
        return False

# Función principal para recorrer departamentos y ciudades
def main():
    global solicitud_exitosa_creada
    
    try:
        # Obtener la lista de departamentos
        driver.get("https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/Crear")
        time.sleep(3)
        
        departamento_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'departamento')]"))
        departamentos = [(option.get_attribute("value"), option.text.strip()) for option in departamento_select.options if option.get_attribute("value") != "0"]
        
        # Reordenar los departamentos para comenzar con ANTIOQUIA
        antioquia_index = -1
        for i, (value, text) in enumerate(departamentos):
            if text == "ANTIOQUIA":
                antioquia_index = i
                break
        
        if antioquia_index != -1:
            # Mover ANTIOQUIA al principio de la lista
            antioquia = departamentos.pop(antioquia_index)
            departamentos.insert(0, antioquia)
            print(f"{INFO} Se ha priorizado el departamento de ANTIOQUIA para la búsqueda.")
        
        # Cargar combinaciones ya utilizadas
        used_combinations = load_used_combinations()
        print(f"{INFO} Se cargaron {len(used_combinations)} combinaciones ya utilizadas.")
        
        # Contador global para combinaciones procesadas
        global combinaciones_procesadas_por_ciudad
        
        for dep_value, dep_text in departamentos:
            # Verificar si ya se creó una solicitud exitosa
            if solicitud_exitosa_creada:
                print(f"{INFO} Ya se creó una solicitud exitosa. Terminando el proceso.")
                break
                
            # Seleccionar departamento
            departamento_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'departamento')]"))
            departamento_select.select_by_value(dep_value)
            print(f"\n{INFO} DEPARTAMENTO: {dep_text} ({dep_value})")
            time.sleep(2)
            
            # Obtener ciudades para este departamento
            ciudad_select = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'ciudad')]"))
            ciudades = [(option.get_attribute("value"), option.text.strip()) for option in ciudad_select.options if option.get_attribute("value") != "0"]
            
            # Si estamos en ANTIOQUIA, priorizar ALEJANDRIA
            if dep_text == "ANTIOQUIA":
                alejandria_index = -1
                for i, (value, text) in enumerate(ciudades):
                    if text == "ALEJANDRIA":
                        alejandria_index = i
                        break
                
                if alejandria_index != -1:
                    # Mover ALEJANDRIA al principio de la lista
                    alejandria = ciudades.pop(alejandria_index)
                    ciudades.insert(0, alejandria)
                    print(f"{INFO} Se ha priorizado la ciudad de ALEJANDRIA para la búsqueda.")
            
            for ciudad_value, ciudad_text in ciudades:
                # Verificar si ya se creó una solicitud exitosa
                if solicitud_exitosa_creada:
                    print(f"{INFO} Ya se creó una solicitud exitosa. Pasando a la siguiente ciudad.")
                    break
                    
                # Reiniciar contador para esta ciudad
                ciudad_key = f"{dep_value}_{ciudad_value}"
                combinaciones_procesadas_por_ciudad[ciudad_key] = 0
                
                print(f"\n   {ARROW} CIUDAD: {ciudad_text} ({ciudad_value})")
                print(f"{INFO} Se guardarán automáticamente las primeras {Max_combinaciones_por_ciudad} combinaciones para esta ciudad")
                
                # Preparar el formulario para esta combinación departamento-ciudad
                if not prepare_form(dep_value, ciudad_value):
                    print(f"{WARNING} No se pudo preparar el formulario para {dep_text} - {ciudad_text}")
                    continue
                
                # Obtener todas las especialidades disponibles para esta ciudad
                especialidades = get_available_specialties()
                if not especialidades:
                    print(f"{WARNING} No hay especialidades disponibles para {dep_text} - {ciudad_text}")
                    continue
                
                print(f"{INFO} Se encontraron {len(especialidades)} especialidades para {ciudad_text}")
                
                # Para cada especialidad, intentar crear una solicitud
                especialidades_procesadas = False
                combinaciones_guardadas = 0
                
                for idx, especialidad_value, especialidad_text in especialidades:
                    # Verificar si ya se creó una solicitud exitosa
                    if solicitud_exitosa_creada:
                        print(f"{INFO} Ya se creó una solicitud exitosa. Pasando a la siguiente especialidad.")
                        break
                        
                    # Verificar si ya alcanzamos el límite de combinaciones automáticas para esta ciudad
                    if combinaciones_procesadas_por_ciudad[ciudad_key] >= Max_combinaciones_por_ciudad:
                        print(f"{INFO} Se alcanzó el límite de {Max_combinaciones_por_ciudad} combinaciones automáticas para {ciudad_text}.")
                        print(f"{INFO} Se seguirán procesando especialidades, pero solo se guardarán las combinaciones utilizadas con éxito.")
                        
                    print(f"\n      {ARROW} ESPECIALIDAD: {especialidad_text} ({especialidad_value})")
                    
                    # Procesar esta especialidad
                    result = procesar_centros_formacion(dep_value, ciudad_value, idx, especialidad_value, especialidad_text, used_combinations)
                    
                    if result:
                        especialidades_procesadas = True
                        combinaciones_guardadas += 1
                        print(f"{CHECK} Solicitud creada exitosamente para {especialidad_text}")
                    
                    # Recargar combinaciones usadas después de cada intento
                    used_combinations = load_used_combinations()
                    
                    # Preparar el formulario para la siguiente especialidad
                    if idx < len(especialidades) - 1:  # Si no es la última especialidad
                        if not prepare_form(dep_value, ciudad_value):
                            print(f"{WARNING} No se pudo preparar el formulario para la siguiente especialidad")
                            break
                
                # Si ya se creó una solicitud exitosa, terminar el proceso
                if solicitud_exitosa_creada:
                    print(f"{INFO} Se ha creado una solicitud exitosa. Terminando el proceso.")
                    break
                
                # Mostrar resumen de combinaciones procesadas para esta ciudad
                print(f"{INFO} Se procesaron {combinaciones_procesadas_por_ciudad[ciudad_key]} combinaciones para {ciudad_text}")
                print(f"{INFO} Se guardaron {len(load_used_combinations()) - len(used_combinations)} nuevas combinaciones en esta ciudad")
                
                # Volver a la página de creación para la siguiente ciudad
                driver.get("https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/Crear")
                time.sleep(3)
            
            # Si ya se creó una solicitud exitosa, terminar el proceso
            if solicitud_exitosa_creada:
                print(f"{INFO} Se ha creado una solicitud exitosa. Terminando el proceso.")
                break
        
        if solicitud_exitosa_creada:
            print(f"\n{CHECK} Proceso finalizado correctamente. Se creó 1 solicitud exitosa.")
        else:
            print(f"\n{WARNING} Proceso finalizado sin crear ninguna solicitud exitosa.")
            
        print(f"{INFO} Total de combinaciones guardadas: {len(load_used_combinations())}")
        return 0  # Código de salida exitoso
    
    except Exception as e:
        print(f"{ERROR} Error general: {e}")
        return 1  # Código de salida con error
    finally:
        print(f"\n{INFO} Cerrando navegador...")
        try:
            driver.quit()
            print(f"{CHECK} Navegador cerrado correctamente.")
        except Exception as e:
            print(f"{WARNING} Error al cerrar el navegador: {e}")

# Ejecutar el script
if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)  # Asegurar que el script termine con el código de salida adecuado  