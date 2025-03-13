from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
import pandas as pd
import time
import pyperclip  # Para copiar al portapapeles

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
        time.sleep(2)  # Esperar un poco más después del inicio de sesión
    except Exception as e:
        print("❌ Error al iniciar sesión:", e)
        driver.quit()
        exit()

# Función para cerrar modal si está abierto
def cerrar_modal_si_existe():
    try:
        # Intentar encontrar el modal por su ID
        modal = driver.find_element(By.ID, "modalAprendiz")
        
        # Verificar si el modal está visible
        if "display: block" in modal.get_attribute("style"):
            # Encontrar y hacer clic en el botón de cerrar
            cerrar_btn = driver.find_element(By.XPATH, "//div[@id='modalAprendiz']//button[@data-dismiss='modal']")
            cerrar_btn.click()
            print("✅ Modal cerrado correctamente.")
            time.sleep(1)  # Esperar a que se cierre el modal
    except:
        # Si no encuentra el modal o no está visible, continuar
        pass

# Función modificada para formatear datos para Google Sheets
def formatear_datos_para_sheets(datos):
    try:
        # Convertir datos a DataFrame para facilitar el manejo
        df = pd.DataFrame(datos, columns=["Acción", "Tipo de Documento", "Documento", "Nombre", "Ciudad", 
                                       "Fecha Inicio Lectiva", "Fecha Inicio Productiva", "Fecha Fin Productiva", 
                                       "Teléfono", "Correo", "Programa de Formación", "Especialidad"])
        
        # Guardar como backup local
        df.to_excel("aprendices_backup.xlsx", index=False)
        print("✅ Datos guardados en Excel local como backup.")
        
        # Preparar datos para Google Sheets (omitir columna Acción)
        datos_para_sheets = []
        for fila in datos:
            # Crear nueva fila con los datos en el orden correcto:
            # B: Tipo de Documento, C: Documento, D: Nombre, E: Ciudad, 
            # F: Inicio Lectiva, G: Inicio Productiva, H: Fin Productiva,
            # I: Técnico (Especialidad), J: Celular, K: Email, L: "Pendiente de seguimiento"
            fila_para_sheets = [
                fila[1],          # B: Tipo de Documento
                fila[2],          # C: Documento
                fila[3],          # D: Nombre
                fila[4],          # E: Ciudad
                fila[5],          # F: Fecha Inicio Lectiva
                fila[6],          # G: Fecha Inicio Productiva
                fila[7],          # H: Fecha Fin Productiva
            ]
            
            # Añadir Especialidad como campo técnico (columna I)
            if len(fila) > 11:
                fila_para_sheets.append(fila[11])  # Especialidad
            else:
                fila_para_sheets.append("")
                
            # Añadir Teléfono (columna J)
            if len(fila) > 8:
                fila_para_sheets.append(fila[8])
            else:
                fila_para_sheets.append("")
                
            # Añadir Correo (columna K)
            if len(fila) > 9:
                fila_para_sheets.append(fila[9])
            else:
                fila_para_sheets.append("")
                
            # Agregar "Pendiente de seguimiento" (columna L)
            fila_para_sheets.append("Pendiente de seguimiento")
                
            datos_para_sheets.append(fila_para_sheets)
        
        return datos_para_sheets
            
    except Exception as e:
        print(f"❌ Error al preparar datos para Google Sheets: {e}")
        return []

# Función modificada para pegar datos en Google Sheets
def pegar_datos_en_google_sheets(datos_formateados):
    try:
        # Abrir Google Sheets en una nueva pestaña
        driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1SK5dZ5IH3oGDwcvpxL-QFCsFQQL3ChZpoTS51pNlWGw/edit?usp=sharing', '_blank');")
        print("✅ Google Sheets abierto en nueva pestaña.")
        
        # Cambiar al contexto de la nueva pestaña
        driver.switch_to.window(driver.window_handles[-1])
        
        # Esperar a que la hoja de cálculo se cargue
        time.sleep(5)
        
        # Convertir datos a formato CSV para pegar
        csv_data = ""
        for fila in datos_formateados:
            csv_data += "\t".join([str(celda) for celda in fila]) + "\n"
        
        # Copiar los datos al portapapeles
        pyperclip.copy(csv_data)
        print("✅ Datos copiados al portapapeles.")
        
        # Instrucciones para el usuario
        print("\n⚠️ INSTRUCCIONES:")
        print("1. Haz clic en la celda B2 (primera celda de datos debajo del encabezado verde)")
        print("2. Presiona Ctrl+V para pegar los datos")
        print("3. Los datos deben aparecer en las columnas correctas")
        
        # Esperar a que el usuario copie los datos
        print("\n⚠️ Espera a que la hoja de Google Sheets cargue completamente...")
        time.sleep(2)
        
        # Intentar hacer clic en la celda B2
        try:
            # Intentar usar la API de Google Sheets para seleccionar B2
            try:
                driver.execute_script("SpreadsheetApp.getActiveSheet().getRange('B2').activate();")
            except:
                # Si falla, intentar hacer clic aproximado
                action = webdriver.ActionChains(driver)
                # Coordenadas ajustadas para B2 (debajo del encabezado)
                action.move_by_offset(150, 230).click().perform()
                print("✅ Se hizo clic en la ubicación aproximada de la celda B2.")
            
            # Simular Ctrl+V (pegar)
            try:
                action = webdriver.ActionChains(driver)
                action.key_down(webdriver.Keys.CONTROL).send_keys('v').key_up(webdriver.Keys.CONTROL).perform()
                print("✅ Se ha intentado pegar los datos automáticamente.")
            except:
                print("⚠️ No se pudo pegar automáticamente. Por favor, pega manualmente con Ctrl+V.")
                
        except Exception as e:
            print(f"⚠️ No se pudo hacer clic automáticamente en la celda B2: {e}")
            print("   Por favor, haz clic manualmente en la celda B2 y presiona Ctrl+V.")
        
    except Exception as e:
        print(f"❌ Error al pegar datos en Google Sheets: {e}")

# Ejecutar inicio de sesión
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
    # Obtener todas las filas de aprendices
    filas = driver.find_elements(By.XPATH, "//tbody[@id='filasTabla']/tr")
    total_aprendices = len(filas)
    print(f"✅ Se encontraron {total_aprendices} aprendices.")
    
    for i, fila in enumerate(filas):
        # Asegurarse de que no haya modales abiertos
        cerrar_modal_si_existe()
        
        # Volver a cargar las filas después de cerrar el modal
        if i > 0:  # Para el segundo aprendiz en adelante
            filas = driver.find_elements(By.XPATH, "//tbody[@id='filasTabla']/tr")
            fila = filas[i]
        
        # Extraer datos básicos de la tabla
        columnas = fila.find_elements(By.TAG_NAME, "td")
        datos_aprendiz = [columna.text for columna in columnas]
        
        try:
            # Hacer clic en el botón 'Detalle' utilizando JavaScript
            detalle_btn = fila.find_element(By.XPATH, ".//button[@id='solicitud_aprendiz']")
            driver.execute_script("arguments[0].click();", detalle_btn)
            print(f"✅ Clic en Detalle para aprendiz {i+1}.")
            time.sleep(2)
            
            # Extraer datos adicionales del modal
            try:
                # Extraer número de teléfono
                telefono = wait.until(EC.presence_of_element_located((By.XPATH, "//label[@id='lbl_apr_telefono']"))).text
                datos_aprendiz.append(telefono)
                
                # Extraer correo electrónico (usando el nuevo ID)
                try:
                    correo = driver.find_element(By.XPATH, "//label[@id='lbl_apr_email']").text
                    datos_aprendiz.append(correo)
                    print(f"✅ Correo capturado: {correo}")
                except Exception as e:
                    print(f"⚠️ No se pudo obtener el correo con el ID 'lbl_apr_email': {e}")
                    try:
                        # Intentar con el ID anterior como respaldo
                        correo = driver.find_element(By.XPATH, "//label[@id='lbl_apr_correo']").text
                        datos_aprendiz.append(correo)
                        print(f"✅ Correo capturado con ID alternativo: {correo}")
                    except:
                        datos_aprendiz.append("")
                        print("❌ No se pudo capturar el correo electrónico.")
                
                # Extraer programa de formación
                try:
                    programa = driver.find_element(By.XPATH, "//label[@id='lbl_apr_programa']").text
                    datos_aprendiz.append(programa)
                except:
                    datos_aprendiz.append("")
                    
                # Extraer especialidad
                try:
                    especialidad = driver.find_element(By.XPATH, "//label[@id='lbl_apr_especialidad']").text
                    datos_aprendiz.append(especialidad)
                    print(f"✅ Especialidad capturada: {especialidad}")
                except Exception as e:
                    print(f"⚠️ No se pudo obtener la especialidad: {e}")
                    datos_aprendiz.append("")
                
                # Agregar los datos completos a nuestra lista
                aprendices_data.append(datos_aprendiz)
                print(f"✅ Datos del aprendiz {datos_aprendiz[3]} copiados correctamente.")
                
            except Exception as e:
                print(f"⚠️ Error al extraer detalles del aprendiz: {e}")
                aprendices_data.append(datos_aprendiz)
            
            # Cerrar el modal de detalles explícitamente
            try:
                cerrar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='modalAprendiz']//button[@data-dismiss='modal']")))
                cerrar_btn.click()
                time.sleep(1.5)  # Esperar a que se cierre el modal
            except:
                print("⚠️ No se pudo cerrar el modal de manera normal. Intentando alternativas...")
                try:
                    # Intentar cerrar con JavaScript
                    driver.execute_script("$('#modalAprendiz').modal('hide');")
                    time.sleep(1.5)
                except:
                    # Si todo falla, refrescar la página y volver a empezar desde donde quedamos
                    if i < total_aprendices - 1:  # Si no es el último aprendiz
                        print("⚠️ Refrescando página para continuar con el siguiente aprendiz...")
                        driver.refresh()
                        time.sleep(3)
                        
                        # Volver a abrir la tabla de aprendices
                        ver_aprendices_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='solicitud_aprendices']")))
                        driver.execute_script("arguments[0].click();", ver_aprendices_btn)
                        time.sleep(2)
                
        except ElementClickInterceptedException:
            print(f"⚠️ El clic fue interceptado para el aprendiz {i+1}. Intentando cerrar modal...")
            cerrar_modal_si_existe()
            continue
        except Exception as e:
            print(f"⚠️ Error general al procesar aprendiz {i+1}: {e}")
            continue
    
    # Preparar y pegar datos en Google Sheets
    if aprendices_data:
        datos_formateados = formatear_datos_para_sheets(aprendices_data)
        if datos_formateados:
            pegar_datos_en_google_sheets(datos_formateados)
        else:
            print("⚠️ No se pudieron formatear los datos correctamente.")
    else:
        print("⚠️ No se recolectaron datos de aprendices para guardar.")
        
except Exception as e:
    print(f"❌ Error al extraer datos de aprendices: {e}")
finally:
    print("✅ Proceso completado. El navegador permanecerá abierto con la página de Google Sheets.")
    # El navegador no se cierra para permitir la copia manual de datos
    input("Presiona Enter cuando hayas terminado para cerrar el navegador...")
    driver.quit()