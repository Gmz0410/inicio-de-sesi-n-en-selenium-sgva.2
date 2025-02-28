from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Configurar el WebDriver (Asegúrate de tener el driver correcto instalado)
driver = webdriver.Chrome()  # Usa el WebDriver según tu navegador

# URL de inicio de sesión
login_url = "https://caprendizaje.sena.edu.co/sgva/SGVA_Diseno/pag/login.aspx"

# Abrir el navegador e ir al sitio web
driver.get(login_url)
time.sleep(3)  # Esperar a que la página cargue completamente

# Buscar y completar el campo NIT (Usuario)
try:
    nit_input = driver.find_element(By.XPATH, "//input[@type='text']")
    nit_input.send_keys("900729437")  
    print("✅ Campo NIT encontrado y llenado.")
except Exception as e:
    print("❌ Error: No se encontró el campo NIT.", e)

# Buscar y completar el campo Contraseña
try:
    password_input = driver.find_element(By.XPATH, "//input[@type='password']")
    password_input.send_keys("719D62B4")  
    print("✅ Campo Contraseña encontrado y llenado.")
except Exception as e:
    print("❌ Error: No se encontró el campo Contraseña.", e)

# Buscar y hacer clic en el botón "Iniciar sesión"
try:
    login_button = driver.find_element(By.XPATH, "//input[@type='submit']")
    login_button.click()
    print("✅ Botón 'Iniciar sesión' encontrado y clickeado.")
except Exception as e:
    print("❌ Error: No se encontró el botón 'Iniciar sesión'.", e)

# Esperar unos segundos para cargar la página después del login
time.sleep(5)

# Ir a la página de solicitudes
solicitudes_url = "https://caprendizaje.sena.edu.co/sgva/Empresa/Solicitudes/"
driver.get(solicitudes_url)
print("✅ Página de solicitudes abierta.")

# Esperar a que la página cargue
time.sleep(3)

# Buscar y hacer clic en el botón "Crear nueva solicitud"
try:
    crear_solicitud_button = driver.find_element(By.LINK_TEXT, "Crear nueva solicitud")
    crear_solicitud_button.click()
    print("✅ Botón 'Crear nueva solicitud' encontrado y clickeado.")
except Exception as e:
    print("❌ Error: No se encontró el botón 'Crear nueva solicitud'.", e)

# Esperar unos segundos para que la página de creación cargue
time.sleep(5)

# Cerrar el navegador (opcional, puedes quitar esto si quieres continuar con más pasos)
driver.quit()
