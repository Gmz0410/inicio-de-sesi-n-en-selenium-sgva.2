from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import csv
from datetime import date

# Configuración del webdriver (asegúrate de tener el ChromeDriver en tu PATH)
driver = webdriver.Chrome()

# URL de la plataforma SVGA
url = "https://caprendizaje.sena.edu.co/sgva/SGVA_Diseno/pag/login.aspx"
driver.get(url)

# Esperar a que la página se cargue
wait = WebDriverWait(driver, 10)

# PASO 1: Iniciar sesión
username = "900729437"
password = "719D62B4"

# Completar los campos de usuario y contraseña
wait.until(EC.presence_of_element_located((By.ID, "txtUsuario"))).send_keys(username)
wait.until(EC.presence_of_element_located((By.ID, "txtContrasena"))).send_keys(password)

# Hacer clic en el botón de iniciar sesión
wait.until(EC.element_to_be_clickable((By.ID, "btnIngresar"))).click()

# PASO 2: Buscar pasantes y crear una nueva postulación
wait.until(EC.presence_of_element_located((By.ID, "campoBuscarPasantes"))).click()
wait.until(EC.element_to_be_clickable((By.ID, "btnCrearPostulacion"))).click()

# PASO 3: Diligenciar los datos requeridos
departamento = "Nombre del Departamento"
ciudad = "Nombre de la Ciudad"
especialidades = ["Especialidad 1", "Especialidad 2"] # Agrega las especialidades que necesites

wait.until(EC.presence_of_element_located((By.ID, "campoDepartamento"))).send_keys(departamento)
wait.until(EC.presence_of_element_located((By.ID, "campoCiudad"))).send_keys(ciudad)

# PASO 4: Buscar y seleccionar especialidades
for especialidad in especialidades:
    wait.until(EC.presence_of_element_located((By.ID, "campoBuscarEspecialidad"))).send_keys(especialidad)
    wait.until(EC.element_to_be_clickable((By.XPATH, f"//option[text()='{especialidad}']"))).click()

# PASO 5: Seleccionar competencias de la especialidad
competencia = "Nombre de la Competencia"
wait.until(EC.element_to_be_clickable((By.XPATH, f"//option[text()='{competencia}']"))).click()

# PASO 6: Diligenciar datos de la postulación
num_aspirantes = "3"
fecha_cierre = date.today().strftime("%Y-%m-%d")

wait.until(EC.presence_of_element_located((By.ID, "campoNumAspirantes"))).send_keys(num_aspirantes)
wait.until(EC.presence_of_element_located((By.ID, "campoFechaCierre"))).send_keys(fecha_cierre)
wait.until(EC.element_to_be_clickable((By.ID, "btnCrearPostulacion"))).click()

# PASO 7: Ir a postulaciones
wait.until(EC.element_to_be_clickable((By.ID, "btnIrAPostulaciones"))).click()

# PASO 8: Buscar aprendices y concursar
wait.until(EC.presence_of_element_located((By.ID, "campoBuscarAprendices"))).click()
wait.until(EC.element_to_be_clickable((By.ID, "btnVerAprendices"))).click()

# PASO 9: Copiar datos del aprendiz y guardarlos en una hoja de cálculo
aprendices = driver.find_elements(By.CLASS_NAME, "datosAprendiz")
with open('aprendices.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Nombre", "Correo", "Teléfono", "Especialidad"])
    for aprendiz in aprendices:
        nombre = aprendiz.find_element(By.CLASS_NAME, "nombre").text
        correo = aprendiz.find_element(By.CLASS_NAME, "correo").text
        telefono = aprendiz.find_element(By.CLASS_NAME, "telefono").text
        especialidad = aprendiz.find_element(By.CLASS_NAME, "especialidad").text
        writer.writerow([nombre, correo, telefono, especialidad])