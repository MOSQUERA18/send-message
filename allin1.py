import pandas as pd
import time
from tkinter import Tk, Label, Button, filedialog, messagebox
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import quote
from webdriver_manager.chrome import ChromeDriverManager
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def cargar_datos(ruta_excel):
    try:
        datos = pd.read_excel(ruta_excel)
        return datos
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")
        return None

def obtener_saludo():
    hora_actual = datetime.now().hour
    if 5 <= hora_actual < 12:
        return "Buenos días"
    elif 12 <= hora_actual < 18:
        return "Buenas tardes"
    else:
        return "Buenas noches"

def generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo):
    saludo = obtener_saludo()
    documentos_lista = [doc.strip() for doc in str(documentos_faltantes).split(',')]
    documentos_formateados = "\n".join([f"{i+1}. {doc}" for i, doc in enumerate(documentos_lista)])
    mensaje = (f"{saludo} {nombre},\n\n"
               "Escribe Marlon Mosquera del Centro Agropecuario La Granja del Espinal. "
               f"Estamos validando los documentos de matrícula para la tecnología en {nombre_tecnologo}. "
               "Hemos identificado que es necesario enviar nuevamente los siguientes documentos:\n\n"
               f"{documentos_formateados}\n\n"
               "Por favor, realiza las correcciones indicadas y carga nuevamente todo el paquete completo.")
    return mensaje

def enviar_correo(correo, mensaje):
    try:
        remitente = ""
        clave = ""
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(remitente, clave)
        
        correo_msg = MIMEMultipart()
        correo_msg["From"] = remitente
        correo_msg["To"] = correo
        correo_msg["Subject"] = "Validación de Documentos"
        correo_msg.attach(MIMEText(mensaje, "plain"))
        
        servidor.sendmail(remitente, correo, correo_msg.as_string())
        servidor.quit()
        print(f"Correo enviado correctamente a {correo}")
    except Exception as e:
        print(f"Error al enviar el correo a {correo}: {e}")

def enviar_mensajes_whatsapp(ruta_excel):
    datos = cargar_datos(ruta_excel)
    if datos is None:
        return
    
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://web.whatsapp.com/")
    
    messagebox.showinfo("WhatsApp Web", "Inicie sesión en WhatsApp Web y presione Aceptar cuando los chats estén cargados.")
    
    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        documentos_faltantes = fila.get("Documentos_Faltantes")
        nombre_tecnologo = fila.get("Nombre_Tecnologo")
        numero_telefono = fila.get("Numero_Telefono")

        if pd.isna(nombre) or pd.isna(documentos_faltantes) or pd.isna(nombre_tecnologo) or pd.isna(numero_telefono):
            continue

        mensaje = generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo)
        mensaje_codificado = quote(mensaje)
        url = f"https://web.whatsapp.com/send?phone=+57{int(numero_telefono)}&text={mensaje_codificado}"
        driver.get(url)
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
            )
            send_button = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
            send_button.click()
            time.sleep(3)
        except Exception as e:
            print(f"Error al enviar mensaje a {numero_telefono}: {e}")
    
    driver.quit()
    messagebox.showinfo("Finalizado", "Mensajes enviados correctamente.")

def seleccionar_archivo(opcion):
    ruta_excel = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
    if ruta_excel:
        if opcion == 1:
            enviar_mensajes_whatsapp(ruta_excel)
        elif opcion == 2:
            procesar_correo(ruta_excel)

def procesar_correo(ruta_excel):
    datos = cargar_datos(ruta_excel)
    if datos is None:
        return
    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        documentos_faltantes = fila.get("Documentos_Faltantes")
        nombre_tecnologo = fila.get("Nombre_Tecnologo")
        correo = fila.get("Correo")
        
        if pd.isna(nombre) or pd.isna(documentos_faltantes) or pd.isna(nombre_tecnologo) or pd.isna(correo):
            continue
        
        mensaje = generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo)
        enviar_correo(correo, mensaje)
    
    messagebox.showinfo("Finalizado", "Correos enviados correctamente.")

def iniciar_interfaz():
    ventana = Tk()
    ventana.title("Opciones de Envío")
    ventana.geometry("400x250")
    
    Label(ventana, text="Seleccione una opción", font=("Arial", 16)).pack(pady=10)
    Button(ventana, text="Enviar mensajes a WhatsApp", command=lambda: seleccionar_archivo(1), width=30).pack(pady=5)
    Button(ventana, text="Enviar mensajes por Gmail", command=lambda: seleccionar_archivo(2), width=30).pack(pady=5)
    Button(ventana, text="Salir", command=ventana.quit, width=30).pack(pady=10)
    
    ventana.mainloop()

if __name__ == "__main__":
    iniciar_interfaz()
