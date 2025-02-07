import pandas as pd
import time
import tkinter as tk
from tkinter import ttk
import threading
from tkinter import Tk, Label, Button, filedialog, messagebox, Toplevel,PhotoImage
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
import os
import tkinter.font as tkFont



def mostrar_vista_previa(df, ruta_excel, opcion, ventana):
    """ Muestra una vista previa del archivo Excel antes de enviarlo. """
    
    # Crear ventana emergente
    vista_previa = tk.Toplevel(ventana)
    vista_previa.title("Vista previa del archivo")
    vista_previa.geometry("600x400")

    # Etiqueta de instrucciones
    label = tk.Label(vista_previa, text="Verifica los datos antes de enviarlos:", font=("Arial", 12, "bold"))
    label.pack(pady=5)

    # Crear un frame con scroll para mostrar la tabla
    frame_tabla = tk.Frame(vista_previa)
    frame_tabla.pack(fill=tk.BOTH, expand=True)

    # Crear tabla con ttk.Treeview
    tree = ttk.Treeview(frame_tabla, columns=list(df.columns), show="headings")

    # Agregar encabezados de columnas
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)

    # Agregar filas de la tabla (solo las primeras 10 filas para no saturar la ventana)
    for _, row in df.head(10).iterrows():
        tree.insert("", tk.END, values=list(row))

    tree.pack(fill=tk.BOTH, expand=True)

    # Botones para aceptar o cancelar el envío
    def aceptar_envio():
        vista_previa.destroy()
        if opcion == 1:
            threading.Thread(target=enviar_mensajes_whatsapp, args=(ruta_excel, ventana)).start()
        elif opcion == 2:
            threading.Thread(target=procesar_correo, args=(ruta_excel, ventana)).start()

    tk.Button(vista_previa, text="Aceptar y Enviar", command=aceptar_envio, bg="#2ecc71", fg="white").pack(side=tk.LEFT, padx=20, pady=10)
    tk.Button(vista_previa, text="Cancelar", command=vista_previa.destroy, bg="#e74c3c", fg="white").pack(side=tk.RIGHT, padx=20, pady=10)



def obtener_ruta_descargas():
    """ Obtiene la ruta de la carpeta de Descargas del usuario """
    if os.name == "nt":  # Windows
        return os.path.join(os.environ["USERPROFILE"], "Downloads")
    else:  # MacOS y Linux
        return os.path.join(os.path.expanduser("~"), "Downloads")

def descargar_plantilla(tipo):
    """ Descarga una plantilla de Excel en la carpeta de Descargas del usuario """
    ruta_descargas = obtener_ruta_descargas()
    os.makedirs(ruta_descargas, exist_ok=True)  # Asegurar que la carpeta existe

    if tipo == "whatsapp":
        ruta_guardado = os.path.join(ruta_descargas, "plantilla_whatsapp.xlsx")
        columnas = ["Nombre", "Numero_Telefono", "Remitente", "Mensaje"]
    else:
        ruta_guardado = os.path.join(ruta_descargas, "plantilla_correo.xlsx")
        columnas = ["Nombre", "Correo", "Remitente", "Mensaje"]

    df = pd.DataFrame(columns=columnas)
    df.to_excel(ruta_guardado, index=False)

    from tkinter import messagebox
    messagebox.showinfo("Éxito", f"Plantilla de {tipo.capitalize()} guardada en: {ruta_guardado}")



def mostrar_cargando(ventana):
    """ Muestra una ventana emergente de carga """
    global ventana_carga
    ventana_carga = Toplevel(ventana)
    ventana_carga.title("Cargando...")
    ventana_carga.geometry("250x100")
    ventana_carga.resizable(False, False)
    
    Label(ventana_carga, text="Procesando, por favor espere...", font=("Arial", 10)).pack(pady=10)
    
    # Deshabilitar la ventana principal mientras se carga
    ventana_carga.grab_set()
    ventana.update()

def ocultar_cargando():
    """ Oculta la ventana emergente de carga """
    ventana_carga.destroy()

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

def generar_mensaje(nombre, remitente, mensaje_base):
    """ Genera el mensaje con negritas y emojis para WhatsApp. """
    saludo = obtener_saludo()  # Obtiene "Buenos días", "Buenas tardes" o "Buenas noches"

    # Si el mensaje base está vacío, se usa un mensaje predeterminado
    if pd.isna(mensaje_base) or mensaje_base.strip() == "":
        mensaje_base = "📌 *Por favor, revisa esta información importante.*"

    # Formato del mensaje mejorado con negritas y emojis
    mensaje = (f"👋 *Hola {nombre}*, {saludo}.\n\n"
               f"✍️ *Escribe {remitente}* desde el *Centro Agropecuario La Granja.*\n\n"
               f"📢 {mensaje_base}\n\n"
               "Gracias por tu atención. ✅")

    return mensaje


def enviar_mensajes_whatsapp(ruta_excel, ventana):
    """ Envía mensajes a WhatsApp usando los datos del Excel. """
    mostrar_cargando(ventana)

    datos = cargar_datos(ruta_excel)
    if datos is None:
        ocultar_cargando()
        return

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://web.whatsapp.com/")

    messagebox.showinfo("WhatsApp Web", "Inicie sesión en WhatsApp Web y presione Aceptar cuando los chats estén cargados.")

    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        numero_telefono = fila.get("Numero_Telefono")
        remitente = fila.get("Remitente")  # Nombre de quien envía el mensaje
        mensaje_base = fila.get("Mensaje")  # Mensaje personalizado

        if pd.isna(nombre) or pd.isna(numero_telefono) or pd.isna(remitente):
            continue

        mensaje = generar_mensaje(nombre, remitente, mensaje_base)
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
    ocultar_cargando()
    messagebox.showinfo("Finalizado", "Mensajes enviados correctamente.")


def validar_archivo_excel(ruta_excel):
    """ Verifica que el archivo tenga el formato correcto: 
    - Solo archivos .xlsx
    - No debe estar vacío
    - Debe contener exactamente las columnas esperadas
    """

    try:
        # Cargar el archivo
        df = pd.read_excel(ruta_excel)

        # Verificar que el archivo no esté vacío
        if df.empty:
            messagebox.showerror("Error", "⚠️ El archivo Excel está vacío. Carga un archivo válido.")
            return False

        # Definir las columnas esperadas según el tipo de mensaje
        columnas_esperadas = {"Nombre", "Numero_Telefono", "Remitente", "Mensaje"}

        # Obtener las columnas reales del archivo
        columnas_actuales = set(df.columns)

        # Verificar si faltan columnas
        columnas_faltantes = columnas_esperadas - columnas_actuales
        if columnas_faltantes:
            messagebox.showerror("Error", f"⚠️ Faltan las siguientes columnas en el archivo: {', '.join(columnas_faltantes)}")
            return False

        # Verificar si hay columnas adicionales no permitidas
        columnas_extras = columnas_actuales - columnas_esperadas
        if columnas_extras:
            messagebox.showwarning("Advertencia", f"⚠️ El archivo contiene columnas adicionales: {', '.join(columnas_extras)}.\n"
                                                 "Solo se procesarán las columnas esperadas.")

        return True  # Si todo está correcto, retorna True

    except Exception as e:
        messagebox.showerror("Error", f"⚠️ No se pudo leer el archivo Excel. Verifica que sea un archivo válido.\n\nError: {e}")
        return False


def seleccionar_archivo(opcion, ventana):
    """ Permite seleccionar solo archivos Excel (.xlsx) y valida su contenido. """
    ruta_excel = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )

    if not ruta_excel:
        return  # Si el usuario cancela la selección, no hacer nada

    # Verificar si el archivo es válido
    if not validar_archivo_excel(ruta_excel):
        return  # Si hay un error, mostrar alerta y salir

    # Cargar el archivo para mostrar vista previa
    df = pd.read_excel(ruta_excel)
    mostrar_vista_previa(df, ruta_excel, opcion, ventana)  # Mostrar la vista previa antes de enviar


def procesar_correo(ruta_excel, ventana):
    """ Procesa y envía correos electrónicos a partir del archivo Excel. """
    mostrar_cargando(ventana)

    datos = cargar_datos(ruta_excel)
    if datos is None:
        ocultar_cargando()
        return

    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        correo = fila.get("Correo")
        remitente = fila.get("Remitente")  # Nombre de quien envía el mensaje
        mensaje_base = fila.get("Mensaje")  # Mensaje personalizado

        if pd.isna(nombre) or pd.isna(correo) or pd.isna(remitente):
            continue

        mensaje = generar_mensaje(nombre, remitente, mensaje_base)
        enviar_correo(correo, mensaje)

    ocultar_cargando()
    messagebox.showinfo("Finalizado", "Correos enviados correctamente.")


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
        correo_msg["Subject"] = "Información Importante"
        correo_msg.attach(MIMEText(mensaje, "plain"))
        
        servidor.sendmail(remitente, correo, correo_msg.as_string())
        servidor.quit()
        print(f"Correo enviado correctamente a {correo}")
    except Exception as e:
        print(f"Error al enviar el correo a {correo}: {e}")



def obtener_tiempo():
    hora = time.strftime('%H:%M:%S')
    fecha = time.strftime('%A %d %B %Y')
    texto_hora.config(text=hora)
    texto_fecha.config(text=fecha)
    texto_hora.after(1000, obtener_tiempo)

def agregar_reloj(ventana):
    global texto_hora, texto_fecha
    reloj_frame = Toplevel(ventana)
    reloj_frame.geometry("1300x100")  # Ajusta la posición en la esquina superior derecha
    reloj_frame.overrideredirect(True)  # Eliminar barra de título
    reloj_frame.config(bg='gray')
    reloj_frame.wm_attributes('-transparentcolor', 'gray')
    
    texto_hora = Label(reloj_frame, fg='white', bg='black', font=('Arial', 20, 'bold'))
    texto_hora.pack()
    texto_fecha = Label(reloj_frame, fg='white', bg='black', font=('Arial', 12))
    texto_fecha.pack()
    
    obtener_tiempo()


def iniciar_interfaz():
    global ventana
    
    ventana = Tk()
    
    # Colores suaves para el fondo
    ventana.title("Envío de Mensajes Masivos a WhatsApp y Gmail")
    ventana.geometry("500x400")
    ventana.resizable(False, False)
    ventana.config(bg='#f4f4f9')  # Fondo de color suave (blanco sucio)

    # Establecer fuente personalizada
    font_button = tkFont.Font(family="Segoe UI", size=12, weight="bold")
    font_label = tkFont.Font(family="Segoe UI", size=16, weight="bold")

    # Cambiar íconos
    ventana.iconbitmap("C:\\send-message\\backend\\iconos\\send.ico") 

    # Cargar iconos y ajustar tamaño
    icono_whatsapp = PhotoImage(file="C:\\send-message\\backend\\iconos\\wasap.png").subsample(8, 8)
    icono_gmail = PhotoImage(file="C:\\send-message\\backend\\iconos\\gmaili.png").subsample(8, 8)
    icono_exit = PhotoImage(file="C:\\send-message\\backend\\iconos\\exit.png").subsample(11, 11)
    
    agregar_reloj(ventana)  # Agregar el reloj en la esquina superior derecha

    # Label con color de fondo claro
    Label(ventana, text="Seleccione una opción", font=font_label, bg='#f4f4f9').pack(pady=10)
    
    # Botones con colores suaves, bordes redondeados y sombras
    Button(ventana, text="Enviar mensajes a WhatsApp", command=lambda: seleccionar_archivo(1, ventana), 
           width=30, height=2, font=font_button, relief="flat", bg='#3498db', fg='white', 
           activebackground='#2980b9').pack(pady=5)
    Button(ventana, text="Enviar mensajes por Gmail", command=lambda: seleccionar_archivo(2, ventana), 
           width=30, height=2, font=font_button, relief="flat", bg='#e74c3c', fg='white', 
           activebackground='#c0392b').pack(pady=5)

    Button(ventana, text="  ""Descargar plantilla WhatsApp", image=icono_whatsapp, compound="left", 
           command=lambda: descargar_plantilla("whatsapp"), width=300, height=50, font=font_button, 
           relief="flat", bg='#2ecc71', fg='white', activebackground='#27ae60').pack(pady=5)
    
    Button(ventana, text="  ""Descargar plantilla Gmail", image=icono_gmail, compound="left", 
           command=lambda: descargar_plantilla("gmail"), width=300, height=50, font=font_button, 
           relief="flat", bg='#f39c12', fg='white', activebackground='#e67e22').pack(pady=5)
    
    Button(ventana, text="Salir", image=icono_exit, compound="left", 
        command=ventana.quit, width=285, height=50, font=font_button, 
        relief="flat", bg='#7f8c8d', fg='white', activebackground='#95a5a6',
        padx=10).pack(pady=2)


    ventana.mainloop()
if __name__ == "__main__":
    iniciar_interfaz()