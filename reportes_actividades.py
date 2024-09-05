import tkinter as tk
from tkinter import PhotoImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from collections import Counter
import random
import re
from tkinter import messagebox
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import sys


# Mapeo de meses en español a números
meses_map = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
}
mensajes_variados = [
    "Se resolvió el problema relacionado con:",
    "Se atendió la solicitud de:",
    "Se procesó la incidencia de:",
    "Se completó la tarea relacionada con:",
    "Se dio solución al incidente de:",
    "Se revisó el requerimiento de:",
    "Se ejecutó la operación para:",
    "Se gestionó la incidencia sobre:",
    "Se tomó acción sobre la incidencia de:",
    "Se realizó el seguimiento a:",
    "Se completó la gestión de:",
    "Se brindó asistencia a:",
    "Se cerró el caso relacionado con:",
    "Se monitoreó el incidente de:",
    "Se diagnosticó y resolvió el problema de:",
    "Se solucionó el inconveniente de:",
    "Se trabajó en la solicitud de:"
]

meses_map_inverso = {v: k.capitalize() for k, v in meses_map.items()}



def resource_path(relative_path):
    """ Retorna la ruta absoluta del recurso para PyInstaller """
    try:
        # PyInstaller guarda los archivos temporales en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def crear_pdf_desde_lista():
    global actividades_por_fecha, nombre, puesto, mes_nombre, ano,mes, funciones

    ano=int(ano)
    mes=int(mes)
    # Verifica si hay actividades para el año y mes seleccionados
    if ano in actividades_por_fecha and mes in actividades_por_fecha[ano]:
        nombre_pdf = f"reporte_actividades_{ano}_{mes}.pdf"
        crear_pdf(actividades_por_fecha[ano][mes], nombre, puesto, mes_nombre, ano, funciones, nombre_archivo=nombre_pdf)
        messagebox.showinfo("PDF Generado", f"El PDF ha sido generado con éxito: {nombre_pdf}")
    else:
        messagebox.showinfo("Sin actividades", "No se encontraron actividades para el mes y año seleccionados.")
def transformar_mes(mes):
    try:
        mes_num = int(mes)
        if 1 <= mes_num <= 12:
            return mes_num
        else:
            raise ValueError
    except ValueError:
        mes = mes.lower().strip()
        if mes in meses_map:
            return meses_map[mes]
        else:
            messagebox.showerror("Error", f"Mes '{mes}' no es válido. Introduce un mes válido.")
            return None

def limpiar_asunto(asunto):
    # Remover los prefijos comunes de correos electrónicos como "Fwd:", "Re:", etc.
    asunto_limpio = re.sub(r'^(Fwd:|Re:|Fw:|RE:)\s*', '', asunto).strip()
    return asunto_limpio

# Función para crear el PDF estilizado
def crear_pdf(datos_actividades, nombre, puesto, mes_nombre, ano, funciones, nombre_archivo="reporte_actividades.pdf"):
    doc = SimpleDocTemplate(nombre_archivo, pagesize=A4)
    styles = getSampleStyleSheet()
    
    # Definir estilos básicos
    normal_style = styles['Normal']
    normal_style.fontSize = 12
    title_style = styles['Title']
    title_style.fontSize = 18
    heading_style = styles['Heading2']
    heading_style.fontSize = 14

    story = []

    # Título del reporte
    titulo = Paragraph("Reporte de actividades", title_style)
    story.append(titulo)
    story.append(Spacer(1, 0.3 * inch))
    
    # Encabezado con nombre, puesto, mes y año
    encabezado = Paragraph(f"{nombre}<br/>{puesto}<br/>{mes_nombre} - {ano}", normal_style)
    story.append(encabezado)
    story.append(Spacer(1, 0.3 * inch))
    
    # Línea de separación
    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.grey))
    story.append(Spacer(1, 0.3 * inch))

    # Sección de funciones
    story.append(Paragraph("Funciones:", heading_style))
    if funciones:
        for funcion in funciones:
            story.append(Paragraph(f"● {funcion.strip()}", normal_style))
    else:
        story.append(Paragraph("No se especificaron funciones.", normal_style))
    story.append(Spacer(1, 0.3 * inch))
    
    # Línea de separación
    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.grey))
    story.append(Spacer(1, 0.3 * inch))
    
    # Sección de actividades
    story.append(Paragraph("Actividades del mes:", heading_style))

    if datos_actividades:
        # Limpiar y contar los asuntos
        asuntos_limpios = [limpiar_asunto(actividad['asunto']) for actividad in datos_actividades]
        contador_asuntos = Counter(asuntos_limpios)

        for asunto, count in contador_asuntos.items():
            # Seleccionar un mensaje aleatorio de la matriz
            mensaje_aleatorio = random.choice(mensajes_variados)

            if count > 1:
                story.append(Paragraph(f"● {mensaje_aleatorio} '{asunto}' N {count} veces", normal_style))
            else:
                story.append(Paragraph(f"● {mensaje_aleatorio} '{asunto}'", normal_style))
    else:
        story.append(Paragraph("No se encontraron actividades para el mes seleccionado.", normal_style))
    
    story.append(Spacer(1, 0.5 * inch))
    
    # Línea de firma
    firma = Paragraph("<br/><br/>__________________________<br/>Firma", normal_style)
    story.append(firma)
    story.append(Spacer(1, 0.5 * inch))
    
    # Pie de página con el nombre del desarrollador y fecha
    footer = Paragraph(f"{nombre}<br/>{puesto}<br/>{mes_nombre} - {ano}", normal_style)
    story.append(footer)

    # Construir el documento PDF
    doc.build(story)
    messagebox.showinfo("PDF Generado", f"El PDF ha sido generado con éxito: {nombre_archivo}")

def rellenar_campos():
    # Rellenar los campos con los datos proporcionados
    entry_nombre.delete(0, tk.END)
    entry_nombre.insert(0, "Christofer Giovanny Luevanos Torres")
    
    entry_puesto.delete(0, tk.END)
    entry_puesto.insert(0, "Desarrollador")
    
    entry_usuario.delete(0, tk.END)
    entry_usuario.insert(0, "cluevanos")
    
    entry_contrasena.delete(0, tk.END)
    entry_contrasena.insert(0, "Gatita5682.")
    
    # Rellenar las funciones
    funciones = """Revisión de sitios institucionales.
Mantenimiento de soluciones tecnológicas.
Creación de soluciones tecnológicas.
Maquetado de sitios web.
Actualizaciones del sitio institucional.
Atención directa a alumnos, personal docente y administrativo.
Atención en mesa de ayuda."""
    
    text_funciones.delete("1.0", tk.END)
    text_funciones.insert(tk.END, funciones)

def ejecutar_bot():
    global actividades_por_fecha, nombre, puesto, mes_nombre, ano,mes, funciones
    nombre = entry_nombre.get()
    puesto = entry_puesto.get()
    usuario = entry_usuario.get()
    contrasena = entry_contrasena.get()
    ano = entry_ano.get()
    mes = transformar_mes(entry_mes.get())
    funciones = text_funciones.get("1.0", tk.END).strip().split("\n")
    if not nombre or not puesto or not usuario or not contrasena or not ano or not mes:
        messagebox.showwarning("Advertencia", "Todos los campos son obligatorios.")
        return
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    webdriver_service = Service(resource_path('chromedriver.exe'))
    driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)
    try:
        url = 'https://itop3.ceti.mx'
        driver.get(url)
        time.sleep(3)
        username_input = driver.find_element(By.ID, 'user')
        username_input.send_keys(usuario)

        password_input = driver.find_element(By.ID, 'pwd')
        password_input.send_keys(contrasena)
        password_input.send_keys(Keys.RETURN)
        time.sleep(6)
        # Extraer la propiedad 'data-tooltip-content'
        img_element = driver.find_element(By.CSS_SELECTOR, 'img.ibo-navigation-menu--user-picture--image')
        tooltip_content = img_element.get_attribute('data-tooltip-content')
        
        # Eliminar "Conectado como " y guardar el resto en 'user_logged'
        user_logged = tooltip_content.replace("Conectado como ", "").replace("(Administrator)", "").strip()
        
        print(f"Usuario logueado: {user_logged}")
        
        driver.get("https://itop3.ceti.mx/pages/UI.php?operation=search&filter=%255B%2522SELECT+%2560UserRequest%2560+FROM+UserRequest+AS+%2560UserRequest%2560+WHERE+1%2522%252C%255B%255D%252C%255B%255D%255D&c[menu]=SearchUserRequests&c[org_id]=")
        time.sleep(3)
        # Aumentar el tiempo de espera para el botón "Agregar nuevo criterio"
        agregar_criterio_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.sfm_toggler'))
        )
        # Desplazar el botón "Agregar nuevo criterio" a la vista
        driver.execute_script("arguments[0].scrollIntoView(true);", agregar_criterio_btn)
        time.sleep(2)
        agregar_criterio_btn.click()
        # Aumentar el tiempo de espera para el checkbox de "Analista"
        analista_checkbox = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'li[data-field-ref="UserRequest.agent_id"]'))
        )
        # Verificar si el texto del elemento contiene "Analista"
        if "Analista" in analista_checkbox.text:
            analista_checkbox.click()
            print("Checkbox de 'Analista' seleccionado.")
        else:
            print("El elemento no contiene el texto 'Analista'. No se seleccionó el checkbox.")
        time.sleep(10)
        
        search_wrapper = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'span.sff_input_wrapper'))
        )
        # Luego, buscar el input dentro del contenedor
        filtro_input = search_wrapper.find_element(By.CSS_SELECTOR, 'input[placeholder="Búsqueda..."]')
        # Desplazar el campo de búsqueda a la vista
        driver.execute_script("arguments[0].scrollIntoView(true);", filtro_input)
        time.sleep(2)
        # Verificar si está visible e interactuar con Selenium ActionChains
        if filtro_input.is_displayed():
            filtro_input.click()
            
            # Usar ActionChains para simular la escritura de cada carácter
            actions = ActionChains(driver)
            actions.move_to_element(filtro_input)
            for character in user_logged:
                actions.send_keys(character)
                time.sleep(0.1)  # Dar un pequeño tiempo entre cada tecla
            actions.send_keys(Keys.RETURN)
            actions.perform()
        else:
            print("El campo de búsqueda no está visible. Haciendo clic con JavaScript.")
            driver.execute_script("arguments[0].click();", filtro_input)
            
            # Usar ActionChains para escribir de nuevo si no funcionó el clic normal
            actions = ActionChains(driver)
            for character in user_logged:
                actions.send_keys(character)
                time.sleep(0.1)
            
            actions.send_keys(Keys.RETURN)
            actions.perform()
        time.sleep(2)
        wait = WebDriverWait(driver, 10)
        # Aumentar el tiempo de espera para el checkbox del valor correspondiente
        valor_checkbox = wait.until(EC.presence_of_element_located((By.XPATH, f'//label[contains(text(), "{user_logged}")]')))
        valor_checkbox.click()
        
        # Espera a que el elemento esté presente y completamente cargado
        select_element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.NAME, 'datatable_result_1_length'))
        )
        select = Select(select_element)
        select.select_by_value('-1')  # Seleccionar "All"
        
        # Espera a que el elemento esté presente y completamente cargado
        select_element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.NAME, 'datatable_table_inner_id_sf_search_1_length'))
        )
        
        # Asegurarse de que el select esté visible en pantalla antes de interactuar
        driver.execute_script("arguments[0].scrollIntoView(true);", select_element)
        time.sleep(2)  # Breve pausa para asegurarse de que el elemento esté visible

        # Seleccionar la opción "All" con Selenium
        try:
            select = Select(select_element)
            select.select_by_value('-1')  # Seleccionar "All"
            print("Opción 'All' seleccionada correctamente.")
        except Exception as e:
            print(f"Error al seleccionar la opción 'All' con Selenium: {e}")
            
            # Forzar la selección con JavaScript si Selenium falla
            driver.execute_script("arguments[0].value = '-1';", select_element)
            print("Opción 'All' seleccionada con JavaScript.")

        time.sleep(5)  # Espera para que la página cargue los nuevos datos

        # Inicializar diccionario para almacenar actividades por año y mes
        actividades_por_fecha = {}

        # Extraer las filas de la tabla
        rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
        
        rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
        actividades_por_fecha = {}

        for row in rows:
            try:
                # Extraer fecha
                fecha_td = row.find_element(By.CSS_SELECTOR, 'td[data-attribute-code="start_date"]')
                fecha_text = fecha_td.get_attribute('data-value-raw')
                #print(f"Fecha extraída: {fecha_text}")  # Para depuración
                
                fecha = datetime.strptime(fecha_text, "%Y-%m-%d %H:%M:%S")
                
                # Extraer enlace
                enlace = row.find_element(By.CSS_SELECTOR, 'a.object-ref-link').get_attribute('href')
                #print(f"Enlace extraído: {enlace}")  # Para depuración
                
                # Extraer asunto
                asunto_td = row.find_element(By.CSS_SELECTOR, 'td[data-attribute-code="title"]')
                asunto = asunto_td.get_attribute('data-value-raw')
                #print(f"Asunto extraído: {asunto}")  # Para depuración
                
                # Organizar las actividades por año y mes
                year = fecha.year
                month = fecha.month
                #print(f"Año: {year}, Mes: {month}")  # Para verificar

                if year not in actividades_por_fecha:
                    actividades_por_fecha[year] = {}
                
                if month not in actividades_por_fecha[year]:
                    actividades_por_fecha[year][month] = []

                actividades_por_fecha[year][month].append({'asunto': asunto, 'enlace': enlace, 'fecha': fecha_text})
            
            except Exception as e:
                print(f"Ocurrió un error en la fila: {e}")

        # Verificar las actividades por año y mes
        #print(f"Actividades organizadas por fecha: {actividades_por_fecha}")  # Para depuración

        # Generar el PDF con los datos extraídos

        # Pasa las actividades del año y mes seleccionados al PDF
        mes_nombre = meses_map_inverso[mes]
        print(f"Generacion del PDf para el año {ano}, mes {mes} disponible.")  # Depuración
        boton_crear_pdf.config(state=tk.NORMAL)
        messagebox.showinfo("Operación Completa", "La lista de actividades ha sido generada.")
    finally:
        driver.quit()

# Nueva función para iniciar sesión en PD (Proyectos Desarrollo)
def ejecutar_bot_pd():
    usuario = entry_usuario.get()
    contrasena = entry_contrasena.get()

    if not usuario or not contrasena:
        messagebox.showwarning("Advertencia", "Usuario y contraseña son obligatorios.")
        return

    chrome_options = Options()
    #chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    
    webdriver_service = Service('./chromedriver.exe')
    driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)

    try:
        # Navegar al sitio web de PD
        url = 'https://proyectosdesarrollo.ceti.mx/login?back_url=https%3A%2F%2Fproyectosdesarrollo.ceti.mx%2Flogin'
        driver.get(url)
        time.sleep(3)
        
        # Iniciar sesión
        username_input = driver.find_element(By.ID, 'username')
        username_input.send_keys(usuario)

        password_input = driver.find_element(By.ID, 'password')
        password_input.send_keys(contrasena)
        password_input.send_keys(Keys.RETURN)
        time.sleep(6)

        # Esperar y presionar el botón "Módulos"
        modulos_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[title="Módulos"]'))
        )
        modulos_button.click()

        # Esperar a que el menú se despliegue y presionar "Paquetes de trabajo"
        paquetes_trabajo_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.work-packages-menu-item.op-menu--item-action[title="Paquetes de trabajo"]'))
        )
        paquetes_trabajo_link.click()

        time.sleep(2)

        # Esperar a que el botón de "Activar Filtro" esté disponible y hacer clic
        filtro_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-qa-selector="wp-filter-button"]'))
        )
        filtro_button.click()
        time.sleep(2)
        # Esperar a que el botón "Borrar" esté disponible y hacer clic
        borrar_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.spot-link[title="Borrar"]'))
        )
        borrar_button.click()

        # Seleccionamos el input dentro del div con la clase 'ng-input'
        combobox_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.ng-input input[role="combobox"][type="text"]'))
        )

        # Desplazar hasta el elemento si no está visible
        driver.execute_script("arguments[0].scrollIntoView(true);", combobox_input)

        # Intentar hacer clic usando JavaScript en lugar de .click() directo
        driver.execute_script("arguments[0].click();", combobox_input)


        # Esperar a que la opción "Fecha de inicio" esté disponible y hacer clic en ella
        fecha_inicio_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.ng-option[role="option"] span.ng-option-label'))
        )

        # Verificar que el texto sea "Fecha de inicio" y hacer clic
        if fecha_inicio_option.text == "Fecha de inicio":
            fecha_inicio_option.click()

        messagebox.showinfo("Operación Completa", "Se ha iniciado sesión y accedido a 'Paquetes de trabajo' en Proyectos de Desarrollo.")
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un error: {e}")
    finally:
        driver.quit()

root = tk.Tk()
root.title("Buggie-Generador de reportes")
icon = PhotoImage(file="buggie.png")
root.iconphoto(True, icon)
root.geometry("420x530")
root.configure(bg="#2c3e50")

style = ttk.Style()
style.configure("TLabel", font=("Helvetica", 12), background="#2c3e50", foreground="white")
style.configure("TEntry", font=("Helvetica", 12))
style.configure("TButton", font=("Helvetica", 12, "bold"), background="#1abc9c", foreground="white")
style.map("TButton", background=[('active', '#16a085')])

# Creación de marco principal con padding
frame = ttk.Frame(root, padding="20 20 20 20", style="TFrame")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Cabecera
header_label = tk.Label(frame, text="Generador de reportes v1", font=("Helvetica", 16, "bold"), bg="#34495e", fg="white")
header_label.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky="ew")

# Campos de entrada
ttk.Label(frame, text="Nombre:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
entry_nombre = ttk.Entry(frame, width=30)
entry_nombre.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(frame, text="Puesto:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
entry_puesto = ttk.Entry(frame, width=30)
entry_puesto.grid(row=2, column=1, padx=5, pady=5)

ttk.Label(frame, text="Usuario:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
entry_usuario = ttk.Entry(frame, width=30)
entry_usuario.grid(row=3, column=1, padx=5, pady=5)

ttk.Label(frame, text="Contraseña:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
entry_contrasena = ttk.Entry(frame, show="*", width=30)
entry_contrasena.grid(row=4, column=1, padx=5, pady=5)

ttk.Label(frame, text="Año:").grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)

# Creación de lista de años desde 2023 hasta 2100
years = [str(year) for year in range(2023, 2101)]
entry_ano = ttk.Combobox(frame, values=years, width=13, state="readonly")
entry_ano.grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)
entry_ano.current(0)  # Establece el valor por defecto en 2023

ttk.Label(frame, text="Mes:").grid(row=6, column=0, padx=5, pady=5, sticky=tk.W)

# Creación de lista de meses en español con la primera letra en mayúscula
months = [mes.capitalize() for mes in meses_map.keys()]
entry_mes = ttk.Combobox(frame, values=months, width=15, state="readonly")
entry_mes.grid(row=6, column=1, padx=5, pady=5, sticky=tk.W)
entry_mes.current(0)  # Establece el valor por defecto en Enero

# Text area para funciones
ttk.Label(frame, text="Funciones:").grid(row=7, column=0, padx=5, pady=5, sticky=tk.N)
text_funciones = tk.Text(frame, height=6, width=30)
text_funciones.grid(row=7, column=1, padx=5, pady=5, sticky=tk.W)

# Botón para rellenar automáticamente
boton_rellenar = ttk.Button(frame, text="Rellenar Automáticamente", command=rellenar_campos)
boton_rellenar.grid(row=8, column=0, columnspan=2, pady=10, sticky="ew")

# Cambia el texto del botón de ejecución por "Ingresar iTop"
boton_ejecutar = ttk.Button(frame, text="Ingresar iTop", command=ejecutar_bot)
boton_ejecutar.grid(row=9, column=0, columnspan=2, pady=20, sticky="ew")

# Nuevo botón "Ingresar PD"
boton_ingresar_pd = ttk.Button(frame, text="Ingresar PD", command=ejecutar_bot_pd)
boton_ingresar_pd.grid(row=10, column=0, columnspan=2, pady=10, sticky="ew")

boton_crear_pdf = ttk.Button(frame, text="Crear PDF", command=crear_pdf_desde_lista)
boton_crear_pdf.grid(row=11, column=0, columnspan=2, pady=10, sticky="ew")

# Deshabilitar el botón al principio
boton_crear_pdf.config(state=tk.DISABLED)

# Configuración de padding en los widgets del frame
for widget in frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

root.mainloop()
