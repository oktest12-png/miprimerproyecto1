# Importamos las librerías necesarias
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from collections import defaultdict
import random
import json

# Importamos las librerías para el código QR
import qrcode
from PIL import Image, ImageTk

### NUEVO: Importamos librerías para el reporte HTML ###
import glob
import webbrowser
from matplotlib import pyplot as plt
import locale

# --- Variables Globales ---
ruta_archivo_log = "" 
ultima_posicion = 0
conteo_por_hora = defaultdict(int) 
conteo_total_dia = 0
fecha_actual = datetime.now().date()
qr_image_ref = None 

# --- INICIA NUEVA SECCIÓN DE REPORTE HTML ---

def procesar_datos_reporte(lista_archivos_json):
    """
    Lee una lista de archivos JSON, los procesa y calcula estadísticas.
    Aplica la regla de la jornada laboral de 6:00 a 14:59.
    """
    datos_procesados = {
        "resumen": {},
        "dias": []
    }
    horas_jornada = range(6, 15) # De 6:00 a 14:59
    
    # Intentamos configurar el idioma a español para los nombres de los días
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
        except locale.Error:
            print("Advertencia: No se pudo configurar el local a español. Los días aparecerán en inglés.")

    for archivo in lista_archivos_json:
        with open(archivo, 'r') as f:
            data = json.load(f)
        
        fecha_obj = datetime.strptime(data['fecha'], '%Y-%m-%d').date()
        nombre_dia = fecha_obj.strftime('%A').capitalize()
        
        conteo_horas = {}
        total_dia = 0
        desglose_json = {int(k): v for k, v in data['desglose_por_hora'].items()}

        for hora in horas_jornada:
            piezas = desglose_json.get(hora, 0) # Si no hay registro para la hora, es 0
            conteo_horas[hora] = piezas
            total_dia += piezas
            
        datos_procesados["dias"].append({
            "fecha": data['fecha'],
            "nombre_dia": nombre_dia,
            "conteo_por_hora": conteo_horas,
            "total_dia": total_dia
        })

    # Calcular resumen general
    if datos_procesados["dias"]:
        total_semanal = sum(d['total_dia'] for d in datos_procesados["dias"])
        promedio_diario = total_semanal / len(datos_procesados["dias"])
        dia_max = max(datos_procesados["dias"], key=lambda x: x['total_dia'])
        dia_min = min(datos_procesados["dias"], key=lambda x: x['total_dia'])
        
        # Calcular hora pico
        promedio_por_hora = {h: 0 for h in horas_jornada}
        for hora in horas_jornada:
            suma_hora = sum(d['conteo_por_hora'][hora] for d in datos_procesados["dias"])
            promedio_por_hora[hora] = suma_hora / len(datos_procesados["dias"])
        
        hora_pico = max(promedio_por_hora, key=promedio_por_hora.get)

        datos_procesados["resumen"] = {
            "total_semanal": total_semanal,
            "promedio_diario": round(promedio_diario, 2),
            "dia_max_produccion": f"{dia_max['nombre_dia']} ({dia_max['total_dia']} piezas)",
            "dia_min_produccion": f"{dia_min['nombre_dia']} ({dia_min['total_dia']} piezas)",
            "hora_pico": f"{hora_pico:02d}:00 - {hora_pico:02d}:59",
            "rango_fechas": f"{datos_procesados['dias'][0]['fecha']} al {datos_procesados['dias'][-1]['fecha']}"
        }
    
    return datos_procesados

def crear_graficas(datos_procesados):
    """Genera y guarda las gráficas de producción."""
    dias = datos_procesados['dias']
    resumen = datos_procesados['resumen']
    
    # 1. Gráfica de Barras: Producción Total por Día
    nombres_dias = [d['nombre_dia'] for d in dias]
    totales_diarios = [d['total_dia'] for d in dias]
    
    plt.figure(figsize=(10, 5))
    plt.bar(nombres_dias, totales_diarios, color='#4A90E2')
    plt.title('Producción Total por Día', fontsize=16)
    plt.ylabel('Piezas Producidas')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    path_barras = 'grafica_produccion_diaria.png'
    plt.savefig(path_barras)
    plt.close()

    # 2. Gráfica de Líneas: Ritmo de Producción Horario (Promedio)
    horas_jornada = range(6, 15)
    promedio_por_hora = []
    for hora in horas_jornada:
        suma_hora = sum(d['conteo_por_hora'][hora] for d in dias)
        promedio_por_hora.append(suma_hora / len(dias))
        
    etiquetas_horas = [f"{h:02d}:00" for h in horas_jornada]
    
    plt.figure(figsize=(10, 5))
    plt.plot(etiquetas_horas, promedio_por_hora, marker='o', color='#D0021B', linestyle='-')
    plt.title('Ritmo de Producción Horario (Promedio Semanal)', fontsize=16)
    plt.ylabel('Promedio de Piezas')
    plt.xlabel('Hora del Día')
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.tight_layout()
    path_lineas = 'grafica_ritmo_horario.png'
    plt.savefig(path_lineas)
    plt.close()

    return path_barras, path_lineas

def generar_contenido_html(datos, path_grafica_barras, path_grafica_lineas):
    """Construye el string HTML del reporte."""
    # Estilos CSS para que se vea bien
    html_style = """
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background-color: #f4f7f6; color: #333; }
        .container { max-width: 1200px; margin: 20px auto; padding: 20px; background-color: #fff; box-shadow: 0 0 10px rgba(0,0,0,0.1); border-radius: 8px; }
        h1, h2 { color: #005A9C; text-align: center; border-bottom: 2px solid #005A9C; padding-bottom: 10px; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; text-align: center; margin-bottom: 40px; }
        .summary-item { background-color: #eaf4ff; padding: 20px; border-radius: 8px; }
        .summary-item h3 { margin: 0 0 10px 0; color: #004080; }
        .summary-item p { font-size: 1.5em; font-weight: bold; margin: 0; color: #005A9C; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 40px; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: center; }
        th { background-color: #005A9C; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #eaf4ff; }
        .graph-container { text-align: center; margin-bottom: 20px; }
        .graph-container img { max-width: 90%; height: auto; border-radius: 8px; box-shadow: 0 0 8px rgba(0,0,0,0.1); }
    </style>
    """
    
    # --- Encabezado y Resumen ---
    html_header = f"""
    <h1>Reporte de Producción Semanal</h1>
    <h2>{datos['resumen']['rango_fechas']}</h2>
    <div class="summary">
        <div class="summary-item"><h3>Producción Total</h3><p>{datos['resumen']['total_semanal']:,}</p></div>
        <div class="summary-item"><h3>Promedio Diario</h3><p>{datos['resumen']['promedio_diario']}</p></div>
        <div class="summary-item"><h3>Día de Mayor Producción</h3><p>{datos['resumen']['dia_max_produccion']}</p></div>
        <div class="summary-item"><h3>Hora Pico (Promedio)</h3><p>{datos['resumen']['hora_pico']}</p></div>
    </div>
    """
    
    # --- Tabla Comparativa ---
    html_table = "<h2>Tabla Comparativa Diaria</h2><table><thead><tr><th>Fecha (Día)</th>"
    horas_jornada = range(6, 15)
    for hora in horas_jornada:
        html_table += f"<th>{hora:02d}:00</th>"
    html_table += "<th>Total Día</th></tr></thead><tbody>"
    
    for dia in datos['dias']:
        html_table += f"<tr><td>{dia['fecha']} ({dia['nombre_dia']})</td>"
        for hora in horas_jornada:
            html_table += f"<td>{dia['conteo_por_hora'][hora]}</td>"
        html_table += f"<td><b>{dia['total_dia']:,}</b></td></tr>"
    html_table += "</tbody></table>"
    
    # --- Gráficas ---
    html_graphs = f"""
    <h2>Gráficas de Rendimiento</h2>
    <div class="graph-container">
        <img src="{path_grafica_barras}" alt="Gráfica de Producción Diaria">
    </div>
    <div class="graph-container">
        <img src="{path_grafica_lineas}" alt="Gráfica de Ritmo Horario">
    </div>
    """

    # --- Unimos todo ---
    full_html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Reporte de Producción</title>
        {html_style}
    </head>
    <body>
        <div class="container">
            {html_header}
            {html_table}
            {html_graphs}
        </div>
    </body>
    </html>
    """
    return full_html

def generar_reporte_semanal():
    """
    Función principal que orquesta la creación del reporte HTML.
    """
    messagebox.showinfo("Generando Reporte", "Esto puede tardar unos segundos. Por favor, espera...")
    
    try:
        # 1. Buscar los últimos 5 archivos JSON de reporte
        archivos_json = sorted(glob.glob('reporte_*.json'))
        if not archivos_json:
            messagebox.showerror("Error", "No se encontraron archivos de reporte (.json) en esta carpeta.")
            return
        
        archivos_a_analizar = archivos_json[-5:] # Tomamos los últimos 5
        
        # 2. Procesar los datos
        datos = procesar_datos_reporte(archivos_a_analizar)
        
        # 3. Crear las gráficas
        path_barras, path_lineas = crear_graficas(datos)
        
        # 4. Generar el contenido HTML
        html_content = generar_contenido_html(datos, path_barras, path_lineas)
        
        # 5. Guardar el archivo HTML
        nombre_reporte = "reporte_semanal.html"
        with open(nombre_reporte, 'w', encoding='utf-8') as f:
            f.write(html_content)
            
        # 6. Abrir en el navegador
        webbrowser.open('file://' + os.path.realpath(nombre_reporte))
        
    except Exception as e:
        messagebox.showerror("Error Inesperado", f"Ocurrió un error al generar el reporte:\n{e}")

# --- FIN NUEVA SECCIÓN DE REPORTE HTML ---


def guardar_reporte_json():
    """
    Guarda el conteo del día actual en un archivo JSON si se ha producido algo.
    """
    if conteo_total_dia == 0:
        return

    datos_reporte = {
        "fecha": str(fecha_actual),
        "conteo_total_dia": conteo_total_dia,
        "desglose_por_hora": dict(conteo_por_hora) 
    }
    nombre_archivo = f"reporte_{fecha_actual.strftime('%Y-%m-%d')}.json"
    try:
        with open(nombre_archivo, 'w', encoding='utf-8') as f:
            json.dump(datos_reporte, f, indent=4)
    except Exception as e:
        messagebox.showerror("Error al Guardar", f"No se pudo guardar el reporte JSON:\n{e}")

def al_cerrar():
    """
    Guarda el reporte final y luego cierra la aplicación.
    """
    print("Guardando reporte antes de salir...")
    guardar_reporte_json()
    ventana.destroy()

# --- Lógica de la Aplicación (Funciones principales) ---
def seleccionar_archivo():
    global ruta_archivo_log, ultima_posicion, conteo_total_dia, conteo_por_hora, fecha_actual
    ruta = filedialog.askopenfilename(title="Selecciona el archivo de log (.txt)", filetypes=[("Archivos de texto", "*.txt")])
    if not ruta: return
    ruta_archivo_log = ruta
    try:
        ultima_posicion = os.path.getsize(ruta_archivo_log) 
    except FileNotFoundError:
        messagebox.showerror("Error", f"No se pudo encontrar el archivo en la ruta:\n{ruta_archivo_log}")
        return
    conteo_total_dia = 0
    conteo_por_hora.clear()
    fecha_actual = datetime.now().date()
    etiqueta_archivo.config(text=f"Vigilando: {os.path.basename(ruta_archivo_log)}")
    etiqueta_total.config(text=f"Piezas: {conteo_total_dia}")
    actualizar_lista_horas()
    vigilar_archivo()

def vigilar_archivo():
    global ruta_archivo_log, ultima_posicion, conteo_total_dia, conteo_por_hora, fecha_actual
    if not ruta_archivo_log: return
    try:
        if datetime.now().date() != fecha_actual:
            guardar_reporte_json()
            fecha_actual = datetime.now().date()
            conteo_total_dia = 0
            conteo_por_hora.clear()

        tamano_actual = os.path.getsize(ruta_archivo_log)
        if tamano_actual > ultima_posicion:
            with open(ruta_archivo_log, 'r', errors='ignore') as f:
                f.seek(ultima_posicion)
                lineas_nuevas = f.readlines()
                for linea in lineas_nuevas:
                    if linea.strip():
                        conteo_total_dia += 1
                        hora_actual = datetime.now().hour
                        conteo_por_hora[hora_actual] += 1
                ultima_posicion = f.tell()
            actualizar_interfaz()
    except FileNotFoundError:
        messagebox.showerror("Error", "Se perdió la conexión con el archivo. Por favor, selecciónalo de nuevo.")
        etiqueta_archivo.config(text="Archivo no seleccionado")
        ruta_archivo_log = "" 
        return
    ventana.after(1000, vigilar_archivo)

def actualizar_interfaz():
    etiqueta_total.config(text=f"Piezas: {conteo_total_dia}")
    actualizar_lista_horas()

def actualizar_lista_horas():
    lista_horas.delete(0, tk.END)
    horas_ordenadas = sorted(conteo_por_hora.keys())
    for hora in horas_ordenadas:
        piezas = conteo_por_hora[hora]
        hora_12 = hora % 12
        if hora_12 == 0: hora_12 = 12
        periodo = "AM" if hora < 12 else "PM"
        texto_item = f"{hora_12:02d}:00 {periodo} - {hora_12:02d}:59 {periodo} | {piezas} piezas"
        lista_horas.insert(tk.END, texto_item)

def mostrar_qr_desbloqueo():
    # ... (El código de esta función no cambia)
    global qr_image_ref
    ventana_qr = tk.Toplevel(ventana)
    ventana_qr.title("Código de Desbloqueo")
    ventana_qr.geometry("300x350")
    ventana_qr.resizable(False, False)
    ventana_qr.config(padx=10, pady=10)
    ventana_qr.attributes('-topmost', True)
    label_instruccion = tk.Label(ventana_qr, text="Escanee este código en la máquina", font=("Helvetica", 12))
    label_instruccion.pack(pady=10)
    password = "Calidadtorque"
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
    qr.add_data(password)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    photo_image = ImageTk.PhotoImage(img)
    qr_image_ref = photo_image
    label_qr = tk.Label(ventana_qr, image=photo_image)
    label_qr.pack(pady=10)

def eliminar_ultimo_registro():
    # ... (El código de esta función no cambia)
    global ruta_archivo_log, ultima_posicion, conteo_total_dia, conteo_por_hora
    if not ruta_archivo_log:
        messagebox.showwarning("Advertencia", "Primero debes seleccionar un archivo de log.")
        return
    confirmar = messagebox.askyesno("Confirmar Eliminación", "¿Estás seguro de que quieres eliminar el último registro?\n\nEsta acción no se puede deshacer.")
    if not confirmar: return
    try:
        with open(ruta_archivo_log, 'r', errors='ignore') as f: lineas = f.readlines()
        if not lineas:
            messagebox.showinfo("Información", "El archivo ya está vacío.")
            return
        with open(ruta_archivo_log, 'w') as f: f.writelines(lineas[:-1])
        hora_actual = datetime.now().hour
        if conteo_total_dia > 0: conteo_total_dia -= 1
        if conteo_por_hora[hora_actual] > 0: conteo_por_hora[hora_actual] -= 1
        ultima_posicion = os.path.getsize(ruta_archivo_log)
        actualizar_interfaz()
        messagebox.showinfo("Éxito", "El último registro fue eliminado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al modificar el archivo:\n{e}")

def generar_pieza_aleatoria():
    # ... (El código de esta función no cambia)
    numero_pieza_fijo = "3QG971793B"
    serial_aleatorio = "".join(random.choice('0123456789') for _ in range(16))
    codigo_completo = f"{numero_pieza_fijo} {serial_aleatorio}"
    ventana.clipboard_clear()
    ventana.clipboard_append(codigo_completo)
    messagebox.showinfo("Código Generado", f"Se copió al portapapeles:\n\n{codigo_completo}")

# --- Creación de la Interfaz Gráfica (Ventana) ---

ventana = tk.Tk()
ventana.title("Monitor Producción")
ventana.geometry("450x420") # ### MODIFICADO: Hice la ventana un poco más ancha para el nuevo botón ###
ventana.minsize(360, 250)
ventana.config(padx=10, pady=10)
ventana.attributes('-topmost', True)

# --- Configuración de la Grilla (Grid) ---
ventana.grid_columnconfigure(0, weight=1)
ventana.grid_columnconfigure(1, weight=1)
ventana.grid_rowconfigure(2, weight=1)

# --- Barra de Herramientas ---
frame_botones = tk.Frame(ventana)
frame_botones.grid(row=0, column=0, columnspan=2, pady=(0, 5), sticky="ew")

boton_seleccionar = tk.Button(frame_botones, text="Abrir Log", command=seleccionar_archivo)
boton_seleccionar.pack(side="left", padx=2, expand=True, fill="x")

boton_desbloqueo = tk.Button(frame_botones, text="Desbloquear", command=mostrar_qr_desbloqueo)
boton_desbloqueo.pack(side="left", padx=2, expand=True, fill="x")

boton_eliminar = tk.Button(frame_botones, text="Eliminar", command=eliminar_ultimo_registro)
boton_eliminar.pack(side="left", padx=2, expand=True, fill="x")

boton_generador = tk.Button(frame_botones, text="Generar", command=generar_pieza_aleatoria)
boton_generador.pack(side="left", padx=2, expand=True, fill="x")

### NUEVO: Botón para generar el reporte semanal ###
boton_reporte = tk.Button(frame_botones, text="Reporte", command=generar_reporte_semanal, bg="#C8E6C9")
boton_reporte.pack(side="left", padx=2, expand=True, fill="x")

# --- Panel de Información ---
etiqueta_archivo = tk.Label(ventana, text="Archivo no seleccionado", fg="gray", font=("Helvetica", 9))
etiqueta_archivo.grid(row=1, column=0, sticky="w", padx=5)

etiqueta_total = tk.Label(ventana, text="Piezas: 0", font=("Helvetica", 14, "bold"))
etiqueta_total.grid(row=1, column=1, sticky="e", padx=5)

# --- Lista de Producción ---
lista_horas = tk.Listbox(ventana, height=10, borderwidth=1, relief="solid")
lista_horas.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(5, 0))

# --- Scrollbar para la lista ---
scrollbar = tk.Scrollbar(lista_horas, orient="vertical", command=lista_horas.yview)
lista_horas.config(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

# --- Capturamos el evento de cierre de la ventana ---
ventana.protocol("WM_DELETE_WINDOW", al_cerrar)

ventana.mainloop()