import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import os
import re
from datetime import datetime

archivos = []
horas = {}
indice_actual = 0
ruta_historico = ""

pd.options.mode.chained_assignment = None

COLUMNAS_FINALES = [
    "POINT",
    "SSID",
    "BSSID",
    "SIGNAL",
    "AUTHENTICATION",
    "ENCRYPTION",
    "DAY_DATE",
    "MONTH_DATE",
    "YEAR_DATE",
    "HOUR_DATE",
    "MERIDIEM_DATE",
    "MES"
]

# =========================
# VALIDAR FORMATO HORA
# =========================
def validar_hora(hora):
    patron = r"^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$"
    return re.match(patron, hora) is not None

# =========================
# DETECTAR ZONA
# =========================
def detectar_zona(path):
    nombre = os.path.basename(path).upper()
    match = re.search(r'\b([A-Z])\b', nombre)
    if match:
        return match.group(1)
    return nombre[0]

# =========================
# OBTENER MES DINÁMICO
# =========================
def obtener_mes(fecha_usuario):
    fecha = datetime.strptime(fecha_usuario, "%d/%m/%Y")

    meses = [
        "Enero","Febrero","Marzo","Abril",
        "Mayo","Junio","Julio","Agosto",
        "Septiembre","Octubre","Noviembre","Diciembre"
    ]

    mes = meses[fecha.month-1]
    anio = str(fecha.year)[-2:]

    return f"{mes}{anio}"

# =========================
# DESCOMPONER FECHA 
# =========================
def descomponer_fecha(fecha_usuario, hora):

    fecha = datetime.strptime(fecha_usuario, "%d/%m/%Y")
    hora_obj = datetime.strptime(hora, "%H:%M:%S").time()

    fecha_final = datetime.combine(fecha, hora_obj)

    return {
        "DAY_DATE": fecha_final.day,
        "MONTH_DATE": fecha_final.month,
        "YEAR_DATE": fecha_final.year,
        "HOUR_DATE": fecha_final.strftime("%I:%M:%S").lstrip("0"),
        "MERIDIEM_DATE": fecha_final.strftime("%p")
    }

# =========================
# VALIDAR SI EXCEL ESTA ABIERTO
# =========================
def archivo_esta_abierto(ruta):
    try:
        with open(ruta, 'a'):
            pass
        return False
    except PermissionError:
        return True

# =========================
# SELECCIONAR HISTORICO
# =========================
def seleccionar_historico():
    global ruta_historico

    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo histórico",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if ruta:
        ruta_historico = ruta
        label_hist.config(text=f"Histórico: {os.path.basename(ruta)}")

# =========================
# SELECCIONAR CSV
# =========================
def seleccionar_archivos():
    global archivos, horas, indice_actual

    files = filedialog.askopenfilenames(
        title="Seleccionar escaneos WiFi",
        filetypes=[("CSV files", "*.csv")]
    )

    if not files:
        return

    archivos = sorted(list(files), key=lambda x: detectar_zona(x))

    horas = {}
    indice_actual = 0

    actualizar_lista()
    pedir_siguiente()

# =========================
# ACTUALIZAR LISTA
# =========================
def actualizar_lista():
    lista.delete(0, tk.END)

    for archivo in archivos:
        zona = detectar_zona(archivo)

        if zona in horas:
            lista.insert(tk.END, f"{zona} ✓ {horas[zona]}")
        else:
            lista.insert(tk.END, f"{zona} ⏳ pendiente")

# =========================
# PEDIR HORA
# =========================
def pedir_siguiente():
    global indice_actual

    if indice_actual >= len(archivos):
        zona_label.config(text="Todas las zonas tienen hora asignada")
        return

    zona = detectar_zona(archivos[indice_actual])

    zona_label.config(
        text=f"Ingrese hora HH:MM:SS para zona {zona}"
    )

# =========================
# GUARDAR HORA
# =========================
def guardar_hora():
    global indice_actual

    if indice_actual >= len(archivos):
        return

    hora = entrada_hora.get().strip()

    if not validar_hora(hora):
        messagebox.showerror("Formato inválido","Formato requerido HH:MM:SS")
        return

    zona = detectar_zona(archivos[indice_actual])

    horas[zona] = hora

    entrada_hora.delete(0, tk.END)

    indice_actual += 1

    actualizar_lista()
    pedir_siguiente()

# =========================
# EDITAR HORA
# =========================
def editar_hora():
    seleccion = lista.curselection()

    if not seleccion:
        return

    index = seleccion[0]
    zona = detectar_zona(archivos[index])

    hora = entrada_hora.get().strip()

    if not validar_hora(hora):
        messagebox.showerror("Formato inválido","Formato requerido HH:MM:SS")
        return

    horas[zona] = hora

    actualizar_lista()
    entrada_hora.delete(0, tk.END)

# =========================
# PROCESAR ETL
# =========================
def procesar_etl():
    global ruta_historico

    if ruta_historico == "":
        messagebox.showerror("Error", "Debes seleccionar el archivo histórico")
        return

    if len(horas) != len(archivos):
        messagebox.showerror("Error", "Debes asignar hora a todas las zonas")
        return

    if archivo_esta_abierto(ruta_historico):
        messagebox.showwarning("Archivo abierto", "Debes cerrar el Excel")
        return

    fecha_usuario = fecha_entry.get().strip()
    try:
        datetime.strptime(fecha_usuario, "%d/%m/%Y")
    except:
        messagebox.showerror("Error", "Formato fecha DD/MM/YYYY")
        return

    mes_valor = obtener_mes(fecha_usuario)
    df_hist = pd.read_excel(ruta_historico, sheet_name="BD")

    df_hist.columns = (
    df_hist.columns
    .astype(str)
    .str.replace("\ufeff", "", regex=False)
    .str.replace("ï»¿", "", regex=False)
    .str.replace("Ï»¿", "", regex=False)
    .str.strip()
    .str.upper()
)

    # Limpieza inicial del histórico
    if "DATE(UTC)" in df_hist.columns:
        df_hist.drop(columns=["DATE(UTC)"], inplace=True)

    datos_total = []
    barra["maximum"] = len(archivos)
    barra["value"] = 0

    # =========================
    # PRIMERA PASADA: leer CSVs y construir df_nuevos
    # =========================
    for i, archivo in enumerate(archivos):
        zona = detectar_zona(archivo)

        try:
            # 1. Leer el CSV
            df_raw = pd.read_csv(archivo, encoding="latin1")
            df_raw.columns = df_raw.columns.str.strip()

            # 2. Crear un DataFrame que solo tenga las columnas finales
            df = pd.DataFrame()

            # 3. Mapeo seguro de columnas
            mapeo = {
                "SSID": "SSID",
                "BSSID": "BSSID",
                "SIGNAL": "SIGNAL",
                "AUTHENTICATION": "AUTHENTICATION",
                "ENCRYPTION": "ENCRYPTION"
            }

            for col_final, col_csv in mapeo.items():
                if col_csv in df_raw.columns:
                    df[col_final] = df_raw[col_csv].fillna("").astype(str)
                else:
                    df[col_final] = ""

            # 4. Agregar campos calculados
            fecha_partes = descomponer_fecha(fecha_usuario, horas[zona])
            df["DAY_DATE"] = fecha_partes["DAY_DATE"]
            df["MONTH_DATE"] = fecha_partes["MONTH_DATE"]
            df["YEAR_DATE"] = fecha_partes["YEAR_DATE"]
            df["HOUR_DATE"] = fecha_partes["HOUR_DATE"]
            df["MERIDIEM_DATE"] = fecha_partes["MERIDIEM_DATE"]
            df["POINT"] = zona
            df["MES"] = mes_valor

            # Asegurar que todas las columnas existen y están en orden
            for col in COLUMNAS_FINALES:
                if col not in df.columns:
                    df[col] = ""
            df = df[COLUMNAS_FINALES]

            datos_total.append(df)

        except Exception as e:
            print(f"Error procesando {archivo}: {e}")
            continue

        barra["value"] = i + 1
        ventana.update_idletasks()

    if not datos_total:
        messagebox.showerror("Error", "No se pudieron procesar archivos")
        return

    # =========================
    # PRIMER GUARDO: df_nuevos a partir de los datos de CSVs
    # =========================
    df_nuevos = pd.concat(datos_total, ignore_index=True)

    # Asegurar columnas en df_nuevos
    df_nuevos.columns = df_nuevos.columns.str.strip()
    for col in COLUMNAS_FINALES:
        if col not in df_nuevos.columns:
            df_nuevos[col] = ""
    df_nuevos = df_nuevos[COLUMNAS_FINALES]

    # CONCAT parcial con histórico
    df_final_parcial = pd.concat([df_hist, df_nuevos], ignore_index=True)

    try:
        with pd.ExcelWriter(ruta_historico, engine="openpyxl", mode="w") as writer:
            df_final_parcial.to_excel(writer, sheet_name="BD", index=False)
        messagebox.showinfo("Proceso completado", f"Registros agregados: {len(df_nuevos)}")
    except Exception as e:
        messagebox.showerror("Error al guardar", str(e))
        return

    barra["value"] = 0

    # =========================
    # SEGUNDA PASADA: reconstruir df_nuevos para historial y asegurar SSID
    # =========================
    datos_total = []  # reiniciar para la segunda pasada

    for i, archivo in enumerate(archivos):
        zona = detectar_zona(archivo)

        try:
            # 1. Leer el archivo
            df_raw = pd.read_csv(archivo, encoding="latin1")

            # 2. LIMPIEZA CRÍTICA: Eliminar espacios en los nombres de las columnas
            df_raw.columns = df_raw.columns.str.strip()

            # 3. Crear DataFrame objetivo
            df = pd.DataFrame(columns=COLUMNAS_FINALES)

            # 4. Mapeo seguro: usamos .get() para evitar errores si el nombre varía ligeramente
            df["SSID"] = df_raw.get("SSID", "")
            df["BSSID"] = df_raw.get("BSSID", "")
            df["SIGNAL"] = df_raw.get("SIGNAL", "")
            # Ajusta AUTHENTICATION/ENCRYPTION si el CSV los tiene truncados (ej: AUTHENTI)
            df["AUTHENTICATION"] = df_raw.get("AUTHENTICATION", df_raw.get("AUTHENTI", ""))
            df["ENCRYPTION"] = df_raw.get("ENCRYPTION", df_raw.get("ENCRYPTIO", ""))

            # 5. Rellenar lo demás
            fecha_partes = descomponer_fecha(fecha_usuario, horas[zona])
            df["DAY_DATE"] = fecha_partes["DAY_DATE"]
            df["MONTH_DATE"] = fecha_partes["MONTH_DATE"]
            df["YEAR_DATE"] = fecha_partes["YEAR_DATE"]
            df["HOUR_DATE"] = fecha_partes["HOUR_DATE"]
            df["MERIDIEM_DATE"] = fecha_partes["MERIDIEM_DATE"]
            df["POINT"] = zona
            df["MES"] = mes_valor

        except Exception as e:
            print(f"Error procesando {archivo}: {e}")
            continue

        # =========================
        # ELIMINAR DATE(UTC) si existiera
        # =========================
        if "DATE(UTC)" in df.columns:
            df.drop(columns=["DATE(UTC)"], inplace=True)

        # =========================
        # NUEVAS COLUMNAS DE FECHA
        # =========================
        fecha_partes = descomponer_fecha(fecha_usuario, horas[zona])
        df["DAY_DATE"] = fecha_partes["DAY_DATE"]
        df["MONTH_DATE"] = fecha_partes["MONTH_DATE"]
        df["YEAR_DATE"] = fecha_partes["YEAR_DATE"]
        df["HOUR_DATE"] = fecha_partes["HOUR_DATE"]
        df["MERIDIEM_DATE"] = fecha_partes["MERIDIEM_DATE"]

        # =========================
        # CAMPOS FIJOS
        # =========================
        df["POINT"] = zona
        df["MES"] = mes_valor

        # =========================
        # SSID
        # =========================
        # Normalizamos de nuevo por si acaso
        df.columns = df.columns.str.strip()

        if "SSID" not in df.columns:
            df["SSID"] = ""
        else:
            df["SSID"] = df["SSID"].fillna("").astype(str).str.strip()

        # =========================
        # AJUSTAR COLUMNAS
        # =========================
        for col in COLUMNAS_FINALES:
            if col not in df.columns:
                df[col] = ""

        df = df[COLUMNAS_FINALES]

        datos_total.append(df)

        barra["value"] = i + 1
        ventana.update_idletasks()

    if not datos_total:
        messagebox.showerror("Error","No se pudieron procesar archivos")
        return

    # LIMPIAR HISTÓRICO (estructura final)
    df_hist.columns = df_hist.columns.str.strip()
    for col in COLUMNAS_FINALES:
        if col not in df_hist.columns:
            df_hist[col] = ""
    df_hist = df_hist[COLUMNAS_FINALES]

    # UNIR NUEVOS obtenidos en segunda pasada
    df_nuevos = pd.concat(datos_total, ignore_index=True)
    df_nuevos.columns = df_nuevos.columns.str.strip()
    for col in COLUMNAS_FINALES:
        if col not in df_nuevos.columns:
            df_nuevos[col] = ""
    df_nuevos = df_nuevos[COLUMNAS_FINALES]


    # CONCAT FINAL
    df_final = pd.concat([df_hist, df_nuevos], ignore_index=True)

    try:
        with pd.ExcelWriter(
            ruta_historico,
            engine="openpyxl",
            mode="w"
        ) as writer:
            df_final.to_excel(writer, sheet_name="BD", index=False)
    except Exception as e:
        messagebox.showerror("Error al guardar", str(e))
        return

    messagebox.showinfo(
        "Proceso completado",
        f"Registros agregados: {len(df_nuevos)}"
    )

    barra["value"] = 0

# =========================
# INTERFAZ
# =========================
ventana = tk.Tk()
ventana.title("WiFi Scan Processor - Giesecke+Devrient")
ventana.geometry("720x700")

titulo = tk.Label(ventana,text="WiFi Scan ETL Processor",font=("Segoe UI",16,"bold"))
titulo.pack(pady=10)

sub = tk.Label(ventana,text="Herramienta interna de procesamiento de escaneos WiFi",font=("Segoe UI",10))
sub.pack()

boton_hist = tk.Button(ventana,text="Seleccionar archivo histórico Excel",command=seleccionar_historico)
boton_hist.pack(pady=5)

label_hist = tk.Label(ventana,text="Histórico no seleccionado")
label_hist.pack()

boton = tk.Button(ventana,text="Seleccionar archivos CSV",command=seleccionar_archivos)
boton.pack(pady=10)

fecha_label = tk.Label(ventana,text="Fecha escaneo (DD/MM/YYYY)")
fecha_label.pack()

fecha_entry = tk.Entry(ventana,width=20)
fecha_entry.pack(pady=5)

zona_label = tk.Label(ventana,text="Seleccione archivos para comenzar")
zona_label.pack()

entrada_hora = tk.Entry(ventana,width=20)
entrada_hora.pack(pady=5)

guardar = tk.Button(ventana,text="Guardar hora",command=guardar_hora)
guardar.pack()

lista = tk.Listbox(ventana,width=70,height=10)
lista.pack(pady=15)

editar = tk.Button(ventana,text="Editar hora zona seleccionada",command=editar_hora)
editar.pack()

barra = Progressbar(ventana,orient="horizontal",length=500,mode="determinate")
barra.pack(pady=20)

procesar = tk.Button(
    ventana,
    text="Procesar ETL",
    command=procesar_etl,
    bg="#0B7A0B",
    fg="white",
    font=("Segoe UI",11,"bold"),
    width=20
)
procesar.pack(pady=10)

ventana.mainloop()