# archivo para manejar la logica del sistema
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt

import pandas as pd

from PIL import Image, ImageDraw, ImageFont
import qrcode
from reportlab.lib.pagesizes import A4, letter
from reportlab.pdfgen import canvas

pdf = canvas

# inicializamos la ventana principal de tkinter
root = tk.Tk()
root.title("Sistema de QR inventario")  # titulo del sistema
root.geometry("800x600")  # tamaño de la ventana del sistema

# variable global para los dataframes
df = pd.DataFrame()


# funcion para cargar nuestros archivos excel
def load_csv():
    global df
    file_path = filedialog.askopenfilename(
        filetypes=[("XLSX files", "*.xlsx")]
    )
    if file_path:
        try:
            df = pd.read_excel(file_path, dtype={"Material": str})
            display_data(df)
            print(df[["Material", "Descripción del material"]].to_numpy())
            print(df["Material"].to_numpy())
        except Exception as e:
            messagebox.showerror("error", f"fallo al cargar el archivo: {e}")


def generate_qr():
    if df.empty:
        messagebox.showwarning("Atención", "Primero debes cargar un archivo Excel.")
        return

    data_num_descripcion = df[["Material", "Descripción del material"]]
    data = data_num_descripcion

    imagenes_para_pdf = []

    # --- RUTA EN EL ESCRITORIO ---
    carpeta_escritorio = Path.home() / "Desktop" / "material-qr-listos"
    os.makedirs(carpeta_escritorio, exist_ok=True)

    for index, q in data.iterrows():
        num_material = q["Material"]
        desc_material = q["Descripción del material"]

        datos_completo_desc = desc_material
        datos_completo_num = f"{num_material}"

        # Configuración del QR
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )

        qr.add_data(datos_completo_num)
        qr.make(fit=True)

        qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        qr_width, qr_height = qr_img.size

        texto = datos_completo_desc
        extra_pantalla_abajo = 50

        nuevo_ancho = qr_width
        nuevo_alto = qr_height + extra_pantalla_abajo
        imagen_final = Image.new("RGB", (nuevo_ancho, nuevo_alto), "white")

        imagen_final.paste(qr_img, (0, 0))

        draw = ImageDraw.Draw(imagen_final)

        try:
            font = ImageFont.truetype("arial.ttf", 12)
        except IOError:
            font = ImageFont.load_default()

        bbox = draw.textbbox((0, 0), texto, font=font, stroke_width=5, align="center")
        text_width = bbox[2] - bbox[0]

        x = (nuevo_ancho - text_width) // 2
        y = qr_height + 10

        draw.text((x, y), texto, fill="black", font=font)

        nombre_archivo = f"{num_material}.png"
        ruta_completa = carpeta_escritorio / nombre_archivo

        # Guardar la imagen en la carpeta del Escritorio
        imagen_final.save(ruta_completa, "png")

        imagenes_para_pdf.append(imagen_final)

        qr.clear()

    # --- GENERACIÓN DEL PDF EN EL ESCRITORIO ---
    if not imagenes_para_pdf:
        print("no hay imagenes para procesar")
        return

    ancho_hoja = 2480
    alto_hoja = 3508

    columnas = 3
    filas = 4
    qrs_por_pagina = columnas * filas

    ancho_item_pdf = 600
    alto_item_pdf = 700

    margen_izquierdo = 180
    margen_superior = 200
    espacio_horizontal = 160
    espacio_vertical = 120

    paginas_pdf = []
    hoja_actual = None

    for indice, img_qr_final in enumerate(imagenes_para_pdf):
        if indice % qrs_por_pagina == 0:
            if hoja_actual:
                paginas_pdf.append(hoja_actual)
            hoja_actual = Image.new("RGB", (ancho_hoja, alto_hoja), "white")

        indice_en_pagina = indice % qrs_por_pagina
        col = indice_en_pagina % columnas
        fila = indice_en_pagina // columnas

        x = margen_izquierdo + col * (ancho_item_pdf + espacio_horizontal)
        y = margen_superior + fila * (alto_item_pdf + espacio_vertical)

        img_redimensionada = img_qr_final.resize(
            (ancho_item_pdf, alto_item_pdf), Image.Resampling.LANCZOS
        )

        hoja_actual.paste(img_redimensionada, (x, y))

    if hoja_actual:
        paginas_pdf.append(hoja_actual)

    archivo_pdf_salida = "qrs_generados.pdf"
    if paginas_pdf:
        ruta_pdf_escritorio = Path.home() / "Desktop" / archivo_pdf_salida

        paginas_pdf[0].save(
            ruta_pdf_escritorio,
            save_all=True,
            append_images=paginas_pdf[1:],
        )
        print(f"🎉 ¡Proceso completo! PNGs guardados en: {carpeta_escritorio}")
        print(f"📄 PDF creado en: {ruta_pdf_escritorio}")


def display_data(data):
    tree.delete(*tree.get_children())

    tree["columns"] = list(data.columns)
    tree["show"] = "headings"

    for col in data.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))


def add_new_data_qr():
    global df
    input_num_material = codigo_material.get()
    input_descripcion_material = descripcion_material.get()

    if not input_num_material:
        messagebox.showwarning("Atención", "Debes ingresar un código de material.")
        return

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )

    qr.add_data(input_num_material)
    qr.make(fit=True)

    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    qr_width, qr_height = qr_img.size

    texto = input_descripcion_material
    extra_pantalla_abajo = 50

    nuevo_ancho = qr_width
    nuevo_alto = qr_height + extra_pantalla_abajo
    imagen_final = Image.new("RGB", (nuevo_ancho, nuevo_alto), "white")

    imagen_final.paste(qr_img, (0, 0))

    draw = ImageDraw.Draw(imagen_final)

    try:
        font = ImageFont.truetype("arial.ttf", 12)
    except IOError:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), texto, font=font)
    text_width = bbox[2] - bbox[0]

    x = (nuevo_ancho - text_width) // 2
    y = qr_height + 10

    draw.text((x, y), texto, fill="black", font=font)

    # --- RUTA EN EL ESCRITORIO ---
    carpeta_escritorio = Path.home() / "Desktop" / "material-qr-listos"
    os.makedirs(carpeta_escritorio, exist_ok=True)

    nombre_archivo = f"{input_num_material}.png"
    ruta_completa = carpeta_escritorio / nombre_archivo

    imagen_final.save(ruta_completa, "png")
    messagebox.showinfo("Éxito", f"QR guardado en: {ruta_completa}")


def filter_data():
    global df
    column = entry_column.get()
    value = entry_value.get()

    if column not in df.columns:
        messagebox.showwarning("Cuidado", "Columna no encontrada")
        return

    filtered_df = df[df[column].astype(str) == value]

    if filtered_df.empty:
        messagebox.showinfo("Info", "No coinciden los registros encontrados")

    display_data(filtered_df)


def printqr():
    ruta_archivo = filedialog.askopenfilenames(
        title="Selecciona el archivo para imprimir",
        filetypes=[
            ("Archivos PDF", "*.pdf"),
            ("Docs de texto", "*.txt"),
            ("Archivos de imagen", "*.*"),
        ],
    )

    if ruta_archivo:
        exitos = 0
        errores = 0
        for ruta in ruta_archivo:
            try:
                ruta_normalizada = os.path.abspath(ruta)
                os.startfile(ruta_normalizada, "print")
                exitos += 1
            except Exception as e:
                messagebox.showerror(
                    "Error", f"No se pudo imprimir el archivo.\nDetalles: {e}"
                )
                errores += 1

        if errores == 0:
            messagebox.showinfo("Éxito!", f"Se enviaron {exitos} archivos a imprimir")
        else:
            messagebox.showwarning(
                "Resultado",
                f"Se enviaron {exitos} a imprimir, pero {errores} fallaron",
            )
    else:
        messagebox.showwarning("Cancelado", "No se seleccionó ningún archivo")


# ----------------------------------------------------- CONFIGURACIÓN DE WIDGETS ---------------------------------------------------------------

btn_load = tk.Button(
    root,
    text="Cargar excel",
    command=load_csv,
    bg="#0b5399",
    fg="white",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5,
)
btn_load.pack(pady=10)

frame_filter = tk.Frame(root)
frame_filter.pack(pady=10)

entry_column = tk.Entry(frame_filter, width=20)
entry_column.grid(row=0, column=0, padx=5)
entry_column.insert(0, "Código de material")

entry_value = tk.Entry(frame_filter, width=20)
entry_value.grid(row=0, column=1, padx=5)
entry_value.insert(0, "Descripción")

btn_filter = tk.Button(
    frame_filter,
    text="Filtrar busqueda",
    command=filter_data,
    bg="#0b5399",
    fg="white",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5,
)
btn_filter.grid(row=0, column=2, padx=5, pady=5)

btn_generateqr = tk.Button(
    frame_filter,
    text="Generar QRs",
    command=generate_qr,
    bg="#0b5399",
    fg="white",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5,
)
btn_generateqr.grid(row=0, column=3, padx=5, pady=5)

codigo_material = tk.Entry(frame_filter, width=20)
codigo_material.grid(row=1, column=0, padx=5)

descripcion_material = tk.Entry(frame_filter, width=20)
descripcion_material.grid(row=1, column=1, padx=5)

btn_generate_new_qr = tk.Button(
    frame_filter,
    text="Generar nuevo QR",
    command=add_new_data_qr,
    bg="#0b5399",
    fg="white",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5,
)
btn_generate_new_qr.grid(row=1, column=2, padx=5, pady=5)

btn_print = tk.Button(
    frame_filter,
    text="Imprimir QRs",
    command=printqr,
    bg="#0b5399",
    fg="white",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5,
)
btn_print.grid(row=1, column=3, padx=5, pady=5)


# ---------------------------------------seccion para la configuracion de agregar un nuevo material-------------------------------------
codigo_material = tk.Entry(frame_filter, width=20, fg="gray")
codigo_material.grid(row=1, column=0, padx=5)
codigo_material.insert(0, "Código de material")


def on_focus_in_codigo(event):
  if codigo_material.get() == "Código de material" or entry_column.get() == "Código de material":
    codigo_material.delete(0, tk.END)
    codigo_material.config(fg="black")
   

def on_focus_in_codigo_filter(event):
    if entry_column.get() == "Código de material":
        entry_column.delete(0, tk.END)
        entry_column.config(fg="black")

def on_focus_out_codigo_filter(event):
    if not entry_column.get():
        entry_column.insert(0, "Código de material")
        entry_column.config(fg="gray")


def on_focus_out_codigo(event):
  if not codigo_material.get() or entry_column.get():
    codigo_material.insert(0, "Código de material")
    codigo_material.config(fg="gray")



codigo_material.bind("<FocusIn>", on_focus_in_codigo)
codigo_material.bind("<FocusOut>", on_focus_out_codigo)
entry_column.bind("<FocusIn>", on_focus_in_codigo_filter)
entry_column.bind("<FocusOut>", on_focus_out_codigo_filter)


descripcion_material = tk.Entry(frame_filter, width=20, fg="gray")
descripcion_material.grid(row=1, column=1, padx=5)
descripcion_material.insert(0, "Descripción")


def on_focus_in_desc(event):
  if descripcion_material.get() == "Descripción" or entry_value.get() == "Descripción":
    descripcion_material.delete(0, tk.END)
    descripcion_material.config(fg="black")
    


def on_focus_in_desc_filter(event):
    if entry_value.get() == "Descripción":
        entry_value.delete(0, tk.END)
        entry_value.config(fg="black")


def on_focus_out_desc_filter(event):
    if not entry_value.get():
        entry_value.insert(0, "Descripción")
        entry_value.config(fg="gray")

def on_focus_out_desc(event):
  if not descripcion_material.get() or entry_value.get():
    descripcion_material.insert(0, "Descripción")
    descripcion_material.config(fg="gray")
    


descripcion_material.bind("<FocusIn>", on_focus_in_desc)
descripcion_material.bind("<FocusOut>", on_focus_out_desc)
entry_value.bind("<FocusIn>", on_focus_in_desc_filter)
entry_value.bind("<FocusOut>", on_focus_out_desc_filter)

tree = ttk.Treeview(root)
tree.pack(expand=True, fill="both")

root.mainloop()