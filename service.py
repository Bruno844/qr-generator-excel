#archivo para manejar la logica del sistema
import tkinter as tk
from tkinter import ttk , filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os
from pathlib import Path
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas


pdf = canvas

#inicializamos la ventana principal de tkinter
root = tk.Tk()
root.title("Sistema de QR inventario") #titulo del sistema
root.geometry("800x600") #tamaño de la ventana del sistema

#variable global para los dataframes
df = pd.DataFrame()


#funcion para cargar nuestros archivos excel
def load_csv():
    global df #poner global df sirve para indicarle a python que vamos a modificar o reasignar una variable llamada df en el entorno global, sin la necesidad de crear una nueva variable local para actualizarla, si no usar la existente mejor
    file_path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")]) #capturamos el tipo de archivo que se seleccione, y se pueden esos dos que configuramos
    if file_path:
        try:
            #aca especificamos con la propiedad llamada dtype, para poder forzar a cambiar el tipo de dato que puede ser una columna del archivo que seleccionemos, en este caso la columna material era de tipo entero o numerico, y lo pasamos a string(str) para mejor lectura de datos(recomendado asi)
            df = pd.read_excel(file_path, dtype={'Material': str}) #pandas y su funcion de read_excel para leer el archivo seleccionado
            display_data(df) #funcion que usamos para poder mostrar la data
            print(df[["Material", "Descripción del material"]].to_numpy())
            print(df['Material'].to_numpy())
            #print(df[0:3])
        except Exception as e:
            messagebox.showerror("error", f"fallo al cargar el archivo: {e}")



#num_material: any, desc_material: any
def generate_qr():
   
    #seleccion de las dos tablas
    data_num_descripcion = df[["Material", "Descripción del material"]]
    # Añadir los datos
    data = data_num_descripcion

    imagenes_para_pdf = []


    for index, q in data.iterrows():
        num_material = q["Material"]
        desc_material = q["Descripción del material"]

        datos_completo_desc = desc_material
        datos_completo_num = f"{num_material}"

        # Configuración del QR
        qr = qrcode.QRCode(
            version=1, # Tamaño del QR (del 1 al 40). El 1 es el más pequeño.
            error_correction=qrcode.constants.ERROR_CORRECT_L, # Tolerancia a daños o que salga medio pixeleado
            box_size=10, # Cuántos píxeles mide cada "cuadrado" del QR
            border=4, # Grosor del borde blanco (mínimo 4)
        )

        #solo vamos a agregar el dato del numero de material
        #qr.add_data(datos_completo_desc)
        qr.add_data(datos_completo_num)
        qr.make(fit=True)

        qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        qr_width, qr_height = qr_img.size

        #configuramos el texto y el espacio extra
        texto = datos_completo_desc
        extra_pantalla_abajo = 50 #pixeles extra abajo para el texto

        #creamos un lienzo nuevo mas alto que el qr original
        nuevo_ancho = qr_width
        nuevo_alto = qr_height + extra_pantalla_abajo
        imagen_final = Image.new("RGB", (nuevo_ancho, nuevo_alto), "white")

        imagen_final.paste(qr_img, (0,0))

        canvas = ImageDraw.Draw(imagen_final)

        try:
            font = ImageFont.truetype("arial.ttf", 12)
        except IOError:
            font = ImageFont.load_default()

        # Calcular la posición para que el texto quede centrado horizontalmente
        # textbbox nos da las coordenadas del rectángulo del texto: (left, top, right, bottom)
        bbox = canvas.textbbox((0, 0), texto, font=font, stroke_width=5, align="center")
        text_width = bbox[2] - bbox[0]

        x = (nuevo_ancho - text_width) // 2
        y = qr_height + 10 # 10 píxeles por debajo del QR

        # Dibujar el texto (en color negro)
        canvas.text((x, y), texto, fill="black", font=font) #aca podemos especificar la fuente, pero es opcional, ya que si queremos darle otro estilo ahi lo configuramos con la propiedad llamada font=valor, y ahi en dond dice valor iria el nombre de la tipografia


        #en esta parte creamos la funcionalidad para que se pueda generar una carpeta en el escritorio o donde sea, y almacenemos nuestros qrs
        nombre_carpeta = "material-qr-listos"
        nombre_archivo = f"{num_material}.png"
        contenido = imagen_final

        #creamos la carpeta (ponemos exist_ok=true para que verifique si ya existe esa carpeta)
        os.makedirs(nombre_carpeta, exist_ok=True)

        ruta_escritorio = Path.home() / "Desktop"

        #unimos la ruta de la carpeta con el nombre del archivo
        ruta_completa = os.path.join(nombre_carpeta, nombre_archivo)
        
       
        # Crear la imagen con colores personalizados (puedes usar nombres o códigos HEX)
        #img = qr.make_image(fill_color="darkblue", back_color="white")

        # Guardar el archivo - lo guardamos con la nomenclatura del numero de material
        imagen_final.save(ruta_completa, "png")

        imagenes_para_pdf.append(imagen_final)

        #esta linea de codigo, limpia el QR para la siguiente iteracion de la generacion de los qr, para que en cada vuelta se limpie el qr de los datos, y vaya generando los nuevos
        qr.clear()


    ##------------------------esta seccion para abajo, es para la generacion del PDF y copiando los qrs en las hojas de los pdf -------------##

    if not imagenes_para_pdf:
        print("no hay imagenes para procesar")
        return
    

    #definimos el alto y el ancho de la hoja del pdf
    ancho_hoja = 2480
    alto_hoja = 3508


    #grilla de 3 columnas x 4 filas = 12 qrs en cada hoja o por pagina
    columnas = 3
    filas = 4
    qrs_por_pagina = columnas * filas


    # Tamaño al que vamos a redimensionar cada bloque (QR+Texto) dentro del PDF
    # Al ser el QR original de aprox 300x350px, lo agrandamos un toque para la hoja
    ancho_item_pdf = 600
    alto_item_pdf = 700  # Más alto por el espacio del texto
    
    # Márgenes de distribución en la hoja
    margen_izquierdo = 180
    margen_superior = 200
    espacio_horizontal = 160
    espacio_vertical = 120

    #un arreglo de las paginas del pdf, para saber cuantas se van agregando de acuerdo a las cantidades de pdf
    paginas_pdf = []
    #nos paramos sobre la hoja actual
    hoja_actual = None

    # tomamos como primer argumento, el indice de cada qr, y como segundo la imagen qr final, y enumeramos el arreglo de donde se guardaron los qrs a ese arreglo
    for indice, img_qr_final in enumerate(imagenes_para_pdf):

        #si corresponde, crear una hoja nueva, aca lo que hace es que si el porcentaje o el resto del indice y qrs por pagina(que son 12) da 0 (osea que no sobra)
        if indice % qrs_por_pagina == 0:
            if hoja_actual:
                #si hay hoja actual, agregamos una hoja al arreglo de paginas_pdf
                paginas_pdf.append(hoja_actual)
            #y de la hoja actual creamos una nueva imagen, como que la generamos(rgb, le pasamos el ancho, alto y de que color de fondo sea, en este caso blanca)
            hoja_actual = Image.new('RGB', (ancho_hoja, alto_hoja), 'white') 

        
        #calculamos posicion en la grilla de la hoja del pdf
        indice_en_pagina = indice % qrs_por_pagina
        col = indice_en_pagina % columnas
        fila = indice_en_pagina // columnas

        #aca ya organizamos los elementos por coordenadas, en x es horizontal, mientras que y es vertical
        ##lo que hacemos es calcular el margen_izquierdo mas el tamaño de la col(que seria el qr en si, y lo multiplicamos por el ancho y espacio horizontal que hay entre columnas)
        x = margen_izquierdo + col * (ancho_item_pdf + espacio_horizontal)
        y = margen_superior + fila * (alto_item_pdf + espacio_vertical)

        #redimensionamos el bloque del qr mas el texto debajo para que no pierda la calidad
        #aca le pasamos el nuevo ancho y alto de los qrs, asi no pierden la calidad
        #el segundo argumento, llamado Image.Resampling es un algoritmo que redimensiona lo que le pasemos, recalculando los pixeles de la imagen para que se vea nitido.
        img_redimensionada = img_qr_final.resize((ancho_item_pdf, alto_item_pdf), Image.Resampling.LANCZOS)

        #pegamos el bloque en la hoja del pdf, y le pasamos como argumento la imagen_redimensionada, y las coordenadas finales que serian las variables de x e y
        hoja_actual.paste(img_redimensionada, (x,y))

    #si hay hoja actual, se la pegamos en el arreglo de paginas_pdf
    if hoja_actual:
        paginas_pdf.append(hoja_actual)

    #creamos una variable con el nombre de los qr
    archivo_pdf_salida = "qrs_listos.pdf"
    #guardamos todas las hojas procesadas en un unico archivo pdf
    if paginas_pdf:
       
       #hacemos que la ruta donde se guarde sea en el home, o desktop, y el nombre sea el valor de la variable que creamos arriba
        ruta_nuevo_escritorio = Path.home() / "Desktop" / archivo_pdf_salida

        
        paginas_pdf[0].save(
            ruta_nuevo_escritorio,
            save_all=True,
            append_images=paginas_pdf[1:]
        )
        print(f"🎉 ¡Proceso completo! PNGs guardados y PDF creado en el Escritorio: {ruta_escritorio}")

    


   

    




#funcion para poder mostrar la info del excel en una tabla en el programa
def display_data(data):
    tree.delete(*tree.get_children())

    tree["columns"] = list(data.columns)
    tree["show"] = "headings"

    for col in data.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))



#funcion para agregar un material nuevo y que genere un qr a partir de esos datos
def add_new_data_qr():
    global df
    input_num_material = codigo_material.get() #con el .get() capturamos lo que el usuario vaya escribiendo en ese input
    input_descripcion_material = descripcion_material.get()

    # Configuración del QR
    qr = qrcode.QRCode(
        version=1, # Tamaño del QR (del 1 al 40). El 1 es el más pequeño.
        error_correction=qrcode.constants.ERROR_CORRECT_L, # Tolerancia a daños o que salga medio pixeleado
        box_size=10, # Cuántos píxeles mide cada "cuadrado" del QR
        border=4, # Grosor del borde blanco (mínimo 4)
    )

    #solo vamos a agregar el dato del numero de material
    #qr.add_data(datos_completo_desc)
    qr.add_data(input_num_material)
    qr.make(fit=True)

    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    qr_width, qr_height = qr_img.size

    #configuramos el texto y el espacio extra
    texto = input_descripcion_material
    extra_pantalla_abajo = 50 #pixeles extra abajo para el texto

    #creamos un lienzo nuevo mas alto que el qr original
    nuevo_ancho = qr_width
    nuevo_alto = qr_height + extra_pantalla_abajo
    imagen_final = Image.new("RGB", (nuevo_ancho, nuevo_alto), "white")

    imagen_final.paste(qr_img, (0,0))

    canvas = ImageDraw.Draw(imagen_final)

    try:
        font = ImageFont.truetype("arial.ttf", 12)
    except IOError:
        font = ImageFont.load_default()

    # Calcular la posición para que el texto quede centrado horizontalmente
    # textbbox nos da las coordenadas del rectángulo del texto: (left, top, right, bottom)
    bbox = canvas.textbbox((0, 0), texto, font=font)
    text_width = bbox[2] - bbox[0]

    x = (nuevo_ancho - text_width) // 2
    y = qr_height + 10 # 10 píxeles por debajo del QR

    # Dibujar el texto (en color negro)
    canvas.text((x, y), texto, fill="black", font=font) #aca podemos especificar la fuente, pero es opcional, ya que si queremos darle otro estilo ahi lo configuramos con la propiedad llamada font=valor, y ahi en dond dice valor iria el nombre de la tipografia


    #en esta parte creamos la funcionalidad para que se pueda generar una carpeta en el escritorio o donde sea, y almacenemos nuestros qrs
    nombre_carpeta = "material-qr-listos"
    nombre_archivo = f"{input_num_material}.png"
    contenido = imagen_final

    #creamos la carpeta (ponemos exist_ok=true para que verifique si ya existe esa carpeta)
    os.makedirs(nombre_carpeta, exist_ok=True)

    #unimos la ruta de la carpeta con el nombre del archivo
    ruta_completa = os.path.join(nombre_carpeta, nombre_archivo)

    

    # Crear la imagen con colores personalizados (puedes usar nombres o códigos HEX)
    #img = qr.make_image(fill_color="darkblue", back_color="white")

    # Guardar el archivo - lo guardamos con la nomenclatura del numero de material
    imagen_final.save(ruta_completa, "png")

 
    



#funcion para filtrar la data del excel
def filter_data():
    global df
    column = entry_column.get()
    value = entry_value.get()

    if column not in df.columns:
        messagebox.showwarning("cuidado", "columna no encontrada")
        return
    

    filtered_df = df[df[column].astype(str) == value]

    if filtered_df.empty:
        messagebox.showinfo("info", "no machea los registros encontrados")
    
    display_data(filtered_df)



#funcion para imprimir los qr
def printqr():

    #consultamos cual archivo va a abrir para imprimir
    ruta_archivo = filedialog.askopenfilenames(
        title="Selecciona el archivo para imprimir",
        filetypes=[("Archivos PDF", "*.pdf"), ("Docs de texto", "*.txt"), ("archivos de imagen", "*.*")]
    )

    #si el usuario selecciono un archivo, hacemos esta accion
    if ruta_archivo:
      
        exitos = 0
        errores = 0
        #recorremos ya que si enviamos mas de un archivo, el startfile solo reconoce un archivo a la vez, entonces lo recorremos esas ruta y vaya tomando cada archivo por vuelta
        for ruta in ruta_archivo:
            try:
                
                #reemplazamos las barras inclinadas para evitar problemas en entornos windows
                ruta_normalizada = os.path.abspath(ruta) #retorna o quita la ruta absoluta, que es lo de las barras inclinadas, por cada ruta encontrada

                os.startfile(ruta_normalizada, "print")
                exitos += 1 #se suma uno a los exitosos

                #messagebox.showinfo("exito!", "se envio el archivo a la impresora para iniciar la impresion")
            except Exception as e:
                messagebox.showerror("Error", f"no se pudo imprimir el archivo. \nDetalles: {e}")
                errores += 1
        
        if errores == 0:
            messagebox.showinfo("Exito!", f"se enviaron {exitos} archivos a imprimir")
        else:
            messagebox.showwarning("Resultado", f"se enviaron {exitos} a imprimir, pero {errores} fallaron")
    else:
        messagebox.showwarning("Cancelado", "no se selecciono ningun archivo")
    #os.startfile("nombre del archivo a imprimir", "print")















#ahora configuramos los widgets de tkinter

#1er argumento: referencia a la app principal, 2do argumento: titulo del boton, 3er argumento: comando o accion a realizar cuando apretamos el boton, en este caso la funcion de cargar el excel
btn_load = tk.Button(root, text="Cargar excel", command=load_csv, bg="#0b5399", 
    fg="white", 
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5)
btn_load.pack(pady=10) #pack() acomoda los elementos uno abajo del otro o de lado, es sencillo de manejarlo

frame_filter = tk.Frame(root) #Un contenedor invisible para organizar otros widgets. tk.Frame(ventana) -> que se llama root en este caso
frame_filter.pack(pady=10)


#-----------------------------------------------seccion donde podemos filtrar o buscar un elemento en particular basandonos en la columna y su valor correspondiente
entry_column = tk.Entry(frame_filter, width=20, textvariable="Columna") #Un campo de texto de una sola línea (para que el usuario escriba).
entry_column.grid(row=0, column=0, padx=5)
entry_column.insert(0, "Columna") #input donde buscamos por nombre de la columna



entry_value = tk.Entry(frame_filter, width=20)
entry_value.grid(row=0, column=1, padx=5)
entry_value.insert(0, "Valor") #input donde buscamos por el valor de esa fila o elemento

btn_filter = tk.Button(frame_filter, text="Filtrar busqueda", command=filter_data, bg="#0b5399", 
    fg="white", 
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5) #el boton donde ejecuta la funcion de filtrar la data
btn_filter.grid(row=0, column=2, padx=5,pady=5)
#----------------------------------------------------------

btn_generateqr = tk.Button(frame_filter, text="Generar QRs", command=generate_qr, bg="#0b5399", 
    fg="white", 
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5)
btn_generateqr.grid(row=0, column=3, padx=5, pady=5)


#---------------------------------------seccion para la configuracion de agregar un nuevo material-------------------------------------
codigo_material = tk.Entry(frame_filter, width=20)
codigo_material.grid(row=1, column=0, padx=5)

descripcion_material = tk.Entry(frame_filter, width=20)
descripcion_material.grid(row=1, column=1, padx=5)

btn_generate_new_qr = tk.Button(frame_filter, text="generar nuevo QR", command=add_new_data_qr, bg="#0b5399", 
    fg="white", 
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5)
btn_generate_new_qr.grid(row=1, column=2, padx=5, pady=5)


##----------------------------------------------configuracion del boton para imprimir--------------------------------------------------------
btn_print = tk.Button(frame_filter, text="imprimir qrs", command=printqr, bg="#0b5399", 
    fg="white", 
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5)
btn_print.grid(row=1,column=3,padx=5, pady=5)

# Treeview to display data
tree = ttk.Treeview(root)
tree.pack(expand=True, fill='both')


root.mainloop()