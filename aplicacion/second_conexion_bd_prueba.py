# Importamos las librerías necesarias
import gc
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd
import pyodbc
from datetime import datetime
import tkinter as tk
from tkinter import PhotoImage

# Se aplica un try/except para la conexión a la BDD
try:

    # Se crea la variable "conexion" que contendrá el valor de la base de datos tras conectarse
    conexion = pyodbc.connect(

        # Se coloca la información de la BDD
        'Trusted_Connection=Yes; Driver={ODBC Driver 17 for SQL Server}; UID=sa; Server=Emily; Database=mundojoven_db_2'
        #'Driver={ODBC Driver 17 for SQL Server};'
        #'Server=operadora;'
        #'UID=sa;'
        #'PWD=S5pN3_F7o;'
        #'Trusted_Connection=No;'
        #'Database=mundojoven_db;'

    )

    # Se imprime el siguiente texto si la conexión fue correcta
    print("Conexión exitosa")

    # Se crean diccionarios para FechaAlta y FechaIn que contendrán todos los valores para convertirlos en formato xlsx
    diccionario_para_alta = {}
    diccionario_para_in = {}

    # Se crea la función que contiene toda la lógica basandose en el filtro aplicado al campo "FechasAlta"
    def FechasAltaAndOut(dictionario, fecha_uno, fecha_dos):

        # Se crean los cursores que van a leer la BDD
        cursor_fechas = conexion.cursor()

        # El cursor ejecuta un select
        cursor_fechas.execute("SELECT * FROM OTSFiles_M1")

        # Las filas son obtenidas una por una gracias a "fetchone"
        all_rows = cursor_fechas.fetchone()

        # Se crean las variables que filtran el rango de fechas a seleccionar
        #primera_fecha = datetime.strptime(input("Ingrese la primera fecha (YYYY-MM-DD): "), '%Y-%m-%d')
        primera_fecha = datetime.strptime(fecha_uno, "%Y-%m-%d")
        #segunda_fecha = datetime.strptime(input("Ingrese la segunda fecha(YYYY-MM-DD): "), '%Y-%m-%d')
        segunda_fecha = datetime.strptime(fecha_dos, "%Y-%m-%d")

        # Creación de las listas en donde se almacenarán los datos de cada columna de la base de datos
        counter = 0
        negocio_list = list()
        servicio_list = list()
        proveedor_nombre = list()
        cliente_nombre = list()
        descripcion_2 = list()
        oficina_nombre = list()
        pedido = list()
        vendedor_nombre = list()
        art_categoria = list()
        cantidad_list = list()
        gran_total_mx = list()
        costo_usd = list()
        beneficio_usd = list()
        tipo_cambio = list()
        categoria_cte = list()
        mes = list()
        anio = list()
        codigo_guia = list()
        referencia = list()
        origen = list()
        fechaAlta_list = list()
        fecha_in = list()
        fecha_out = list()
        semana_alta = list()
        mes_alta = list()
        nueva_fecha_list = list()


        # Se crea un bucle en el cual se se evalua que si la fila proveniente de la base de datos no es None se repite
        # el bucle, de lo contrario se termina
        while all_rows is not None:

            # Creamos una variable date que guarda el valor de la columna 39 de la base de datos que es FechaAlta
            date = datetime.strptime(str(all_rows[38]), '%Y-%m-%d %H:%M:%S')

            # Despues guardamos ese dato en la lista fechaAlta_list
            fechaAlta_list.append(date)

            # Posteriormente evaluamos cada fecha de la lista con un contador, si la fecha esta dentro del rango
            # propuesto entonces entramos en los cuatro IF's para limpiar datos numericos:
            if fechaAlta_list[counter] >= primera_fecha and fechaAlta_list[counter] <= segunda_fecha:

                # Si el valor del dato de la columna 22 de la BD es igual a None entonces se guarda en la lista
                # gran_total_mxn el valor 0
                if type(all_rows[21]) == type(None):

                    gran_total_mx.append(0)

                # De lo contrario se guarda el valor de la columna 22 en la lista gran_total_mxn pero se transforma a
                # float
                else:
                    gran_total_mx.append(float(all_rows[21]))

                # Si el valor del dato de la columna 28 de la BD es igual a None entonces se guarda en la lista
                # costo_usd el valor 0
                if type(all_rows[28]) == type(None):

                    costo_usd.append(0)

                # De lo contrario se guarda el valor de la columna 28 en la lista gran_total_usd pero se transforma a
                # float
                else:

                    costo_usd.append(float(all_rows[28]))

                # Si el valor del dato de la columna 30 de la BD es igual a None entonces se guarda en la lista
                # beneficio_usd el valor 0
                if type(all_rows[29]) == type(None):

                    beneficio_usd.append(0)

                # De lo contrario se guarda el valor de la columna 30 en la lista beneficio_usd pero se transforma a
                # float
                else:

                    beneficio_usd.append(float(all_rows[29]))

                # Si el valor del dato de la columna 31 de la BD es igual a None entonces se guarda en la lista
                # tipo_cambio el valor 0
                if type(all_rows[30]) == type(None):

                    tipo_cambio.append(0)

                # De lo contrario se guarda el valor de la columna 31 en la lista tipo_cambio pero se transforma a
                # float
                else:

                    tipo_cambio.append(float(all_rows[30]))

                # Despues de las validaciones para limpiar los datos numericos de la BD se procede a guardar cada
                # uno ce los datos de las columnas seleccionadas para cada lista creada
                negocio_list.append(all_rows[1])
                servicio_list.append(all_rows[4])
                proveedor_nombre.append(all_rows[6])
                cliente_nombre.append(all_rows[8])
                descripcion_2.append(all_rows[9])
                oficina_nombre.append(all_rows[10])
                pedido.append(all_rows[12])
                vendedor_nombre.append(all_rows[15])
                art_categoria.append(all_rows[16])
                cantidad_list.append(all_rows[17])
                categoria_cte.append(all_rows[32])
                mes.append(all_rows[33])
                anio.append(all_rows[34])
                codigo_guia.append(all_rows[35])
                referencia.append(all_rows[36])
                origen.append(all_rows[37])
                # En la lista nueva_fecha_list se guarda el dato de la lista fechaAlta que se creo al inicio,
                # esto para no tener conflicto al pasar los datos al DataFrame y al mismo bucle
                nueva_fecha_list.append(fechaAlta_list[counter])
                fecha_in.append(all_rows[39])
                fecha_out.append(all_rows[40])
                # semana_alta.append(all_rows[42])
                # mes_alta.append(all_rows[44])

            # Al terminar el bucle se vuelve a gurardar la siguiente fila de la BD en all_rows para que continue el
            # bucle, así hasta que se terminen todas las filas
            all_rows = cursor_fechas.fetchone()

            # Se aumenta a uno el contador para el IF que funciona como filtro para fechaAlta_list
            counter += 1

        # Al terminar el bucle while se guarda en el diccionario cada lista con el nombre que tienen en la BD
        dictionario["Negocio"] = negocio_list
        dictionario["Servicio"] = servicio_list
        dictionario["Proveedor_Nombre"] = proveedor_nombre
        dictionario["Cliente_Nombre"] = cliente_nombre
        dictionario["Descripcion2"] = descripcion_2
        dictionario["Oficina_Nombre"] = oficina_nombre
        dictionario["Pedido"] = pedido
        dictionario["Vendedor_Nombre"] = vendedor_nombre
        dictionario["Art_Categoria"] = art_categoria
        dictionario["Cantidad"] = cantidad_list
        dictionario["Gran_total_mxn"] = gran_total_mx
        dictionario["Costo_usd"] = costo_usd
        dictionario["Beneficio_Usd"] = beneficio_usd
        dictionario["Tipo_Cambio"] = tipo_cambio
        dictionario["Categoria_cte"] = categoria_cte
        dictionario["Mes"] = mes
        dictionario["Anio"] = anio
        dictionario["Codigo_Guia"] = codigo_guia
        dictionario["Referencia"] = referencia
        dictionario["Origen"] = origen
        dictionario["FechaAlta"] = nueva_fecha_list
        dictionario["FechaIn"] = fecha_in
        dictionario["Fecha_Out"] = fecha_out
        # dictionario["Semana_Alta"] = semana_alta
        # dictionario["Mes_Alta"] = mes_alta

        # Se cierra el cursor que apunta a la base de datos
        cursor_fechas.close()

        # Eliminar la información de las listas para liberar memoria
        del negocio_list
        del servicio_list
        del proveedor_nombre
        del cliente_nombre
        del descripcion_2
        del oficina_nombre
        del pedido
        del vendedor_nombre
        del art_categoria
        del cantidad_list
        del gran_total_mx
        del costo_usd
        del beneficio_usd
        del tipo_cambio
        del categoria_cte
        del mes
        del anio
        del codigo_guia
        del referencia
        del origen
        del fechaAlta_list
        del fecha_in
        del fecha_out
        # del semana_alta
        # del mes_alta
        del nueva_fecha_list

        gc.collect()

        # Retornamos el diccionario
        return dictionario

    # Se crea funcion para filtrar por FechaIn
    def FechaIn(diccionario_in, fecha_uno, fecha_dos):

        # Se crean los cursores que van a leer la BDD
        cursor_fechas_in = conexion.cursor()

        # El cursor ejecuta un select
        cursor_fechas_in.execute("SELECT * FROM OTSFiles_M1")

        # Las filas son obtenidas una por una gracias a "fetchone"
        all_rows_in = cursor_fechas_in.fetchone()

        # Se crean las variables que filtran el rango de fechas in a seleccionar
        # primera_fecha = datetime.strptime(input("Ingrese la primera fecha (YYYY-MM-DD): "), '%Y-%m-%d')
        primera_fecha_in = datetime.strptime(fecha_uno, "%Y-%m-%d")
        # segunda_fecha = datetime.strptime(input("Ingrese la segunda fecha(YYYY-MM-DD): "), '%Y-%m-%d')
        segunda_fecha_in = datetime.strptime(fecha_dos, "%Y-%m-%d")

        # Creación de las listas en donde se almacenarán los datos de cada columna de la base de datos
        contador2 = 0
        servicio_list_in = list()
        proveedor_nombre_in = list()
        cliente_nombre_in = list()
        descripcion_2_in = list()
        oficina_nombre_in = list()
        pedido_in = list()
        vendedor_nombre_in = list()
        fechaIn = list()
        nueva_fecha_in = list()

        # Se crea un bucle en el cual se se evalua que si la fila proveniente de la base de datos no es None se repite
        # el bucle, de lo contrario se termina
        while all_rows_in is not None:

            # print(contador2)

            # print(str(all_rows_in[39]))

            if all_rows_in[39] is not None:

                # print(str(all_rows_in[38]))

                # Creamos una variable date que guarda el valor de la columna 39 de la base de datos que es FechaAlta
                date = datetime.strptime(str(all_rows_in[39]), '%Y-%m-%d %H:%M:%S')

                # Despues guardamos ese dato en la lista fechaAlta_list
                fechaIn.append(date)

                # Posteriormente evaluamos cada fecha de la lista con un contador, si la fecha esta dentro del rango
                # propuesto entonces entramos en los cuatro IF's para limpiar datos numericos:
                if fechaIn[contador2] >= primera_fecha_in and fechaIn[contador2] <= segunda_fecha_in:

                    servicio_list_in.append(all_rows_in[4])
                    proveedor_nombre_in.append(all_rows_in[6])
                    cliente_nombre_in.append(all_rows_in[8])
                    descripcion_2_in.append(all_rows_in[9])
                    oficina_nombre_in.append(all_rows_in[10])
                    pedido_in.append(all_rows_in[12])
                    vendedor_nombre_in.append(all_rows_in[15])
                    nueva_fecha_in.append(fechaIn[contador2])

            else:

                fechaIn.append(0)

            # Al terminar el bucle se vuelve a gurardar la siguiente fila de la BD en all_rows para que continue el
            # bucle, así hasta que se terminen todas las filas
            all_rows_in = cursor_fechas_in.fetchone()

            # Se aumenta a uno el contador para el IF que funciona como filtro para fechaAlta_list
            contador2 += 1

        diccionario_in["Servicio"] = servicio_list_in
        diccionario_in["Proveedor_Nombre"] = proveedor_nombre_in
        diccionario_in["Cliente_Nombre"] = cliente_nombre_in
        diccionario_in["Descripcion_2"] = descripcion_2_in
        diccionario_in["Oficina_Nombre"] = oficina_nombre_in
        diccionario_in["Pedido"] = pedido_in
        diccionario_in["Vendedor_Nombre"] = vendedor_nombre_in
        diccionario_in["FechaIn"] = nueva_fecha_in

        # Eliminar información de listas
        del servicio_list_in
        del proveedor_nombre_in
        del cliente_nombre_in
        del descripcion_2_in
        del oficina_nombre_in
        del pedido_in
        del vendedor_nombre_in
        del fechaIn
        del nueva_fecha_in

        # Colector de basura
        gc.collect()

        # Se cierra el cursor que apunta a la base de datos
        cursor_fechas_in.close()

        # Retornamos el diccionario
        return diccionario_in

    # Se crea la ventana principal
    ventana = tk.Tk()

    # se le añade titulo, dimensiones, icono, no ajustable y atributos
    ventana.title("Travel Shop Ventas")
    ventana.geometry("400x500+450+60")
    ventana.iconbitmap("images/travel_icon.ico")
    ventana.configure(background="lightblue")
    ventana.resizable(width=False, height=False)
    ventana.attributes("-alpha", 0.95)

    # Se crea el primer frame que contiene la imagen
    frame_for_image = tk.Frame(ventana)
    frame_for_image.pack(padx=5, pady=5)

    # Se crea la variable que contiene a la imagen y se le asigna al frame para la imagen
    travel_image = tk.PhotoImage(file="./images/travel_shop_icon.png")
    label_for_image = tk.Label(frame_for_image, image=travel_image)
    label_for_image.pack()

    # Se crea el frame para las fechas
    frame_for_dates = tk.Frame(ventana)
    frame_for_dates.pack(padx=5, pady=5)

    # Se crean el primer label con la leyenda de ingresar el rango de FechaAlta
    label_for_fecha_1 = tk.Label(frame_for_dates, text="Ingrese el rango de FechaAlta\rDel:",
                                 font=("Arial", 11, "italic"))
    label_for_fecha_1.pack()

    # Se crea el primer entry para ingresar la primera fecha
    entry_for_date_1 = tk.Entry(frame_for_dates, fg="white", bg="gray", font=("Arial", 10, "italic"))
    entry_for_date_1.pack()
    entry_for_date_1.insert(0, "YYYY-MM-DD")

    # Se crea el segundo label con la leyenda Al
    label_for_fecha_2 = tk.Label(frame_for_dates, text="Al:", font=("Arial", 11, "italic"))
    label_for_fecha_2.pack()

    # Se crea el segundo entry para ingresar la segunda fecha
    entry_for_date_2 = tk.Entry(frame_for_dates, fg="white", bg="gray", font=("Arial", 10, "italic"))
    entry_for_date_2.pack(padx=5, pady=5)
    entry_for_date_2.insert(0, "YYYY-MM-DD")

    # Se crea el frame para le boton
    frame_for_boton = tk.Frame(ventana)
    frame_for_boton.pack(padx=5, pady=5)

    # Funcion que crea los graficos
    def CreacionGrafico():

        # Se lee el archivo Excel que se crea primero para sacar la informacion
        df2 = pd.read_excel(r'./Archivos_Excel/Archivo_FechasAlta.xlsx')

        # Se crea el data frame para copiar solo cierta informacion de df2
        df_plot = pd.DataFrame()

        # Se copia la informacion en el data frame
        df_plot[['Vendedor_Nombre', 'Proveedor_Nombre', 'Pedido', 'Descripcion2', 'Negocio', 'Costo_usd']] = df2[[
            'Vendedor_Nombre', 'Proveedor_Nombre', 'Pedido', 'Descripcion2', 'Negocio', 'Costo_usd']]

        # Se crea el data frame especifico para la grafica de seguros
        df_plot_seguro = df_plot[df_plot[
            'Proveedor_Nombre'].isin(['ASSISTCARD', 'UNIVERSAL ASSISTANCE (UA ASSISTANCE S.A DE C.V)'])]

        # ***** GRAFICO SEGURO VENTAS ******

        # Agrupar las ventas por Descripcion2, sumando el Gran_total_usd
        ventas_por_descripcion = df_plot_seguro.groupby('Descripcion2')['Costo_usd'].sum().reset_index()

        # Ordenar las descripciones por el total de ventas, de mayor a menor
        ventas_por_descripcion = ventas_por_descripcion.sort_values(by='Costo_usd', ascending=False)

        # Configuración de la figura
        plt.figure(figsize=(12, 8))

        # Gráfico de barras
        sns.barplot(x='Costo_usd', y='Descripcion2', data=ventas_por_descripcion, palette='viridis')

        # Añadir título y etiquetas
        plt.title('Ventas por Descripción', fontsize=16)
        plt.xlabel('Total Costo en USD', fontsize=12)
        plt.ylabel('Descripción', fontsize=12)

        # Mostrar los valores sobre las barras
        for p in plt.gca().patches:
            plt.gca().annotate(f'{p.get_width():,.2f}',  # Formatear el número con 2 decimales
                               (p.get_width(), p.get_y() + p.get_height() / 2),
                               ha='left', va='center', fontsize=10, color='black')

        # Mostrar la gráfica
        plt.tight_layout()

        # Se guarda la grafica en formato png
        plt.savefig(r'./Graficos/ventas_seguros.png', format="png")

        # ***** GRAFICO TOP 10 PROVEEDOR VENTAS *****

        # Sumar las ventas por proveedor
        ventas_por_proveedor = df_plot.groupby('Proveedor_Nombre')['Costo_usd'].sum().reset_index()

        # Ordenar los proveedores por ventas totales en orden descendente
        ventas_por_proveedor = ventas_por_proveedor.sort_values(by='Costo_usd', ascending=False)

        # Seleccionar solo los 10 primeros proveedores con más ventas
        ventas_por_proveedor_top10 = ventas_por_proveedor.head(10)

        # Identificar el proveedor con más ventas
        max_ventas = ventas_por_proveedor_top10['Costo_usd'].max()

        # Crear una columna para resaltar al proveedor con más ventas
        ventas_por_proveedor_top10['resaltado'] = ventas_por_proveedor_top10['Costo_usd'] == max_ventas

        # Configuración de la figura
        plt.figure(figsize=(12, 8))  # Aumentar el tamaño de la figura

        # Gráfico de barras horizontal
        ax = sns.barplot(x='Costo_usd', y='Proveedor_Nombre', data=ventas_por_proveedor_top10, palette='viridis',
                         hue='resaltado', dodge=False)

        # Mostrar los valores de "Gran_total_usd" sobre las barras
        for p in ax.patches:
            ax.annotate(f'{p.get_width():,.2f}',  # Formatear el número con 2 decimales
                        (p.get_width(), p.get_y() + p.get_height() / 2),
                        ha='left', va='center', fontsize=10, color='black')

        # Ajustar el tamaño de las letras de las etiquetas
        plt.xticks(fontsize=10)  # Tamaño de las etiquetas del eje x (valores numéricos)
        plt.yticks(fontsize=10)  # Tamaño de las etiquetas del eje y (nombres de proveedores)

        # Títulos y etiquetas
        plt.title('Top 10 Proveedores con Más Ventas', fontsize=16)
        plt.xlabel('Total Costo en USD', fontsize=12)
        plt.ylabel('Proveedor', fontsize=12)

        # Hacer el fondo cuadriculado
        plt.grid(True, which='both', axis='x', linestyle='--', linewidth=0.5)

        # Guardar la gráfica como un archivo PNG
        plt.savefig(r'./Graficos/top10_ventas_proveedor.png', format="png")

        # --- GRAFICO Top 10 VENDEDORES VENTAS ---

        # Sumar las ventas por vendedor
        ventas_por_vendedor = df_plot.groupby('Vendedor_Nombre')['Costo_usd'].sum().reset_index()

        # Ordenar los vendedores por ventas totales en orden descendente
        ventas_por_vendedor = ventas_por_vendedor.sort_values(by='Costo_usd', ascending=False)

        # Seleccionar solo los 10 primeros vendedores con más ventas
        ventas_por_vendedor_top10 = ventas_por_vendedor.head(10)

        # Identificar el vendedor con más ventas
        max_ventas_vendedor = ventas_por_vendedor_top10['Costo_usd'].max()

        # Crear una columna para resaltar al vendedor con más ventas
        ventas_por_vendedor_top10['resaltado'] = ventas_por_vendedor_top10['Costo_usd'] == max_ventas_vendedor

        # Configuración de la figura para el gráfico de vendedores
        plt.figure(figsize=(12, 8))  # Aumentar el tamaño de la figura

        # Gráfico de barras horizontal para vendedores
        ax = sns.barplot(x='Costo_usd', y='Vendedor_Nombre', data=ventas_por_vendedor_top10, palette='viridis',
                         hue='resaltado', dodge=False)

        # Mostrar los valores de "Gran_total_usd" sobre las barras
        for p in ax.patches:
            ax.annotate(f'{p.get_width():,.2f}',  # Formatear el número con 2 decimales
                        (p.get_width(), p.get_y() + p.get_height() / 2),
                        ha='left', va='center', fontsize=10, color='black')

        # Ajustar el tamaño de las letras de las etiquetas
        plt.xticks(fontsize=10)  # Tamaño de las etiquetas del eje x (valores numéricos)
        plt.yticks(fontsize=10)  # Tamaño de las etiquetas del eje y (nombres de vendedores)

        # Títulos y etiquetas
        plt.title('Top 10 Vendedores con Más Ventas', fontsize=16)
        plt.xlabel('Total Costo en USD', fontsize=12)
        plt.ylabel('Vendedor', fontsize=12)

        # Hacer el fondo cuadriculado
        plt.grid(True, which='both', axis='x', linestyle='--', linewidth=0.5)

        # Guardar la gráfica como un archivo PNG
        plt.savefig(r'./Graficos/top10_ventas_vendedor.png', format="png")

        # ******* GRAFICO NEGOCIO VENTAS ******

        # Sumar las ventas por vendedor
        ventas_por_vendedor = df_plot.groupby('Negocio')['Costo_usd'].sum().reset_index()

        # Ordenar los vendedores por ventas totales en orden descendente
        ventas_por_vendedor = ventas_por_vendedor.sort_values(by='Costo_usd', ascending=False)

        # Seleccionar solo los 10 primeros vendedores con más ventas
        ventas_por_vendedor_top10 = ventas_por_vendedor.head(10)

        # Identificar el vendedor con más ventas
        max_ventas_vendedor = ventas_por_vendedor_top10['Costo_usd'].max()

        # Crear una columna para resaltar al vendedor con más ventas
        ventas_por_vendedor_top10['resaltado'] = ventas_por_vendedor_top10['Costo_usd'] == max_ventas_vendedor

        # Configuración de la figura para el gráfico de vendedores
        plt.figure(figsize=(12, 8))  # Aumentar el tamaño de la figura

        # Gráfico de barras horizontal para vendedores
        ax = sns.barplot(x='Costo_usd', y='Negocio', data=ventas_por_vendedor_top10, palette='viridis',
                         hue='resaltado', dodge=False)

        # Mostrar los valores de "Gran_total_usd" sobre las barras
        for p in ax.patches:
            ax.annotate(f'{p.get_width():,.2f}',  # Formatear el número con 2 decimales
                        (p.get_width(), p.get_y() + p.get_height() / 2),
                        ha='left', va='center', fontsize=10, color='black')

        # Ajustar el tamaño de las letras de las etiquetas
        plt.xticks(fontsize=10)  # Tamaño de las etiquetas del eje x (valores numéricos)
        plt.yticks(fontsize=6)  # Tamaño de las etiquetas del eje y (nombres de negocios)

        # Títulos y etiquetas
        plt.title('Top 10 Negocios con Más Ventas', fontsize=16)
        plt.xlabel('Total Costo en USD', fontsize=12)
        plt.ylabel('Negocio', fontsize=12)

        # Hacer el fondo cuadriculado
        plt.grid(True, which='both', axis='x', linestyle='--', linewidth=0.5)

        # Guardar la gráfica como un archivo PNG
        plt.savefig(r'./Graficos/top10_ventas_negocio.png', format="png")

        # ******* GRAFICA PROVEDOR_SEGUROS_VENTAS *********

        # Agrupar las ventas por Proveedor seguros, sumando el Gran_total_usd
        ventas_por_descripcion = df_plot_seguro.groupby('Proveedor_Nombre')['Costo_usd'].sum().reset_index()

        # Ordenar las proveedores por el total de ventas, de mayor a menor
        ventas_por_descripcion = ventas_por_descripcion.sort_values(by='Costo_usd', ascending=False)

        # Configuración de la figura
        plt.figure(figsize=(12, 7))

        # Gráfico de barras
        sns.barplot(x='Costo_usd', y='Proveedor_Nombre', data=ventas_por_descripcion, palette='viridis')

        # Añadir título y etiquetas
        plt.title('Ventas por Proveedor de seguros', fontsize=16)
        plt.xlabel('Total Costo en USD', fontsize=12)
        plt.ylabel('Proveedor', fontsize=12)

        # Mostrar los valores sobre las barras
        for p in plt.gca().patches:
            plt.gca().annotate(f'{p.get_width():,.3f}',  # Formatear el número con 2 decimales
                               (p.get_width(), p.get_y() + p.get_height() / 2),
                               ha='left', va='center', fontsize=10, color='black')

        # Mostrar la gráfica
        plt.tight_layout()

        # Se guarda la grafica en formato png
        plt.savefig(r'./Graficos/ventas_proveedor_seguros.png', format="png")

        # Se cierran los graficos
        plt.close()

        # Se elimina la informacion de los data frames para liberar memoria
        del df2
        del df_plot

        # Se llama a l garbage collector para limpiar memoria
        gc.collect()

    # Se crea la funcion que añade las pestañas al archivo Excel, este se invoca antes que la funcion de Graficas
    def AnadirPestañasExcel():

        # Se crea el data frame que guarda la informacion del archivo Excel
        df = pd.read_excel(r'./Archivos_Excel/Archivo_FechasAlta.xlsx')

        # Se crean los data frames que guardaran solo informacion precisa
        df2 = pd.DataFrame()
        df3 = pd.DataFrame()
        df4 = pd.DataFrame()
        df5 = pd.DataFrame()
        df6 = pd.DataFrame()
        df7 = pd.DataFrame()

        # Se copia la informacion en cada data frame
        df2[['Vendedor_Nombre', 'Pedido', 'Negocio', 'Descripcion2']] = df[[
            'Vendedor_Nombre', 'Pedido', 'Negocio', 'Descripcion2']]
        df3[['Proveedor_Nombre', 'Vendedor_Nombre', 'Pedido', 'Negocio', 'Descripcion2']] = df[[
            'Proveedor_Nombre', 'Vendedor_Nombre', 'Pedido', 'Negocio', 'Descripcion2']]
        df4[['Descripcion2', 'FechaAlta', 'Proveedor_Nombre', 'Vendedor_Nombre', 'Pedido']] = df[[
            'Descripcion2', 'FechaAlta', 'Proveedor_Nombre', 'Vendedor_Nombre', 'Pedido']]
        df5[['Vendedor_Nombre', 'Descripcion2', 'Pedido', 'Negocio']] = df[[
            'Vendedor_Nombre', 'Descripcion2', 'Pedido', 'Negocio']]
        df6[['Vendedor_Nombre', 'Proveedor_Nombre', 'Pedido', 'Descripcion2', 'Negocio', 'Costo_usd']] = df[[
            'Vendedor_Nombre', 'Proveedor_Nombre', 'Pedido', 'Descripcion2', 'Negocio', 'Costo_usd']]
        df7[['Cliente_Nombre', 'Vendedor_Nombre', 'Pedido']] = df[['Cliente_Nombre', 'Vendedor_Nombre', 'Pedido']]

        # Data frame especifico con filtro para solo mostrar los seguros
        df_seguro = df6[df6['Proveedor_Nombre'].isin(['ASSISTCARD', 'UNIVERSAL ASSISTANCE (UA ASSISTANCE S.A DE C.V)'])]

        # Se escriben las pestañas en el Archivo Excel
        with pd.ExcelWriter(r'./Archivos_Excel/Archivo_FechasAlta.xlsx',
                            engine='openpyxl', mode='a') as writer:

            df2.to_excel(writer, sheet_name='Alta Files', index=False)
            df3.to_excel(writer, sheet_name='Proveedor', index=False)
            df4.to_excel(writer, sheet_name='Producto', index=False)
            df5.to_excel(writer, sheet_name='Asesor', index=False)
            df6.to_excel(writer, sheet_name='Importe', index=False)
            df7.to_excel(writer, sheet_name='Files Alta', index=False)
            df_seguro.to_excel(writer, sheet_name='Seguros', index=False)

        # Se elimina el contenido de los dataframes
        del df
        del df2
        del df3
        del df4
        del df5
        del df6
        del df7
        del df_seguro

        # Se limpia la memoria
        gc.collect()

    # Se crea la funcion que creara el excel al apretar el boton
    def CrearExcelAlta():

        # Se guardan los valores de las fechas
        date_one = entry_for_date_1.get()
        date_two = entry_for_date_2.get()

        # Creamos la variable data_frame que contendra el diccionario de la funcion FechasAltaAndOut
        data_frame = pd.DataFrame(FechasAltaAndOut(diccionario_para_alta, date_one, date_two))

        # Posteriormente exportamos el dataframe como archivo xlsx en la ruta indicada
        data_frame.to_excel(r'./Archivos_Excel/Archivo_FechasAlta.xlsx', index=False,
                            sheet_name='Datos')

        # Se elimina el contenido del data frame
        del data_frame

        # Se elimina el contenido del diccionario
        diccionario_para_alta.clear()

        # Se Ejecuta la funcion AnadirPestanasExcel en 100 milisegundos y posteriormente CreacionGrafico en 150 milis.
        ventana.after(100, AnadirPestañasExcel())
        ventana.after(150, CreacionGrafico())

        # Se libera la memoria
        gc.collect()

    # Se crea la funcion que creara el excel fechasIn al apretar el boton de fechasIn
    def CrearExcelIn():

        # Se guardan los valores de las fechas in
        date_one_in = entry_for_date_in_1.get()
        date_two_in = entry_for_date_in_2.get()

        # Creamos la variable data_frame que contendra el diccionario de la funcion FechasIn
        data_frame_in = pd.DataFrame(FechaIn(diccionario_para_in, date_one_in, date_two_in))

        # Posteriormente exportamos el dataframe como archivo xlsx en la ruta indicada
        data_frame_in.to_excel(r'./Archivos_Excel/Archivo_FechasIn.xlsx', index=False,
                               sheet_name='Datos')

        # Se elimina el data frame
        del data_frame_in

        # Se elimina el contenido del diccionario
        diccionario_para_in.clear()

        # Se libera la memoria
        gc.collect()

    # Se crea el boton para crear archivo Excel
    # Se crea el evento al presionar el boton, mismo que llama a la funcion para crear el excel
    boton = tk.Button(frame_for_boton, text="Precione para generar Excel por FechaAlta!", font=("Arial", 10, "italic"),
                      command=CrearExcelAlta)

    # Se incluye el boton en la ventana
    boton.config(fg="white", bg="green")
    boton.pack()

    # Se crea el frame para las fechas
    frame_for_dates_in = tk.Frame(ventana)
    frame_for_dates_in.pack(padx=5, pady=5)

    # Se crean el primer label con la leyenda de ingresar el rango de FechaIn
    label_for_fecha_in_1 = tk.Label(frame_for_dates_in, text="Ingrese el rango de FechaIn\rDel:",
                                 font=("Arial", 11, "italic"))
    label_for_fecha_in_1.pack()

    # Se crea el primer entry para ingresar la primera fecha in
    entry_for_date_in_1 = tk.Entry(frame_for_dates_in, fg="white", bg="gray", font=("Arial", 10, "italic"))
    entry_for_date_in_1.pack()
    entry_for_date_in_1.insert(0, "YYYY-MM-DD")

    # Se crea el segundo label con la leyenda Al para fechas in
    label_for_fecha_in_2 = tk.Label(frame_for_dates_in, text="Al:", font=("Arial", 11, "italic"))
    label_for_fecha_in_2.pack()

    # Se crea el segundo entry para ingresar la segunda fecha in
    entry_for_date_in_2 = tk.Entry(frame_for_dates_in, fg="white", bg="gray", font=("Arial", 10, "italic"))
    entry_for_date_in_2.pack(padx=5, pady=5)
    entry_for_date_in_2.insert(0, "YYYY-MM-DD")

    # Se crea el frame para le boton
    frame_for_boton_in = tk.Frame(ventana)
    frame_for_boton_in.pack(padx=5, pady=5)

    # Se crea el boton para crear archivo Excel
    # Se crea el evento al presionar el boton, mismo que llama a la funcion para crear el excel
    boton_in = tk.Button(frame_for_boton_in, text="Precione para generar Excel por FechaIn!",
                         font=("Arial", 10, "italic"), command=CrearExcelIn)

    # Se incluye el boton en la ventana
    boton_in.config(fg="white", bg="blue")
    boton_in.pack()

    # Se cierra la ventana grafica
    ventana.mainloop()

# En caso de que la conexión con la BD sea erronea se levanta la excepción e imrpime el error.
except Exception as e:

    print(e)

# Finalmente cerramos la conexión con la BD e imprimimos que finalizó la conexión

finally:

    conexion.close()

    print("Conexión finalizada")





