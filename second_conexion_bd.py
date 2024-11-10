# Importamos las librerías necesarias
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import pyodbc
from datetime import datetime

# Se aplica un try/except para la conexión a la BDD
try:

    # Se crea la variable "conexion" que contendrá el valor de la base de datos tras conectarse
    conexion = pyodbc.connect(

        # Se coloca la información de la BDD
        'Trusted_Connection=Yes; Driver={ODBC Driver 17 for SQL Server}; UID=sa; Server=Emily; Database=mundojoven_db'
    )

    # Se imprime el siguiente texto si la conexión fue correcta
    print("Conexión exitosa")

    # Se crea un diccionario que contendrá todos los valores para convertirlos en formato xlsx
    the_big_dictionary = {}

    # Se crea la función que contiene toda la lógica basandose en el filtro aplicado al campo "FechasAlta"
    def FechasAltaAndOut(dictionario):

        # Se crean los cursores que van a leer la BDD
        cursor_fechas = conexion.cursor()

        # El cursor ejecuta un select
        cursor_fechas.execute("SELECT * FROM VwMJVtasOperadora")

        # Las filas son obtenidas una por una gracias a "fetchone"
        all_rows = cursor_fechas.fetchone()

        # Se crean las variables que filtran el rango de fechas a seleccionar
        primera_fecha = datetime.strptime(input("Ingrese la primera fecha (YYYY-MM-DD): "), '%Y-%m-%d')
        segunda_fecha = datetime.strptime(input("Ingrese la segunda fecha(YYYY-MM-DD): "), '%Y-%m-%d')

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
        gran_total_mxn = list()
        gran_total_usd = list()
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
            date = datetime.strptime(str(all_rows[39]), '%Y-%m-%d %H:%M:%S')

            # Despues guardamos ese dato en la lista fechaAlta_list
            fechaAlta_list.append(date)

            # Posteriormente evaluamos cada fecha de la lista con un contador, si la fecha esta dentro del rango
            # propuesto entonces entramos en los cuatro IF's para limpiar datos numericos:
            if fechaAlta_list[counter] >= primera_fecha and fechaAlta_list[counter] <= segunda_fecha:

                # Si el valor del dato de la columna 22 de la BD es igual a None entonces se guarda en la lista
                # gran_total_mxn el valor 0
                if type(all_rows[22]) == type(None):

                    gran_total_mxn.append(0)

                # De lo contrario se guarda el valor de la columna 22 en la lista gran_total_mxn pero se transforma a
                # float
                else:
                    gran_total_mxn.append(float(all_rows[22]))

                # Si el valor del dato de la columna 28 de la BD es igual a None entonces se guarda en la lista
                # gran_total_usd el valor 0
                if type(all_rows[28]) == type(None):

                    gran_total_usd.append(0)

                # De lo contrario se guarda el valor de la columna 28 en la lista gran_total_usd pero se transforma a
                # float
                else:

                    gran_total_usd.append(float(all_rows[28]))

                # Si el valor del dato de la columna 30 de la BD es igual a None entonces se guarda en la lista
                # beneficio_usd el valor 0
                if type(all_rows[30]) == type(None):

                    beneficio_usd.append(0)

                # De lo contrario se guarda el valor de la columna 30 en la lista beneficio_usd pero se transforma a
                # float
                else:

                    beneficio_usd.append(float(all_rows[30]))

                # Si el valor del dato de la columna 31 de la BD es igual a None entonces se guarda en la lista
                # tipo_cambio el valor 0
                if type(all_rows[31]) == type(None):

                    tipo_cambio.append(0)

                # De lo contrario se guarda el valor de la columna 31 en la lista tipo_cambio pero se transforma a
                # float
                else:

                    tipo_cambio.append(float(all_rows[31]))

                # Despues de las validaciones para limpiar los datos numericos de la BD se procede a guardar cada
                # uno ce los datos de las columnas seleccionadas para cada lista creada
                negocio_list.append(all_rows[1])
                servicio_list.append(all_rows[5])
                proveedor_nombre.append(all_rows[7])
                cliente_nombre.append(all_rows[9])
                descripcion_2.append(all_rows[10])
                oficina_nombre.append(all_rows[11])
                pedido.append(all_rows[13])
                vendedor_nombre.append(all_rows[16])
                art_categoria.append(all_rows[17])
                cantidad_list.append(all_rows[18])
                categoria_cte.append(all_rows[33])
                mes.append(all_rows[34])
                anio.append(all_rows[35])
                codigo_guia.append(all_rows[36])
                referencia.append(all_rows[37])
                origen.append(all_rows[38])
                # En la lista nueva_fecha_list se guarda el dato de la lista fechaAlta que se creo al inicio,
                # esto para no tener conflicto al pasar los datos al DataFrame y al mismo bucle
                nueva_fecha_list.append(fechaAlta_list[counter])
                fecha_in.append(all_rows[40])
                fecha_out.append(all_rows[41])
                semana_alta.append(all_rows[42])
                mes_alta.append(all_rows[44])

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
        dictionario["Gran_total_mxn"] = gran_total_mxn
        dictionario["Gran_total_usd"] = gran_total_usd
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
        dictionario["Semana_Alta"] = semana_alta
        dictionario["Mes_Alta"] = mes_alta

        # Se cierra el cursor que apunta a la base de datos
        cursor_fechas.close()

        # Retornamos el diccionario
        return dictionario

    # Creamos la variable data_frame que contendra el diccionario de la funcion FechasAltaAndOut
    data_frame = pd.DataFrame(FechasAltaAndOut(the_big_dictionary))

    # Posteriormente exportamos el dataframe como archivo xlsx en la ruta indicada
    data_frame.to_excel(r'D:\Prueba_pycharm\conexión_bd\VVta.xlsx', index=False)

# En caso de que la conexión con la BD sea erronea se levanta la excepción e imrpime el error.
except Exception as e:

    print(e)

# Finalmente cerramos la conexión con la BD e imprimimos que finalizó la conexión
finally:

    conexion.close()

    print("Conexión finalizada")
