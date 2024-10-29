import numpy as np
import pandas as pd
import pyodbc
from datetime import datetime

try:

    conexion = pyodbc.connect(
        'Trusted_Connection=Yes; Driver={ODBC Driver 17 for SQL Server}; UID=sa; Server=Emily; Database=mundojoven_db')
    print("Conexión exitosa")

    # cursor = conexion.cursor()
    # cursor.execute("SELECT  * FROM VwMJVtasOperadora ")
    the_big_dictionary = {}

    def FechasAltaAndOut(dictionario):

        cursor_fechas = conexion.cursor()
        cursor_fechas.execute("SELECT * FROM VwMJVtasOperadora")

        all_rows = cursor_fechas.fetchone()

        primera_fecha = datetime.strptime(input("Ingrese la primera fecha (YYYY-MM-DD): "), '%Y-%m-%d')
        segunda_fecha = datetime.strptime(input("Ingrese la segunda fecha(YYYY-MM-DD): "), '%Y-%m-%d')

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
        precioMXN_list = list()
        nueva_fecha_list = list()

        while all_rows is not None:

            date = datetime.strptime(str(all_rows[39]), '%Y-%m-%d %H:%M:%S')
            fechaAlta_list.append(date)

            if fechaAlta_list[counter] >= primera_fecha and fechaAlta_list[counter] <= segunda_fecha:

                if type(all_rows[22]) == type(None):
                    # print(str(all_rows[19]))
                    gran_total_mxn.append(0)

                else:
                    # print(all_rows[19])
                    gran_total_mxn.append(float(all_rows[22]))

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
                #gran_total_mxn.append(all_rows[22])
                gran_total_usd.append(all_rows[28])
                beneficio_usd.append(all_rows[30])
                tipo_cambio.append(all_rows[31])
                categoria_cte.append(all_rows[33])
                mes.append(all_rows[34])
                anio.append(all_rows[35])
                codigo_guia.append(all_rows[36])
                referencia.append(all_rows[37])
                origen.append(all_rows[38])
                nueva_fecha_list.append(fechaAlta_list[counter])
                fecha_in.append(all_rows[40])
                fecha_out.append(all_rows[41])
                semana_alta.append(all_rows[45])
                mes_alta.append(all_rows[47])
                # fechaAlta_list.append(all_rows[counter])


            # print(all_rows[18], all_rows[19], fechaAlta_list[counter])
            # print(fechaAlta_list[counter])

            all_rows = cursor_fechas.fetchone()
            counter += 1

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

        cursor_fechas.close()

        return dictionario

    """
    def PrecioMXN():

        cursor_precios = conexion.cursor()
        cursor_precios.execute("SELECT * FROM VwMJVtasOperadora")

        all_rows = cursor_precios.fetchone()

        precioMXN_list = list()

        counter = 0

        while all_rows is not None:

            if type(all_rows[19]) == type(None):

                # print(str(all_rows[19]))
                precioMXN_list.append(0)
                all_rows = cursor_precios.fetchone()
                counter += 1
            else:

                # print(all_rows[19])
                precioMXN_list.append(float(all_rows[19]))
                all_rows = cursor_precios.fetchone()
                counter += 1

        cursor_precios.close()

        precioMXN_list.append(counter)

        return precioMXN_list
        
        """


    data_frame = pd.DataFrame(FechasAltaAndOut(the_big_dictionary))

    data_frame.to_excel(r'D:\Prueba_pycharm\conexión_bd\VVta.xlsx', index=False)

except Exception as e:

    print(e)

finally:

    conexion.close()

    print("Conexión finalizada")
