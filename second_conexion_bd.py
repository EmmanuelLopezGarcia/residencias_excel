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
        fechaAlta_list = list()
        precioMXN_list = list()
        cantidad_list = list()
        nueva_fecha_list = list()

        while all_rows is not None:

            date = datetime.strptime(str(all_rows[39]), '%Y-%m-%d %H:%M:%S')
            fechaAlta_list.append(date)

            if fechaAlta_list[counter] >= primera_fecha and fechaAlta_list[counter] <= segunda_fecha:

                if type(all_rows[19]) == type(None):
                    # print(str(all_rows[19]))
                    precioMXN_list.append(0)
                else:
                    # print(all_rows[19])
                    precioMXN_list.append(float(all_rows[19]))

                cantidad_list.append(all_rows[18])
                # fechaAlta_list.append(all_rows[counter])
                nueva_fecha_list.append(fechaAlta_list[counter])


            # print(all_rows[18], all_rows[19], fechaAlta_list[counter])
            # print(fechaAlta_list[counter])

            all_rows = cursor_fechas.fetchone()
            counter += 1

        dictionario["PrecioMXN"] = precioMXN_list
        dictionario["Cantidad"] = cantidad_list
        dictionario["FechaAlta"] = nueva_fecha_list

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

    # cursor.close()
    conexion.close()

    print("Conexión finalizada")
