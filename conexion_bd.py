from decimal import Decimal

import pandas as pd
import pyodbc

try:
    conexion = pyodbc.connect('Trusted_Connection=Yes; Driver={SQL Server}; UID=sa; Server=Emily; Database=mundojoven_db')
    print("Conexión exitosa")
    cursor = conexion.cursor()
    cursor.execute("SELECT @@version;")
    row = cursor.fetchone()
    print(row)
    cursor.execute("SELECT  * FROM VwMJVtasOperadora")
    rows = cursor.fetchall()

    df = pd.DataFrame(rows)

    for col in df.columns:

        if df[col].dtype == 'object':

            df[col] = df[col].apply(lambda x: float(x) if isinstance(x, Decimal) else x)

    new_df = df.head(20)

    new_df.to_excel(r'D:\Prueba_pycharm\conexión_bd\VVta.xlsx', sheet_name='Ventas', index=False, engine="xlsxwriter")

    #for row in range(len(rows)):

     #  print(row)


except Exception as e:
    print(e)

finally:

    conexion.close()

    print("Conexión finalizada")
