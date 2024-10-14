from sqlalchemy import create_engine
import pandas as pd

# Configuración de la conexión
usuario = ''
contraseña = ''
servidor = 'Emily'
puerto = '1433'
base_de_datos = 'mundojoven_db'
connection_string = f'mssql+pyodbc://{usuario}:{contraseña}@{servidor}:{puerto}/{base_de_datos}?driver=ODBC Driver 17 for SQL Server'
engine = create_engine(connection_string)

# Leer datos
df = pd.read_sql('SELECT * FROM VwMJVtasOperadora', engine)

# Convertir columnas de tipo Decimal a float
for col in df.select_dtypes(include=['object']):
    df[col] = df[col].apply(lambda x: float(x) if isinstance(x, Decimal) else x)

# Exportar a Excel
df.to_excel(r'D:\Prueba_pycharm\conexión_bd\VVta.xlsx', index=False)

print("Datos exportados a Excel correctamente.")


