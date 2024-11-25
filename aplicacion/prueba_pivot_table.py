import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt


df = pd.read_excel("Archivo_ventas.xlsx")

df2 = pd.DataFrame()
df3 = pd.DataFrame()

df2[['Vendedor_Nombre', 'Semana_Alta', 'Pedido', 'Negocio', 'Descripcion2']] = df[[
    'Vendedor_Nombre', 'Semana_Alta', 'Pedido', 'Negocio', 'Descripcion2']]
df3[['Proveedor_Nombre', 'Vendedor_Nombre', 'Pedido', 'Negocio', 'Descripcion2']] = df[[
    'Proveedor_Nombre', 'Vendedor_Nombre', 'Pedido', 'Negocio', 'Descripcion2']]

with pd.ExcelWriter('Archivo_ventas.xlsx', engine='openpyxl', mode='a') as writer:

    #df2.to_excel(writer,sheet_name='Alta_FILES', index=False)
    df3.to_excel(writer,sheet_name='Proveedor', index=False)

