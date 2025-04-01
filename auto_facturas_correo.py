import os
import glob
import win32com.client as win32
import pandas as pd
import json
from datetime import datetime

# Cargar la configuración desde el archivo JSON
with open('config.json', 'r') as f:
    rutas = json.load(f)

# Rutas de los archivos Excel
ruta_excel_proveedores = rutas.get("ruta_excel_proveedores")
ruta_excel_correos = rutas.get("ruta_excel_correos")

# Carpeta base de proveedores
ruta_base_proveedores = rutas.get("ruta_base_proveedores")
ruta_a_exportar = rutas.get("ruta_a_exportar")
#Correos a copiar (CC):
correos = ["hola@gmail.com"]

# Leer el archivo de Proveedores y el de Correos
df_proveedores = pd.read_excel(ruta_excel_proveedores, parse_dates=['Fecha'])
df_proveedores['Fecha'] = pd.to_datetime(df_proveedores['Fecha'], dayfirst=True)

df_correos = pd.read_excel(ruta_excel_correos)
df_correos_reducido = df_correos[['Codigo', 'Correo1', 'Correo2', 'Correo3','Correo4']]

# Unir ambos dataframes en la columna 'Codigo'
df_merged = df_proveedores.merge(df_correos_reducido, on='Codigo', how='left')
print(df_merged.head(5))

# Recorrer cada fila unificada
for index, row in df_merged.iterrows():
    proveedor = str(row['Proveedor']).strip()  # nombre de la carpeta del proveedor
    factura = str(row['Factura']).strip()
    factura = factura[3:]
    enviado = str(row['Enviado']).strip() if not pd.isna(row['Enviado']) else ''
    
    # Convertir la fecha del Excel (dd/mm/yyyy) al formato requerido (dd-mm-yyyy)
    if pd.isna(row['Fecha']):
        print(f"[ERROR] No se encontró la fecha para el proveedor '{proveedor}'. Se omite.")
        continue

    fecha_obj = row['Fecha']
    carpeta_fecha = fecha_obj.strftime('%d-%m-%Y')

    # Reemplazar '_' por espacio para el saludo
    proveedor_legible = proveedor.replace('_', ' ')

    # Si está marcado como enviado (x), se omite
    if enviado.lower() == 'x':
        print(f"[OMITIDO] Factura '{factura}' para '{proveedor_legible}' (Enviado='x').")
        continue

    # Obtener hasta 3 correos
    posibles_correos = []
    for col_correo in ['Correo1', 'Correo2', 'Correo3','Correo4']:
        if col_correo in row and not pd.isna(row[col_correo]):
            correo_limpio = str(row[col_correo]).strip()
            if correo_limpio:
                posibles_correos.append(correo_limpio)

    # Si no hay correos válidos, no se puede enviar nada
    if not posibles_correos:
        print(f"[ERROR] No hay correos válidos para el proveedor '{proveedor_legible}' (Código={row['Codigo']}).")
        continue

    # Construir la ruta de la carpeta del proveedor
    ruta_proveedor = os.path.join(ruta_base_proveedores, proveedor)
    if not os.path.exists(ruta_proveedor):
        print(f"[ERROR] La carpeta para el proveedor '{proveedor}' no existe. Se omite.")
        continue

    # Construir la ruta de la subcarpeta que coincide con la fecha
    ruta_subcarpeta = os.path.join(ruta_proveedor, carpeta_fecha)
    if not os.path.exists(ruta_subcarpeta):
        print(f"[ERROR] No se encontró la carpeta de fecha '{carpeta_fecha}' en '{ruta_proveedor}'. Se omite.")
        continue

    # Buscar archivos que contengan la factura en su nombre (en la subcarpeta de fecha)
    patron_busqueda = os.path.join(ruta_subcarpeta, f"*{factura}*")
    archivos_factura = glob.glob(patron_busqueda)

    if not archivos_factura:
        print(f"[INFO] No se encontraron archivos que contengan '{factura}' en '{ruta_subcarpeta}'.")
        continue

    # Crear objeto Outlook
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # 0 = nuevo correo

    # Colocar los destinatarios (separados por punto y coma)
    mail.To = ";".join(posibles_correos)

    # Configurar asunto y cuerpo
    mail.Subject = f"Envío de Factura {factura}"
    mail.Body = (
        f"Estimado(a) {proveedor_legible},\n\n"
        f"Adjunto encontrará los archivos correspondientes a la factura {factura} del día {carpeta_fecha}.\n\n"
        "Saludos cordiales."
    )

    # Adjuntar todos los archivos que coinciden con la factura
    for archivo in archivos_factura:
        mail.Attachments.Add(archivo)

    # Enviar el correo
    mail.Send()
    print(f"[ENVIADO] Factura '{factura}' para '{proveedor_legible}' en carpeta '{carpeta_fecha}'. Correos: {posibles_correos}")

    # (Opcional) Marcar como enviado en el DataFrame
    df_merged['Enviado'] = df_merged['Enviado'].astype(object)
    df_merged.at[index, 'Enviado'] = 'x'

    cols = list(df_proveedores.columns) #solo las columnas de interes
    #cols.remove('Enviado')
    #cols.append('Enviado')
    df_merged = df_merged[cols]


# Guardar el DataFrame actualizado en una ruta y con un nombre que incluya la fecha actual en formato yyyymmdd
fecha_actual = datetime.now().strftime("%Y%m%d")
nombre_archivo_exportado = f"Proveedores_enviados_{fecha_actual}.xlsx"
ruta_completa_exportacion = os.path.join(ruta_a_exportar, nombre_archivo_exportado)

df_merged.to_excel(ruta_completa_exportacion, index=False)
print(f"Excel exportado en: {ruta_completa_exportacion}")