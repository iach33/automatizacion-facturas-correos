import os
import glob
import win32com.client as win32
import pandas as pd

# Rutas de los archivos Excel
ruta_excel_proveedores = r'C:\ruta\Proveedores.xlsx'
ruta_excel_correos = r'C:\ruta\Correos.xlsx'

# Carpeta base de proveedores
ruta_base_proveedores = r'C:\ruta\proveedores' #cambiar

# Leer el archivo de Proveedores y el de Correos
df_proveedores = pd.read_excel(ruta_excel_proveedores)
df_correos = pd.read_excel(ruta_excel_correos)

# Unir ambos dataframes en la columna 'Codigo'
df_merged = df_proveedores.merge(df_correos, on='Codigo', how='left')

# Recorrer cada fila unificada
for index, row in df_merged.iterrows():
    proveedor = str(row['Proveedor']).strip()  # nombre de la carpeta del proveedor
    factura = str(row['Factura']).strip()
    enviado = str(row['Enviado']).strip() if not pd.isna(row['Enviado']) else ''

    # Reemplazar '_' por espacio para el saludo
    proveedor_legible = proveedor.replace('_', ' ')

    # Si está marcado como enviado (x), se omite
    if enviado.lower() == 'x':
        print(f"[OMITIDO] Factura '{factura}' para '{proveedor_legible}' (Enviado='x').")
        continue

    # Obtener hasta 3 correos
    posibles_correos = []
    for col_correo in ['Correo1', 'Correo2', 'Correo3']:
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

    # Buscar subcarpetas dentro de la carpeta del proveedor
    subcarpetas = [f for f in os.scandir(ruta_proveedor) if f.is_dir()]
    if not subcarpetas:
        print(f"[ERROR] No hay subcarpetas dentro de '{ruta_proveedor}'. Se omite.")
        continue

    # Encontrar la subcarpeta con la fecha de modificación más reciente
    subcarpeta_mas_reciente = max(subcarpetas, key=lambda x: x.stat().st_mtime).path

    # Buscar archivos que contengan la factura en su nombre (en la subcarpeta más reciente)
    patron_busqueda = os.path.join(subcarpeta_mas_reciente, f"*{factura}*")
    archivos_factura = glob.glob(patron_busqueda)

    if not archivos_factura:
        print(f"[INFO] No se encontraron archivos que contengan '{factura}' en '{subcarpeta_mas_reciente}'.")
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
        f"Adjunto encontrará los archivos correspondientes a la factura {factura}.\n\n"
        "Saludos cordiales."
    )

    # Adjuntar todos los archivos que coinciden con la factura
    for archivo in archivos_factura:
        mail.Attachments.Add(archivo)

    # Enviar el correo
    mail.Send()
    print(f"[ENVIADO] Factura '{factura}' para '{proveedor_legible}'. Correos: {posibles_correos}")

    # (Opcional) Marcar como enviado en el DataFrame
    # df_merged.at[index, 'Enviado'] = 'x'

# (Opcional) Guardar el Excel con la columna 'Enviado' actualizada
# df_merged.to_excel(ruta_excel_proveedores, index=False)
