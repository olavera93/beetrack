import pandas as pd
from datetime import datetime

def obtener_datos_ubicacion(ubicacion, archivo_guias, archivo_direcciones):
    """
    Obtiene el próximo número de guía, iniciales y dirección para una ubicación específica.
    Incrementa el consecutivo en el archivo después de asignarlo.
    """
    # Leer la primera hoja (consecutivos e iniciales)
    guias_df = pd.read_excel(archivo_guias, sheet_name=0)
    
    # Verificar si la ubicación existe en la primera hoja
    if ubicacion not in guias_df['ubicacion'].values:
        raise ValueError(f"La ubicación '{ubicacion}' no está registrada en el archivo de guías.")
    
    # Obtener el índice de la fila correspondiente
    idx = guias_df[guias_df['ubicacion'] == ubicacion].index[0]
    
    # Obtener el consecutivo actual e iniciales
    numero_guia = guias_df.loc[idx, 'consecutivo']
    iniciales = guias_df.loc[idx, 'iniciales']
    
    # Incrementar el consecutivo
    guias_df.loc[idx, 'consecutivo'] += 1
    
    # Guardar los cambios en el archivo de guías
    with pd.ExcelWriter(archivo_guias, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        guias_df.to_excel(writer, sheet_name='Hoja1', index=False)
    
    # Leer el archivo de direcciones
    direcciones_df = pd.read_excel(archivo_direcciones)
    
    # Normalizar las columnas de la segunda hoja
    direcciones_df.columns = direcciones_df.columns.str.strip().str.lower()
    
    # Verificar si la columna 'direccion' existe
    if 'direccion' not in direcciones_df.columns:
        raise ValueError("La columna 'direccion' no existe en el archivo de direcciones.")
    
    # Verificar si la ubicación existe en el archivo de direcciones
    if ubicacion not in direcciones_df['ubicacion'].values:
        raise ValueError(f"La ubicación '{ubicacion}' no está registrada en el archivo de direcciones.")
    
    # Obtener la dirección correspondiente
    direccion = direcciones_df.loc[direcciones_df['ubicacion'] == ubicacion, 'direccion'].values[0]
    
    # Concatenar iniciales con el número de guía
    numero_guia_final = f"{iniciales}{numero_guia}"
    
    return numero_guia_final, direccion


def procesar_archivo(input_file, output_name, placa_vehiculo, last_route, route_choice):
    archivo_guias = 'templates/guias.xlsx'  # Ruta corregida para el archivo de guías
    archivo_direcciones = 'templates/direcciones.xlsx'  # Ruta para el archivo de direcciones
    hoja = 'Sheet1'
    hora_actual = datetime.now().strftime('%Y/%m/%d %H:%M')

    # Carga el archivo Excel existente
    df = pd.read_excel(input_file, sheet_name=hoja)

    # Definir los nuevos encabezados
    nuevos_encabezados = ['guia', 'producto', 'cantidad', 'identificacion', 'contacto', 'direccion', 'horafinal', 'prioridad']
    if len(nuevos_encabezados) != len(df.columns):
        print("Error: El tamaño del array no coincide con el número de columnas en el DataFrame.")
        return None
    df.columns = nuevos_encabezados

    df['vehiculo'] = ''
    df['codigop'] = ''
    df['telefono'] = ''
    df['email'] = ''
    df['latitud'] = ''
    df['longitud'] = ''
    df['horainicio'] = ''
    df['ctdestino'] = ''
    columnas = ['guia', 'vehiculo', 'producto', 'cantidad', 'codigop', 'identificacion', 'contacto', 'telefono', 'email', 'direccion', 'latitud', 'longitud', 'horainicio', 'horafinal', 'ctdestino', 'prioridad']
    df = df[columnas]

    def extraer_antes_coma(texto):
        return texto.split(',')[0] if pd.notna(texto) else texto

    df['contacto'] = df['contacto'].apply(extraer_antes_coma)
    df['horafinal'] = pd.to_datetime(df['horafinal'], errors='coerce')
    df['horafinal'] = df['horafinal'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M') if pd.notna(x) else x)
    df['horainicio'] = hora_actual
    ultima_fila = df.dropna(how='all').index[-1]
    df.loc[:ultima_fila, 'prioridad'] = df.loc[:ultima_fila, 'prioridad'].replace('', 'Normal')
    df.loc[:ultima_fila, 'prioridad'] = df.loc[:ultima_fila, 'prioridad'].fillna('Normal')
    df['vehiculo'] = placa_vehiculo.upper()

    if not last_route:
        # Obtener número de guía y dirección
        numero_guia, direccion = obtener_datos_ubicacion(route_choice, archivo_guias, archivo_direcciones)
        
        # Crear la nueva fila
        nueva_fila = [
            numero_guia,  # Número de guía generado
            placa_vehiculo.upper(),
            f'bodega{route_choice}',
            '1', '', '', 'LFH', '', '',
            direccion, '', '', hora_actual, hora_actual, '', 'Normal'
        ]
        nueva_fila_df = pd.DataFrame([nueva_fila], columns=df.columns)
        df = pd.concat([df, nueva_fila_df], ignore_index=True)

    encabezados_df = pd.read_excel(f'templates/plantilla.xlsx', sheet_name="Formato original")
    nuevos_encabezados = encabezados_df.columns.tolist()
    if len(nuevos_encabezados) != len(df.columns):
        print("Error: El tamaño del array de encabezados no coincide con el número de columnas en el DataFrame.")
        return None
    df.columns = nuevos_encabezados

    output_file = f'uploads/{output_name}.xlsx'
    df.to_excel(output_file, sheet_name=hoja, index=False)
    return output_file
