# src/modelo/tabla_dinamica.py
"""
Lógica para crear Hoja2 con tabla dinámica
"""

import pandas as pd


def crear_hoja2_tabla_dinamica(wb, ws_destino, ultima_fila_datos, headers_destino, estilos, callback=None):
    """Crea Hoja2 con tabla dinámica agrupada por rangos de MONTO CREDITO"""
    try:
        # Buscar o crear Hoja2
        if 'Hoja2' in wb.sheetnames or 'HOJA2' in [s.upper() for s in wb.sheetnames]:
            for nombre in wb.sheetnames:
                if nombre.upper() == 'HOJA2':
                    wb.remove(wb[nombre])
                    break
        
        hoja2 = wb.create_sheet("Hoja2")
        
        # Buscar columna de MONTO CREDITO
        col_monto_credito = None
        for idx, cell in enumerate(headers_destino):
            if cell.value:
                header_str = str(cell.value).strip().upper()
                if 'MONTO CREDITO' in header_str or 'MONTO CRÉDITO' in header_str:
                    col_monto_credito = idx + 1
                    break
        
        if not col_monto_credito:
            if callback:
                callback("  ⚠️ No se encontró columna MONTO CREDITO para tabla dinámica")
            return
        
        # Leer datos de MONTO CREDITO
        datos_monto = []
        for fila in range(6, ultima_fila_datos + 1):
            try:
                cell = ws_destino.cell(fila, col_monto_credito)
                if cell.value is not None:
                    valor = cell.value
                    if isinstance(valor, (int, float)):
                        datos_monto.append(float(valor))
                    elif isinstance(valor, str):
                        try:
                            datos_monto.append(float(valor.replace(',', '.')))
                        except:
                            pass
            except:
                continue
        
        if not datos_monto:
            if callback:
                callback("  ⚠️ No se encontraron datos de MONTO CREDITO")
            return
        
        # Crear DataFrame y agrupar por rangos
        df = pd.DataFrame({'MONTO_CREDITO': datos_monto})
        
        def asignar_rango(valor):
            if valor <= 5000:
                return "1-5000"
            elif valor <= 10000:
                return "5001-10000"
            elif valor <= 15000:
                return "10001-15000"
            elif valor <= 20000:
                return "15001-20000"
            elif valor <= 25000:
                return "20001-25000"
            elif valor <= 30000:
                return "25001-30000"
            elif valor <= 35000:
                return "30001-35000"
            elif valor <= 40000:
                return "35001-40000"
            else:
                inicio = int((valor - 1) // 5000) * 5000 + 1
                fin = inicio + 4999
                return f"{inicio}-{fin}"
        
        df['Rango'] = df['MONTO_CREDITO'].apply(asignar_rango)
        resultado = df.groupby('Rango').agg({'MONTO_CREDITO': ['count', 'sum']}).reset_index()
        resultado.columns = ['Rango', 'Cuenta', 'Suma']
        
        # Ordenar por inicio de rango
        def obtener_inicio_rango(rango_str):
            try:
                return int(rango_str.split('-')[0])
            except:
                return 0
        
        resultado['Orden'] = resultado['Rango'].apply(obtener_inicio_rango)
        resultado = resultado.sort_values('Orden')
        resultado = resultado.drop('Orden', axis=1)
        
        # Escribir encabezados
        hoja2['A1'] = 'Etiquetas de fila'
        hoja2['B1'] = 'Cuenta de MONTO CREDITO'
        hoja2['C1'] = 'Suma de MONTO CREDITO'
        
        for col in ['A1', 'B1', 'C1']:
            cell = hoja2[col]
            cell.font = estilos.fuente_calibri
            cell.alignment = estilos.alineacion_centrada
            cell.fill = estilos.fill_gris
            cell.border = estilos.borde_celda
        
        # Escribir datos
        fila_actual = 2
        for _, row in resultado.iterrows():
            hoja2[f'A{fila_actual}'] = row['Rango']
            hoja2[f'B{fila_actual}'] = int(row['Cuenta'])
            hoja2[f'C{fila_actual}'] = round(row['Suma'], 2)
            
            for col in ['A', 'B', 'C']:
                cell = hoja2[f'{col}{fila_actual}']
                cell.font = estilos.fuente_calibri
                cell.alignment = estilos.alineacion_centrada
                cell.border = estilos.borde_celda
                if col == 'C':
                    cell.number_format = '0.00'
            
            fila_actual += 1
        
        # Fila de Total general
        total_cuenta = int(resultado['Cuenta'].sum())
        total_suma = round(resultado['Suma'].sum(), 2)
        
        hoja2[f'A{fila_actual}'] = 'Total general'
        hoja2[f'B{fila_actual}'] = total_cuenta
        hoja2[f'C{fila_actual}'] = total_suma
        
        for col in ['A', 'B', 'C']:
            cell = hoja2[f'{col}{fila_actual}']
            cell.font = estilos.fuente_calibri
            cell.alignment = estilos.alineacion_centrada
            cell.border = estilos.borde_celda
            cell.fill = estilos.fill_azul
            if col == 'C':
                cell.number_format = '0.00'
        
        # Ancho de columnas
        hoja2.column_dimensions['A'].width = 20
        hoja2.column_dimensions['B'].width = 25
        hoja2.column_dimensions['C'].width = 25
        
        if callback:
            callback("  ✓ Hoja2 creada con tabla dinámica de MONTO CREDITO")
        
    except Exception as e:
        if callback:
            callback(f"  ⚠️ Error al crear Hoja2: {str(e)}")
