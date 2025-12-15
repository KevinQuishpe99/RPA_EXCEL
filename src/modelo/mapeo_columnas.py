# src/modelo/mapeo_columnas.py
"""
Lógica de mapeo inteligente entre columnas origen y destino
"""

import pandas as pd


def obtener_mapeo_columnas(headers_origen, headers_destino, cache_mapeo=None, cache_headers=None):
    """
    Obtiene mapeo de columnas replicando lógica del transformador original
    
    Args:
        headers_origen: Lista de headers del archivo origen
        headers_destino: Lista de celdas de headers del archivo destino
        cache_mapeo: Cache de mapeo previo (opcional)
        cache_headers: Cache de headers previos (opcional)
    
    Returns:
        dict: Mapeo {idx_origen: idx_destino}
    """
    # Usar cache si headers no cambiaron
    if cache_mapeo is not None and cache_headers is not None:
        if len(cache_headers) == len(headers_destino):
            cache_val = cache_headers[0].value if hasattr(cache_headers[0], 'value') else None
            current_val = headers_destino[0].value if hasattr(headers_destino[0], 'value') else None
            if cache_val == current_val:
                return cache_mapeo.copy()

    mapeo = {}
    nombres_destino = {}
    nombres_destino_parciales = {}

    for idx, cell in enumerate(headers_destino):
        if cell.value:
            nombre_limpio = str(cell.value).strip().upper()
            nombres_destino[nombre_limpio] = idx + 1
            nombre_sin_espacios = nombre_limpio.replace('  ', ' ')
            if nombre_sin_espacios != nombre_limpio:
                nombres_destino[nombre_sin_espacios] = idx + 1
            palabras = nombre_limpio.split()
            for palabra in palabras:
                if len(palabra) > 2:
                    if palabra not in nombres_destino_parciales:
                        nombres_destino_parciales[palabra] = []
                    nombres_destino_parciales[palabra].append((idx + 1, nombre_limpio))

    mapeos_conocidos = {
        'PRIMER APELLIDO': ['PRIMER APELLIDO'],
        'SEGUNDO APELLIDO': ['SEGUNDO APELLIDO'],
        'PRIMER NOMBRE': ['PRIMER NOMBRE'],
        'SEGUNDO NOMBRE': ['SEGUNDO NOMBRE'],
        'OFICINA': ['OFICINA'],
        'TIPO IDENTIFICACION': ['TIPO IDENTIFICACION', 'TIPO IDENTIFICACIÓN'],
        'NUMERO DE IDENTIFICACION': ['NUMERO DE IDENTIFICACION', 'NUMERO DE IDENTIFICACIÓN', 'NÚMERO DE IDENTIFICACION'],
        'FECHA DE NACIMIENTO': ['FECHA DE NACIMIENTO'],
        'SEXO/GENERO': ['SEXO/GENERO', 'SEXO', 'GENERO'],
        'ESTADO CIVIL': ['ESTADO CIVIL'],
        'NACIONALIDAD': ['NACIONALIDAD ACTUAL', 'NACIONALIDAD', 'NACIONALIDAD ACTUAL '],
        'PAIS DE ORIGEN': ['PAIS DE ORIGEN', 'PAÍS DE ORIGEN', 'PAIS DE ORIGEN '],
        'PAIS DE RESIDENCIA': ['PAIS DE RESIDENCIA', 'PAÍS DE RESIDENCIA'],
        'PROVINCIA': ['PROVINCIA', 'PROVINCIA ', ' PROVINCIA', 'PROVINCIA DE', 'PROVINCIA DEL'],
        'CIUDAD': ['CIUDAD', 'CIUDAD ', ' CIUDAD', 'CIUDAD DE', 'CIUDAD DEL'],
        'DIRECCION ': ['DIRECCION', 'DIRECCIÓN', 'DIRECCION '],
        'TELEFONO CASA': ['TELEFONO CASA', 'TELÉFONO CASA'],
        'TELEFONO TRABAJO': ['TELEFONO TRABAJO', 'TELÉFONO TRABAJO'],
        'CELULAR': ['CELULAR'],
        'DIRECCION TRABAJO': ['DIRECCION TRABAJO', 'DIRECCIÓN TRABAJO'],
        'EMAIL': ['EMAIL', 'CORREO', 'E-MAIL'],
        'OCUPACION': ['OCUPACION', 'OCUPACIÓN'],
        'ACTIVIDAD ECONOMICA': ['ACTIVIDAD ECONOMICA', 'ACTIVIDAD ECONÓMICA'],
        'INGRESOS': ['INGRESOS'],
        'PATRIMONIO': ['PATRIMONIO'],
        'MONTO CREDITO': ['MONTO CREDITO', 'MONTO CRÉDITO'],
        'SALDO ACTUAL': ['SALDO A LA FECHA', 'SALDO ACTUAL', 'SALDO FINAL'],
        'FECHA DE INICIO DE CREDITO': ['FECHA DE INICIO DE CREDITO', 'FECHA DE INICIO DE CRÉDITO'],
        'FECHA DE TERMINACION DE CREDITO': ['FECHA DE TERMINACION DE CREDITO', 'FECHA DE TERMINACIÓN DE CRÉDITO'],
        'PLAZO DE CREDITO': ['PLAZO DE CREDITO', 'PLAZO DE CRÉDITO'],
        'PRIMA NETA': ['PRIMA NETA'],
    }

    for idx_origen, header_orig in enumerate(headers_origen):
        if pd.notna(header_orig):
            header_orig_str = str(header_orig).strip().upper()
            header_orig_str_limpio = ' '.join(header_orig_str.split())

            for key_orig, posibles_dest in mapeos_conocidos.items():
                key_orig_limpio = ' '.join(key_orig.split())
                if header_orig_str_limpio == key_orig_limpio or key_orig_limpio in header_orig_str_limpio or header_orig_str_limpio in key_orig_limpio:
                    for nombre_dest in posibles_dest:
                        nombre_dest_limpio = ' '.join(nombre_dest.split())
                        if nombre_dest_limpio in nombres_destino:
                            mapeo[idx_origen] = nombres_destino[nombre_dest_limpio]
                            break
                        if nombre_dest in nombres_destino:
                            mapeo[idx_origen] = nombres_destino[nombre_dest]
                            break
                    if idx_origen in mapeo:
                        break
                elif key_orig in ['PROVINCIA', 'CIUDAD'] and key_orig in header_orig_str_limpio:
                    for nom_dest, col_idx in nombres_destino.items():
                        if key_orig in nom_dest and 'PAIS' not in nom_dest:
                            if idx_origen not in mapeo or len(nom_dest) <= len(key_orig) + 5:
                                mapeo[idx_origen] = col_idx
                            break
                    if idx_origen in mapeo:
                        break

    for idx_origen, header_orig in enumerate(headers_origen):
        if idx_origen in mapeo:
            continue
        if pd.notna(header_orig):
            header_orig_str = str(header_orig).strip().upper()
            header_orig_str_limpio = ' '.join(header_orig_str.split())

            if header_orig_str_limpio in nombres_destino:
                mapeo[idx_origen] = nombres_destino[header_orig_str_limpio]
                continue

            mejor_coincidencia = None
            mejor_puntaje = 0
            for nom_dest, col_idx in nombres_destino.items():
                if header_orig_str_limpio == nom_dest:
                    mejor_coincidencia = col_idx
                    mejor_puntaje = 100
                    break
                elif header_orig_str_limpio in nom_dest or nom_dest in header_orig_str_limpio:
                    coincidencia_len = min(len(header_orig_str_limpio), len(nom_dest))
                    if coincidencia_len > mejor_puntaje:
                        mejor_puntaje = coincidencia_len
                        mejor_coincidencia = col_idx

            if mejor_coincidencia and mejor_puntaje >= 5:
                mapeo[idx_origen] = mejor_coincidencia
                continue

            palabras_origen = [p for p in header_orig_str_limpio.split() if len(p) > 2]
            for palabra in palabras_origen:
                if palabra in nombres_destino_parciales:
                    for col_idx, nom_dest in nombres_destino_parciales[palabra]:
                        if col_idx not in mapeo.values():
                            if len(nom_dest.split()) <= 5:
                                mapeo[idx_origen] = col_idx
                                break
                    if idx_origen in mapeo:
                        break

    return mapeo
