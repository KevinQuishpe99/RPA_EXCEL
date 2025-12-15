# src/modelo/mapeo_tc.py
"""
Mapeo específico para flujo TC (455 -> plantilla5924)
Define exactamente qué columna de 455 va a qué columna de TC
Los índices destino son 1-based (para usar directamente con openpyxl)
"""

# Mapeo manual: {índice_columna_455_0based: índice_columna_tc_1based}
# Basado en análisis de estructura real de archivos
MAPEO_TC_MANUAL = {
    1: 1,    # PRIMER APELLIDO
    2: 2,    # SEGUNDO APELLIDO
    3: 3,    # PRIMER NOMBRE
    4: 4,    # SEGUNDO NOMBRE
    5: 5,    # OFICINA
    6: 6,    # TIPO IDENTIFICACION
    7: 7,    # NUMERO DE IDENTIFICACION
    8: 8,    # FECHA DE NACIMIENTO
    # Col 9 (EDAD) se calcula, no se mapea
    12: 12,  # SEXO/GENERO (TC col 12)
    13: 13,  # ESTADO CIVIL (TC col 13)
    14: 14,  # NACIONALIDAD (TC col 14)
    # 15: PAIS DE ORIGEN -> TC col 15 = 239 (fijo, no mapear)
    16: 17,  # CIUDAD (TC col 17)
    15: 16,  # PROVINCIA (TC col 16)
    18: 25,  # DIRECCION (TC col 25)
    19: 26,  # TELEFONO CASA (TC col 26)
    20: 27,  # TELEFONO TRABAJO (TC col 27)
    21: 28,  # CELULAR (TC col 28)
    23: 30,  # EMAIL (TC col 30)
    24: 31,  # OCUPACION (TC col 31)
    25: 32,  # ACTIVIDAD ECONOMICA (TC col 32)
    26: 33,  # INGRESOS (TC col 33)
    27: 34,  # PATRIMONIO (TC col 34)
    29: 35,  # SALDO ACTUAL (455 col 29) -> SALDO A LA FECHA (TC col 35)
    30: 37,  # FECHA DE INICIO DE CREDITO (TC col 37)
    31: 38,  # FECHA DE TERMINACION DE CREDITO (TC col 38)
    32: 39,  # PLAZO DE CREDITO (TC col 39)
    33: 40,  # PRIMA NETA (TC col 40)
    34: 41,  # IMP (TC col 41)
    35: 42,  # PRIMA TOTAL (TC col 42)
    36: 43,  # FECHA EXPIDICION PASAPORTE (TC col 43)
    37: 44,  # FECHA CADUCIDAD PASAPORTE (TC col 44)
    38: 45,  # ESTADO MIGRATORIO (TC col 45)
    40: 47,  # NUMERO PRESTAMO (TC col 47)
    41: 48,  # METODOLOGIA (TC col 48)
    42: 49,  # NOMBRE GRUPO (TC col 49)
    43: 49,  # GRUPO (también va a TC col 49, NOMBRE GRUPO)
    46: 51,  # IDENTIFICACION CONYUGE (TC col 51)
    47: 52,  # NOMBRES CONYUGE (TC col 52)
    48: 53,  # FECHA NACIMIENTO CONYUGE (TC col 53)
}

# Columnas de 455 que se saltan:
SALTAR_455 = {
    0,   # Vacía
    9,   # (vacía o sin nombre)
    10,  # (vacía o sin nombre)
    11,  # (vacía)
    17,  # DIRECCION TRABAJO (se pone en col 25 DIRECCION)
    22,  # (vacía)
    28,  # SALDO 6 MESES ATRAS (no se usa directamente)
    39,  # (vacía)
    44,  # (vacía)
    45,  # DIAS ATRASO (no va a TC)
    49,  # (final, vacía)
}

# Columnas TC que NO se deben sobrescribir (tienen fórmulas):
NO_SOBRESCRIBIR_TC = {
    8,   # EDAD (col 9) - índice 8 0-based
    9,   # EDAD (col 10) - índice 9 0-based
    10,  # % (col 11) - índice 10 0-based
    17,  # CONCATENADO (col 18) - índice 17 0-based
    35,  # SUMA ASEGURAR SDP (col 36) - índice 35 0-based
}

# Campos FIJOS para TC:
CAMPOS_FIJOS_TC = {
    14: '239',  # PAIS DE RESIDENCIA (col 15) - índice 14 0-based
    56: '5924',  # NUMERO POLIZA (col 57) - índice 56 0-based
    57: 'SALDO DE DEUDA T + C',  # NOMBRE PRODUCTO (col 58) - índice 57 0-based
}
