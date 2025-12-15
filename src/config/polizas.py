# Configuración de Pólizas
# Este archivo contiene la configuración de todas las pólizas disponibles

CONFIGURACION_POLIZAS = {
    'DV': {
        'prefijo': 'DV',
        'nombre_archivo': 'Facturación DV',
        'patrones_hoja': [r'DV\s*\(\d+\)', r'DV\s+\d+', 'DV'],
        'descripcion': 'Póliza DV (5852 u otros)',
        'hoja_origen_requerida': 'Report_AseguradoraMensual',  # Nombre de hoja en 413
    },
    'TC': {
        'prefijo': 'TC',
        'nombre_archivo': 'Facturación TC',
        'patrones_hoja': [r'T\+C', r'TC', r'5924'],
        'descripcion': 'Póliza T+C (5924)',
        'hoja_origen_requerida': 'Report_AseguradoraSaldos_COVID',
        'numero_poliza_fijo': '5924',
        'nombre_producto_fijo': 'SALDO DE DEUDA T + C',
        'pais_residencia_fijo': '239',
    },
    # ==================================================================
    # Agregar nuevas pólizas aquí:
    # ==================================================================
    # 'RC': {
    #     'prefijo': 'RC',
    #     'nombre_archivo': 'Facturación RC',
    #     'patrones_hoja': [r'RC\s*\(\d+\)', r'RC\s+\d+', 'RC'],
    #     'descripcion': 'Póliza RC (Responsabilidad Civil)',
    #     'hoja_origen_requerida': 'Report_AseguradoraMensual',
    # },
    # 'AP': {
    #     'prefijo': 'AP',
    #     'nombre_archivo': 'Facturación AP',
    #     'patrones_hoja': [r'AP\s*\(\d+\)', r'AP\s+\d+', 'AP'],
    #     'descripcion': 'Póliza AP (Accidentes Personales)',
    #     'hoja_origen_requerida': 'Report_AseguradoraMensual',
    # },
}

# Configuración General del Sistema
CONFIG_SISTEMA = {
    'ARCHIVOS': {
        'nombre_plantilla': 'plantilla5852.xlsx',
        'hoja_codigos': 'Códigos',  # Hoja a ignorar en plantilla
    },
    'RUTAS': {
        'program_folder': 'programa',  # Carpeta relativa
        'plantillas_folder': 'src/plantillas',  # Carpeta de plantillas
    },
    'UI': {
        'ventana_titulo': 'Transformador de Excel',
        'ventana_ancho': 800,
        'ventana_alto': 700,
    },
    'PROCESAMIENTO': {
        'actualizacion_ui_cada_n_filas': 2000,
        'guardado_cada_n_filas': 3000,
        'max_mensajes_por_ciclo': 10,
        'verificacion_mensajes_ms': 50,
    },
    'VALIDACION': {
        'min_filas_obligatorio': 10,
        'min_columnas_encabezado': 5,
        'min_columnas_validas_encabezado': 3,
        'max_filas_busqueda_encabezados': 10,
        'max_filas_validacion_formulas': 5,
    },
}

# Encabezados requeridos (mappeo de nombres)
ENCABEZADOS_CRITICOS = {
    'TIPO IDENTIFICACION': ['TIPO IDENTIFICACION', 'TIPO IDENT'],
    'NUMERO IDENTIFICACION': ['NUMERO IDENTIFICACION', 'NUMERO IDENT', 'CEDULA'],
    'PRIMER APELLIDO': ['PRIMER APELLIDO'],
    'PRIMER NOMBRE': ['PRIMER NOMBRE'],
    'NACIONALIDAD': ['NACIONALIDAD'],
    'FECHA NACIMIENTO': ['FECHA NACIMIENTO', 'FECHA NACI'],
    'PAIS DE RESIDENCIA': ['PAIS DE RESIDENCIA', 'PAIS RESIDENCIA'],
    'PROVINCIA': ['PROVINCIA'],
    'CIUDAD': ['CIUDAD'],
    'NUMERO DE POLIZA': ['NUMERO DE POLIZA', 'NUMERO POLIZA', 'POLIZA'],
    'MONTO CREDITO': ['MONTO CREDITO', 'MONTO', 'CREDITO'],
    'FECHA DE INICIO DE CREDITO': ['FECHA DE INICIO DE CREDITO', 'FECHA INICIO'],
}

# Palabras clave que indican fin de datos
PALABRAS_CLAVE_TOTALES = ['TOTAL', 'CUADRE', 'PRECANCELACION', 'PRE CANCELACION', 'SUMA']

# Transformaciones especiales de datos
TRANSFORMACIONES = {
    'PAIS_RESIDENCIA_FIJO': '239',  # Valor por defecto para país
    'NACIONALIDAD_SI_TIPO_00': '239',  # Si TIPO IDENTIFICACION = '00'
    'QUITAR_CEROS_INICIALES': ['PROVINCIA', 'CIUDAD'],  # Columnas que pierden ceros
}

# Meses en español
MESES_ESPANOL = {
    1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
    5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
    9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
}
