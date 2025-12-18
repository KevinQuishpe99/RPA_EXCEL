#!/usr/bin/env python3
"""Script para probar la transformación y ver dónde se queda"""

import sys
import os
import traceback

# Agregar el workspace al path
sys.path.insert(0, os.getcwd())

print("[TEST] Iniciando prueba de transformación...")

try:
    # Importar módulos
    from src.modelo import TransformadorDatos
    from src.config.polizas import CONFIGURACION_POLIZAS

    print("[TEST] Módulos importados correctamente")
    print(f"[TEST] Configuración de pólizas: {list(CONFIGURACION_POLIZAS.keys())}")

    # Obtener la póliza DV
    poliza_config = CONFIGURACION_POLIZAS.get('DV')
    print(f"[TEST] Configuración DV: {poliza_config}")

    # Buscar archivo de origen (usar el primero disponible)
    archivos_excel = []
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.xlsx') and ('413' in file or '455' in file):
                archivos_excel.append(os.path.join(root, file))

    if not archivos_excel:
        print("[TEST] No se encontraron archivos 413 o 455 en el workspace")
        sys.exit(1)

    archivo_origen = archivos_excel[0]
    print(f"[TEST] Archivo de origen encontrado: {archivo_origen}")

    # Buscar plantilla
    plantilla_nombre = 'plantilla5852.xlsx'
    posibles_rutas = [
        os.path.join(os.getcwd(), 'src', 'plantillas', plantilla_nombre),
        os.path.join(os.getcwd(), 'plantillas', plantilla_nombre),
        os.path.join(os.getcwd(), plantilla_nombre),
    ]

    ruta_plantilla = None
    for rp in posibles_rutas:
        print(f"[TEST] Buscando plantilla en: {rp}")
        if os.path.exists(rp):
            ruta_plantilla = rp
            break

    if not ruta_plantilla:
        print("[TEST] Plantilla no encontrada")
        sys.exit(1)

    print(f"[TEST] Plantilla encontrada: {ruta_plantilla}")

    # Crear transformador
    print("[TEST] Creando transformador...")
    def callback(msg):
        print(f"[CALLBACK] {msg}")

    transformador = TransformadorDatos(callback_mensaje=callback)
    print("[TEST] Transformador creado")

    # Ejecutar transformación
    print("[TEST] Iniciando transformación...")
    try:
        print("[TEST] Llamando a transformador.transformar()...")
        wb_resultado, nombre_descarga = transformador.transformar(
            archivo_origen=archivo_origen,
            archivo_plantilla=ruta_plantilla,
            poliza_info=poliza_config
        )
        print(f"[TEST] ✓ Transformación completada: {nombre_descarga}")
    except Exception as e:
        print(f"[TEST] ✗ Error: {str(e)}")
        traceback.print_exc()
        sys.exit(1)

    print("[TEST] Prueba finalizada exitosamente")
    
except Exception as e:
    print(f"[ERROR GENERAL] {str(e)}")
    traceback.print_exc()
    sys.exit(1)
