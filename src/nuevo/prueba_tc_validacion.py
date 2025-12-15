#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script para validar que los datos TC se alinearon correctamente"""

import sys
import os
import pandas as pd

# Add parent path to imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modelo.transformador import TransformadorDatos
from config.polizas import CONFIGURACION_POLIZAS

print("=" * 80)
print("VALIDACION: Transformacion TC - Verificar alineacion de datos")
print("=" * 80)

# Rutas
ruta_origen = r"E:\s1\RPA_EXCEL\src\nuevo\455 Report_AseguradoraSaldos_Unificado noviembre 2025 5924.xlsx"
ruta_plantilla = r"E:\s1\RPA_EXCEL\src\nuevo\plantillaTC.xlsx"
ruta_salida = r"E:\s1\RPA_EXCEL\src\nuevo\resultado_tc_validacion.xlsx"

# Leer datos de origen (fila 6 = primer dato)
print("\n[ORIGEN] Leyendo primeras 3 filas de datos...")
df_origen = pd.read_excel(ruta_origen, sheet_name="Report_AseguradoraSaldos_COVID", header=None)
print("  Fila 6 (primer dato) - primeras 20 columnas:")
for i in range(20):
    print(f"    Col {i:2d}: {repr(df_origen.iloc[6, i])}")

print("\n  Fila 7 (segundo dato) - primeras 20 columnas:")
for i in range(20):
    print(f"    Col {i:2d}: {repr(df_origen.iloc[7, i])}")

# Realizar transformacion
print("\n[TRANSFORM] Ejecutando transformacion TC...")
try:
    poliza_config = CONFIGURACION_POLIZAS['TC'].copy()
    poliza_info = {
        'prefijo': poliza_config['prefijo'],
        'numero': poliza_config.get('numero_poliza_fijo'),
        'nombre': poliza_config.get('nombre_producto_fijo'),
        'hoja_origen_requerida': poliza_config.get('hoja_origen_requerida'),
        'pais_residencia': poliza_config.get('pais_residencia_fijo'),
    }
    
    transformador = TransformadorDatos()
    wb_resultado, nombre_descarga = transformador.transformar(
        archivo_origen=ruta_origen,
        archivo_plantilla=ruta_plantilla,
        poliza_info=poliza_info
    )
    wb_resultado.save(ruta_salida)
    print(f"  OK: Transformacion completada")
    print(f"  Nombre: {nombre_descarga}")
    
except Exception as e:
    print(f"  ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Leer resultado
print("\n[RESULTADO] Leyendo resultado TC...")
df_resultado = pd.read_excel(ruta_salida, sheet_name="T+C (5924)", header=None)
print(f"  Dimensiones: {df_resultado.shape}")

print("\n  Fila 7 (primer dato transferido) - primeras 20 columnas:")
for i in range(20):
    val = df_resultado.iloc[7, i]
    print(f"    Col {i:2d}: {repr(val)}")

print("\n  Fila 8 (segundo dato transferido) - primeras 20 columnas:")
for i in range(20):
    val = df_resultado.iloc[8, i]
    print(f"    Col {i:2d}: {repr(val)}")

# Verificar mapeos clave
print("\n[VERIFICACION] Campos mapeados clave:")

# Los mapeos son: origen_col -> destino_col (1-based)
# 1 -> 1: PRIMER APELLIDO (origen col 1 -> destino col 1)
print(f"  PRIMER APELLIDO (origen[6,1] -> destino[7,0])")
print(f"    Origen  col 1: {repr(df_origen.iloc[6, 1])}")
print(f"    Resultado col 0: {repr(df_resultado.iloc[7, 0])}")
print(f"    Match: {df_origen.iloc[6, 1] == df_resultado.iloc[7, 0]}")

# 2 -> 2: SEGUNDO APELLIDO
print(f"\n  SEGUNDO APELLIDO (origen[6,2] -> destino[7,1])")
print(f"    Origen col 2: {repr(df_origen.iloc[6, 2])}")
print(f"    Resultado col 1: {repr(df_resultado.iloc[7, 1])}")
print(f"    Match: {df_origen.iloc[6, 2] == df_resultado.iloc[7, 1]}")

# 29 -> 35: SALDO ACTUAL -> SALDO A LA FECHA
print(f"\n  SALDO A LA FECHA (origen[6,29] -> destino[7,34])")
print(f"    Origen col 29: {repr(df_origen.iloc[6, 29])}")
print(f"    Resultado col 34: {repr(df_resultado.iloc[7, 34])}")
print(f"    Match: {df_origen.iloc[6, 29] == df_resultado.iloc[7, 34]}")

# NUMERO POLIZA (campos fijos col 57 = índice 56)
print(f"\n  NUMERO POLIZA (campo fijo col 57)")
print(f"    Resultado col 56: {repr(df_resultado.iloc[7, 56])}")
print(f"    Esperado: '5924'")
print(f"    Match: {df_resultado.iloc[7, 56] == '5924'}")

# NOMBRE PRODUCTO (campo fijo col 58 = índice 57)
print(f"\n  NOMBRE PRODUCTO (campo fijo col 58)")
print(f"    Resultado col 57: {repr(df_resultado.iloc[7, 57])}")
print(f"    Esperado: 'SALDO DE DEUDA T + C'")
print(f"    Match: {df_resultado.iloc[7, 57] == 'SALDO DE DEUDA T + C'}")

# PAIS DE RESIDENCIA (campo fijo col 15 = índice 14)
print(f"\n  PAIS DE RESIDENCIA (campo fijo col 15)")
print(f"    Resultado col 14: {repr(df_resultado.iloc[7, 14])}")
print(f"    Esperado: '239'")
print(f"    Match: {df_resultado.iloc[7, 14] == '239'}")

print("\n" + "=" * 80)
print("VALIDACION COMPLETADA")
print("=" * 80)
