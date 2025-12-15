#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script para debuguear la transferencia TC - compatible con Windows Terminal"""

import sys
import os
import pandas as pd
import openpyxl

# Add parent path to imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modelo.transformador import TransformadorDatos
from modelo.mapeo_tc import MAPEO_TC_MANUAL, NO_SOBRESCRIBIR_TC, CAMPOS_FIJOS_TC
from config.polizas import CONFIGURACION_POLIZAS

print("=" * 80)
print("DEBUG: Transformacion TC con mapeo mejorado")
print("=" * 80)

# Rutas
ruta_origen = r"E:\s1\RPA_EXCEL\src\nuevo\455 Report_AseguradoraSaldos_Unificado noviembre 2025 5924.xlsx"
ruta_plantilla = r"E:\s1\RPA_EXCEL\src\nuevo\plantillaTC.xlsx"
ruta_salida = r"E:\s1\RPA_EXCEL\src\nuevo\resultado_tc_debug.xlsx"

# Verificar archivos
print("\n[INFO] Verificando archivos...")
print(f"  Origen: {os.path.exists(ruta_origen)}")
print(f"  Plantilla: {os.path.exists(ruta_plantilla)}")

# Leer origen
print("\n[INFO] Leyendo archivo origen...")
try:
    df_origen = pd.read_excel(ruta_origen, sheet_name="Report_AseguradoraSaldos_COVID", header=None)
    print(f"  OK: {df_origen.shape[0]} filas, {df_origen.shape[1]} columnas")
    print(f"  Primera fila (headers esperados):")
    print(f"    {list(df_origen.iloc[5, :10])}")  # Row 5 (0-based) = headers
    print(f"  Segunda fila (primer dato):")
    print(f"    {list(df_origen.iloc[6, :10])}")  # Row 6 (0-based) = first data
except Exception as e:
    print(f"  ERROR: {e}")
    sys.exit(1)

# Leer plantilla
print("\n[INFO] Leyendo plantilla TC...")
try:
    df_plantilla = pd.read_excel(ruta_plantilla, sheet_name="T+C (5924)", header=None)
    print(f"  OK: {df_plantilla.shape[0]} filas, {df_plantilla.shape[1]} columnas")
    print(f"  Headers esperados en fila 6 (0-based):")
    print(f"    {list(df_plantilla.iloc[6, :5])}")
except Exception as e:
    print(f"  ERROR: {e}")
    sys.exit(1)

# Verificar mapeo
print("\n[INFO] Verificando mapeo TC...")
print(f"  Mapeo: {len(MAPEO_TC_MANUAL)} columnas")
print(f"  NO_SOBRESCRIBIR_TC: {NO_SOBRESCRIBIR_TC}")
print(f"  CAMPOS_FIJOS_TC: {CAMPOS_FIJOS_TC}")

# Ver columna 0 del origen (primera columna)
print("\n[DEBUG] Columna 0 del archivo origen (primeros 20 valores):")
col0_values = df_origen.iloc[:20, 0].tolist()
for i, val in enumerate(col0_values):
    print(f"  Fila {i}: {repr(val)}")

# Ver si hay NaN o valores vacios
print(f"\n[DEBUG] Analisis columna 0:")
print(f"  Total non-null: {df_origen.iloc[:, 0].notna().sum()}")
print(f"  Total NaN: {df_origen.iloc[:, 0].isna().sum()}")

# Intentar transferencia manual
print("\n[INFO] Intentando transferencia manual...")
try:
    # Usar el transformador con configuración de TC
    poliza_config = CONFIGURACION_POLIZAS['TC'].copy()
    poliza_info = {
        'prefijo': poliza_config['prefijo'],
        'numero': poliza_config.get('numero_poliza_fijo'),
        'nombre': poliza_config.get('nombre_producto_fijo'),
        'hoja_origen_requerida': poliza_config.get('hoja_origen_requerida'),
        'pais_residencia': poliza_config.get('pais_residencia_fijo'),
    }
    
    print(f"  Poliza info: {poliza_info}")
    
    transformador = TransformadorDatos()
    wb_resultado, nombre_descarga = transformador.transformar(
        archivo_origen=ruta_origen,
        archivo_plantilla=ruta_plantilla,
        poliza_info=poliza_info
    )
    
    print(f"  [OK] Transformacion completada")
    print(f"  Nombre descarga: {nombre_descarga}")
    
    # Guardar resultado
    wb_resultado.save(ruta_salida)
    
    # Verificar resultado
    df_resultado = pd.read_excel(ruta_salida, sheet_name="T+C (5924)", header=None)
    print(f"\n[RESULTADO] Archivo de salida:")
    print(f"  Dimensiones: {df_resultado.shape}")
    print(f"  Fila 7 (primer dato), primeros 10 valores:")
    row7_values = df_resultado.iloc[7, :10].tolist()
    for i, val in enumerate(row7_values):
        print(f"    Col {i}: {repr(val)}")
    
except Exception as e:
    print(f"  ERROR: {e}")
    import traceback
    traceback.print_exc()
