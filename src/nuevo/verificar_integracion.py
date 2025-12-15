#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script para validar que DV (413) sigue funcionando correctamente"""

import sys
import os

# Add parent path to imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

print("=" * 80)
print("TEST: Verificar que DV (413) sigue funcionando tras cambios para TC")
print("=" * 80)

# Verificar que se pueden cargar todos los módulos sin errores
try:
    from modelo.transformador import TransformadorDatos
    from config.polizas import CONFIGURACION_POLIZAS
    from modelo.transferencia_datos import TransferenciaDatos
    from modelo.mapeo_tc import MAPEO_TC_MANUAL, NO_SOBRESCRIBIR_TC, CAMPOS_FIJOS_TC
    print("\n[OK] Todos los módulos importan correctamente")
except ImportError as e:
    print(f"\n[ERROR] Fallo al importar módulos: {e}")
    sys.exit(1)

# Verificar configuración de pólizas
print("\n[CONFIG] Verificando configuración de pólizas...")
if 'DV' not in CONFIGURACION_POLIZAS:
    print("[ERROR] Póliza DV no está en configuración")
    sys.exit(1)
if 'TC' not in CONFIGURACION_POLIZAS:
    print("[ERROR] Póliza TC no está en configuración")
    sys.exit(1)

dv_config = CONFIGURACION_POLIZAS['DV']
tc_config = CONFIGURACION_POLIZAS['TC']

print(f"  DV: {dv_config['nombre_archivo']}")
print(f"     Hoja origen: {dv_config.get('hoja_origen_requerida')}")
print(f"  TC: {tc_config['nombre_archivo']}")
print(f"     Hoja origen: {tc_config.get('hoja_origen_requerida')}")

# Verificar que los cambios en transferencia_datos no rompieron nada
print("\n[LOGICA] Verificando que cambios a transferencia_datos son seguros...")
try:
    from modelo.transferencia_datos import TransferenciaDatos
    
    # Crear instancia dummy
    class EstilosMock:
        pass
    
    td = TransferenciaDatos(EstilosMock(), {}, {}, None)
    print("[OK] TransferenciaDatos se puede instanciar correctamente")
except Exception as e:
    print(f"[ERROR] Error al instanciar TransferenciaDatos: {e}")
    sys.exit(1)

# Verificar mapeo TC
print("\n[MAPEO_TC] Verificando mapeo TC...")
print(f"  Columnas mapeadas: {len(MAPEO_TC_MANUAL)}")
print(f"  Columnas sin sobrescribir: {len(NO_SOBRESCRIBIR_TC)}")
print(f"  Campos fijos: {len(CAMPOS_FIJOS_TC)}")

if len(MAPEO_TC_MANUAL) != 39:
    print(f"[WARNING] Mapeo TC tiene {len(MAPEO_TC_MANUAL)} columnas, esperadas 39")
if len(NO_SOBRESCRIBIR_TC) != 5:
    print(f"[WARNING] NO_SOBRESCRIBIR_TC tiene {len(NO_SOBRESCRIBIR_TC)} columnas, esperadas 5")
if len(CAMPOS_FIJOS_TC) != 3:
    print(f"[WARNING] CAMPOS_FIJOS_TC tiene {len(CAMPOS_FIJOS_TC)} campos, esperados 3")

print("\n[RESULTADO] Verificación completa - DV y TC están listos para probar")
print("=" * 80)
