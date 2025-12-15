"""
Prueba final del flujo TC mejorado
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

import pandas as pd
from openpyxl import load_workbook
from src.modelo.transformador import TransformadorDatos
from src.config.polizas import CONFIGURACION_POLIZAS

print("="*80)
print("PRUEBA FINAL: Transformación TC con mapeo mejorado")
print("="*80)

# Archivos
archivo_455 = "455 Report_AseguradoraSaldos_Unificado noviembre 2025 5924.xlsx"
plantilla_tc = "../plantillas/plantilla5924.xlsx"

# Obtener config TC
poliza_tc = CONFIGURACION_POLIZAS.get('TC')

print(f"\n✓ Configuración TC:")
print(f"  - Número póliza fijo: {poliza_tc.get('numero_poliza_fijo')}")
print(f"  - Nombre producto: {poliza_tc.get('nombre_producto_fijo')}")
print(f"  - País residencia: {poliza_tc.get('pais_residencia_fijo')}")

print(f"\n✓ Archivos:")
print(f"  - Origen: {archivo_455}")
print(f"  - Plantilla: {plantilla_tc}")

# Ejecutar transformación
try:
    print(f"\n{'='*80}")
    print("TRANSFORMANDO...")
    print(f"{'='*80}\n")
    
    def log_msg(msg):
        if msg.strip():
            print(f"  {msg}")
    
    transformador = TransformadorDatos(callback_mensaje=log_msg)
    
    wb_resultado, nombre_descarga = transformador.transformar(
        archivo_origen=archivo_455,
        archivo_plantilla=plantilla_tc,
        poliza_info=poliza_tc
    )
    
    print(f"\n{'='*80}")
    print("GUARDANDO RESULTADO...")
    print(f"{'='*80}\n")
    
    # Guardar temporal
    archivo_salida = "resultado_tc_mejorado.xlsx"
    wb_resultado.save(archivo_salida)
    print(f"✓ Guardado en: {archivo_salida}")
    
    # Verificar resultado
    print(f"\n{'='*80}")
    print("VERIFICACIÓN DE RESULTADO")
    print(f"{'='*80}\n")
    
    wb_check = load_workbook(archivo_salida, data_only=True)
    ws_check = wb_check.active
    
    # Encontrar headers
    fila_headers = None
    for idx in range(1, 15):
        row_text = ' '.join([str(c.value or '').upper() for c in ws_check[idx]])
        if "PRIMER APELLIDO" in row_text:
            fila_headers = idx
            break
    
    if fila_headers:
        headers = [c.value for c in ws_check[fila_headers]]
        fila_datos = ws_check[fila_headers + 1]
        
        print(f"Primeras columnas (fila {fila_headers + 1}):\n")
        
        for i, (h, cell) in enumerate(zip(headers[:20], fila_datos[:20])):
            if h:
                valor = cell.value
                if isinstance(valor, float):
                    valor = f"{valor:.2f}"
                print(f"  Col {i+1:2d}: {h:30s} = {str(valor)[:40]}")
        
        print(f"\n{'='*80}")
        print("VERIFICACIÓN CAMPOS FIJOS")
        print(f"{'='*80}\n")
        
        # Verificar campos fijos
        campos_verificar = [
            ('NUMERO POLIZA', '5924'),
            ('NOMBRE PRODUCTO', 'SALDO DE DEUDA T + C'),
            ('PAIS DE RESIDENCIA', '239')
        ]
        
        for nombre_campo, valor_esperado in campos_verificar:
            encontrado = False
            for col_idx, h in enumerate(headers):
                if h and nombre_campo.upper() in str(h).upper():
                    valor_real = fila_datos[col_idx].value
                    estado = "✓" if str(valor_real) == str(valor_esperado) else "⚠"
                    print(f"{estado} {nombre_campo:30s}")
                    print(f"   Esperado: {valor_esperado}")
                    print(f"   Obtenido: {valor_real}\n")
                    encontrado = True
                    break
            
            if not encontrado:
                print(f"❌ {nombre_campo}: NO ENCONTRADO\n")
        
        print("="*80)
        print("✓ PRUEBA COMPLETADA")
        print("="*80)
    else:
        print("❌ No se encontraron headers")

except Exception as e:
    import traceback
    print(f"\n❌ ERROR: {e}")
    print(traceback.format_exc())
    sys.exit(1)
