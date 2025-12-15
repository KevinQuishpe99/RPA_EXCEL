# TC DATA TRANSFER FIX - IMPLEMENTATION COMPLETE

## PROBLEMA ORIGINAL
El flujo TC (5924) no estaba transfiriendo datos correctamente del archivo 455 a la plantilla TC. 
Aunque la configuración y el mapeo estaban en su lugar, **0 filas se procesaban**, dejando la 
hoja de resultados completamente vacía.

## ROOT CAUSE ANALYSIS

### Issue 1: Row Validation Logic Failed for TC
**Problema:** El archivo 455 tiene una columna vacía (índice 0, todas NaN). La lógica de validación 
de filas en `transferencia_datos.py` verificaba `if pd.notna(primera_col)` donde `primera_col` era 
la columna 0. Como es toda NaN, **todas las filas fueron rechazadas**.

**Línea problema:** `transferencia_datos.py:33` - `primera_columna_array = df_origen.iloc[fila_origen_inicio:, 0].values`

**Solución:** Detectar cuando estamos en TC con mapeo manual y usar la primera columna mapeada 
(que es la 1 = PRIMER APELLIDO) para validación en lugar de la columna 0:

```python
# Para TC, la primera columna podría estar vacía - buscar primera columna mapeada
col_validacion = 0
if poliza_info and poliza_info.get('prefijo') == 'TC':
    # Encontrar la primera columna mapeada (generalmente será 1 o mayor)
    if mapeo:
        col_indices = sorted([k for k in mapeo.keys() if isinstance(k, int)])
        if col_indices and col_indices[0] > 0:
            col_validacion = col_indices[0]

primera_columna_array = df_origen.iloc[fila_origen_inicio:, col_validacion].values
```

### Issue 2: Fixed Fields Not Applied
**Problema:** Los campos fijos (NUMERO POLIZA='5924', NOMBRE PRODUCTO, PAIS DE RESIDENCIA='239') 
no se estaban asignando a las filas transferidas.

**Causa:** La función `_aplicar_campos_fijos_tc()` en `transferencia_datos.py` buscaba campos 
como `numero_poliza_fijo`, pero el `poliza_info` que se pasaba tenía valores en campos como 
`numero`, `nombre`, `pais_residencia`.

**Solución:** En `transformador.py` línea ~130, antes de pasar `poliza_info` a `transferir_datos()`, 
completar los campos necesarios:

```python
# Completar poliza_info con valores fijos para TC
if poliza_info and poliza_info.get('prefijo') == 'TC':
    poliza_info = dict(poliza_info)  # Hacer copia
    if 'numero_poliza_fijo' not in poliza_info:
        poliza_info['numero_poliza_fijo'] = poliza_info.get('numero')
    if 'nombre_producto_fijo' not in poliza_info:
        poliza_info['nombre_producto_fijo'] = poliza_info.get('nombre')
    if 'pais_residencia_fijo' not in poliza_info:
        poliza_info['pais_residencia_fijo'] = poliza_info.get('pais_residencia')
```

## CHANGES IMPLEMENTED

### 1. `src/modelo/transferencia_datos.py` (Lines 19-35)
- **Modificación:** Detectar cuando estamos procesando TC y usar la primera columna mapeada 
  para validación de filas en lugar de la columna 0
- **Impacto:** Las 6203 filas de datos ahora se transfieren correctamente

### 2. `src/modelo/transformador.py` (Lines 127-139)
- **Adición:** Antes de llamar `transferir_datos()`, completar `poliza_info` con los campos 
  esperados por `_aplicar_campos_fijos_tc()`
- **Impacto:** Los campos fijos ahora se aplican correctamente a cada fila

## TEST RESULTS

### Validación TC - prueba_tc_validacion.py
✅ **6203 filas procesadas** (de 6211 total)
✅ **Datos mapeados correctamente:**
- PRIMER APELLIDO: MIRAMAG → MIRAMAG ✓
- SEGUNDO APELLIDO: DELGADO → DELGADO ✓  
- SALDO A LA FECHA: 2020-11-20 → 2020-11-20 ✓

✅ **Campos fijos aplicados:**
- NUMERO POLIZA: '5924' ✓
- NOMBRE PRODUCTO: 'SALDO DE DEUDA T + C' ✓
- PAIS DE RESIDENCIA: '239' ✓

### Integración DV/TC
✅ Todos los módulos importan sin errores
✅ Configuración de ambas pólizas correcta:
   - DV: Hoja origen = "Report_AseguradoraMensual"
   - TC: Hoja origen = "Report_AseguradoraSaldos_COVID"
✅ Mapeo TC íntegro: 39 columnas mapeadas
✅ Protecciones TC en su lugar: 5 columnas sin sobrescribir, 3 campos fijos

## FILES MODIFIED

1. **src/modelo/transferencia_datos.py**
   - Lineas 19-35: Lógica inteligente de validación de filas para TC

2. **src/modelo/transformador.py**
   - Líneas 127-139: Completar campos fijos en poliza_info antes de transferencia

## COMPATIBILITY

✅ **DV (413)** - No se vio afectado
   - La lógica de TC se aplica solo cuando `poliza_info.get('prefijo') == 'TC'`
   - Cambios son backwards-compatible

✅ **TC (5924)** - Ahora totalmente funcional
   - Transfiere 6203 filas en lugar de 0
   - Aplica campos fijos correctamente
   - Mapeo manual MAPEO_TC_MANUAL funciona perfectamente

## VERIFICATION STEPS COMPLETED

1. ✅ Identificar root cause (columna 0 vacía)
2. ✅ Implementar fix para validación de filas
3. ✅ Implementar aplicación de campos fijos
4. ✅ Crear test exhaustivo (prueba_tc_validacion.py)
5. ✅ Verificar integridad de módulos
6. ✅ Validar ambas pólizas (DV y TC)
7. ✅ Sin errores de sintaxis en workspace

## NEXT STEPS (PARA USUARIO)

1. Ejecutar prueba completa con la aplicación Qt (main_qt.py)
2. Confirmar que tanto DV como TC funcionan en la UI
3. Verificar que los nombres de archivo sugeridos aparecen correctamente
4. Validar que los datos se descargan con format correcto

## TECHNICAL NOTES

- **MAPEO_TC_MANUAL:** 39 columnas mapeadas (origen 0-based → destino 1-based)
- **NO_SOBRESCRIBIR_TC:** {8, 9, 10, 17, 35} - Columnas con fórmulas protegidas
- **CAMPOS_FIJOS_TC:** {14: '239', 56: '5924', 57: 'SALDO DE DEUDA T + C'} (1-based indices)
- **Primeras columnas mapeadas:** 1→1 (PRIMER APELLIDO), 2→2 (SEGUNDO APELLIDO), etc.
