# ✅ DESCARGA CON SELECCIÓN DE UBICACIÓN

## 🎯 CAMBIO IMPORTANTE: Usuario elige dónde descargar

Se ha actualizado la funcionalidad de descarga para que el usuario pueda seleccionar la ubicación donde guardar el archivo.

---

## 📥 Cómo Funciona Ahora

### Paso 1: Transformar
```
Usuario selecciona archivo 413
        ↓
Usuario elige póliza (DV)
        ↓
Usuario hace clic "Transformar"
        ↓
Sistema procesa datos
```

### Paso 2: Descargar (NUEVO)
```
✓ Transformación completada
        ↓
Botón "Descargar Resultado" se activa ✨
        ↓
Usuario hace clic "Descargar Resultado"
        ↓
Se abre DIÁLOGO DE SELECCIÓN 📁
        ↓
Usuario elige dónde guardar
        ↓
Usuario hace clic "Guardar"
        ↓
Archivo se guarda en ubicación elegida
        ↓
Se abre carpeta con archivo ✨
```

---

## 🎨 Interfaz de Selección

### Diálogo de Guardado
```
┌─────────────────────────────────────┐
│  Guardar archivo transformado       │
├─────────────────────────────────────┤
│  📁 Mis documentos                  │
│      📄 Facturación_DV_2025-12-13.xlsx
│                                     │
│  Nombre: [Facturación_DV_...]    │
│  Tipo:   [Archivos Excel *.xlsx]  │
│                                     │
│           [Guardar]  [Cancelar]    │
└─────────────────────────────────────┘
```

---

## 🔄 Flujo Técnico

### Vista (`src/vista/principal.py`)
```python
def descargar_resultado(self):
    # 1. Muestra diálogo de selección
    ruta = filedialog.asksaveasfilename(
        title="Guardar archivo transformado",
        defaultextension=".xlsx",
        initialfile="Facturación_DV_2025-12-13.xlsx"
    )
    
    # 2. Si usuario selecciona ubicación
    if ruta and self.callback_descargar:
        # Envía archivos (origen y destino) al controlador
        self.callback_descargar(
            self.archivo_resultado,  # Archivo temporal
            ruta                      # Ubicación elegida
        )
```

### Controlador (`src/controlador/coordinador.py`)
```python
def descargar_archivo(self, ruta_origen, ruta_destino):
    # 1. Copia archivo de temp a ubicación elegida
    shutil.copy2(ruta_origen, ruta_destino)
    
    # 2. Abre carpeta con archivo
    subprocess.Popen(f'explorer /select,"{ruta_destino}"')
    
    # 3. Muestra confirmación
    self.vista.mostrar_exito("Éxito", f"Guardado en:\n{ruta_destino}")
```

---

## 📊 Cambios Realizados

### Archivo: `src/vista/principal.py`
```python
# ANTES: Abría automáticamente ubicación
def descargar_resultado(self):
    if self.callback_descargar:
        self.callback_descargar(self.archivo_resultado)

# AHORA: Muestra diálogo de selección
def descargar_resultado(self):
    ruta_descarga = filedialog.asksaveasfilename(...)
    if ruta_descarga and self.callback_descargar:
        self.callback_descargar(
            self.archivo_resultado,  # Origen
            ruta_descarga            # Destino
        )
```

### Archivo: `src/controlador/coordinador.py`
```python
# ANTES: Solo abría carpeta
def descargar_archivo(self, ruta_archivo):
    subprocess.Popen(f'explorer /select,"{ruta_archivo}"')

# AHORA: Copia a ubicación elegida y abre
def descargar_archivo(self, ruta_origen, ruta_destino):
    shutil.copy2(ruta_origen, ruta_destino)
    subprocess.Popen(f'explorer /select,"{ruta_destino}"')
    self.vista.mostrar_exito("Éxito", ...)
```

### Archivo: `src/controlador/coordinador.py` (_ejecutar_transformacion)
```python
# ANTES: Guardaba en Descargas automáticamente
ruta_descarga = os.path.join(descargas_dir, nombre_descarga)
try:
    wb_resultado.save(ruta_descarga)
except:
    ruta_descarga = ruta_temp

# AHORA: Solo guarda en temp, usuario elige ubicación
temp_dir = tempfile.gettempdir()
ruta_temp = os.path.join(temp_dir, nombre_descarga)
wb_resultado.save(ruta_temp)
self.vista.establecer_archivo_resultado(ruta_temp)
```

---

## ✨ Ventajas

| Aspecto | Antes | Ahora |
|--------|-------|-------|
| Ubicación | Fija (Descargas) | Usuario elige 📁 |
| Flexibilidad | Baja | Alta |
| Control | Automático | Manual |
| Destino | Descargas siempre | Cualquier carpeta |

---

## 🚀 Uso Práctico

### Ejemplo 1: Descargar a Descargas
```
1. Haz clic "Descargar Resultado"
2. Diálogo abre en Descargas (predeterminado)
3. Haz clic "Guardar"
4. Archivo listo en Descargas
```

### Ejemplo 2: Descargar a otra carpeta
```
1. Haz clic "Descargar Resultado"
2. Navega a carpeta deseada (ej: Documentos)
3. Cambia nombre si deseas (opcional)
4. Haz clic "Guardar"
5. Archivo listo en ubicación elegida
```

### Ejemplo 3: Cancelar descarga
```
1. Haz clic "Descargar Resultado"
2. Haz clic "Cancelar"
3. Se cierra diálogo sin descargar
```

---

## 🎯 Mensajes al Usuario

### Durante Transformación
```
✓ Sistema iniciado - Arquitectura MVC
✓ Lógica completa de transformador_excel.py

📌 Instrucciones:
  1. Selecciona archivo 413
  2. Elige póliza
  3. Haz clic en Transformar

════════════════════════════════════════════════════════
```

### Después de Transformar
```
✓ Archivo preparado: Facturación_DV_2025-12-13.xlsx

Haz clic en 'Descargar Resultado' para elegir dónde guardarlo

🎉 ¡Transformación completada exitosamente!
```

### Al Descargar
```
✓ Archivo guardado en:
C:\Users\usuario\Documentos\Facturación_DV_2025-12-13.xlsx

✓ Abriendo carpeta...

Éxito: Archivo guardado en: C:\Users\usuario\Documentos\...
```

---

## 🔧 Detalles Técnicos

### Diálogo de Guardado
```python
filedialog.asksaveasfilename(
    title="Guardar archivo transformado",
    defaultextension=".xlsx",
    initialfile="Facturación_DV_2025-12-13.xlsx",
    filetypes=[
        ("Archivos Excel", "*.xlsx"),
        ("Todos", "*.*")
    ]
)
```

### Copia de Archivo
```python
shutil.copy2(
    ruta_origen,     # Archivo temporal
    ruta_destino     # Ubicación elegida por usuario
)
```

### Abrir Carpeta
```python
# Windows
subprocess.Popen(f'explorer /select,"{ruta_destino}"')

# Mac/Linux
subprocess.Popen(['open', '-R', ruta_destino])
```

---

## ✅ Verificación

✅ Vista: Método `descargar_resultado()` actualizado
✅ Controlador: Método `descargar_archivo()` actualizado
✅ Transformación: Solo guarda en temp
✅ Diálogo: Muestra nombre sugerido
✅ Ubicaciones: Usuario elige dónde guardar
✅ Integración: Abre carpeta después de guardar

---

## 🎉 CONCLUSIÓN

**El flujo de descarga ahora es completamente flexible:**

1. ✅ Transformación automática
2. ✅ **Selección de ubicación por usuario** ← ¡NUEVO!
3. ✅ Copia automática a ubicación elegida
4. ✅ Abre carpeta con archivo

**Flujo completo:**
```bash
python main.py
→ Selecciona archivo
→ Elige póliza
→ Haz clic Transformar
→ Haz clic Descargar Resultado
→ Elige dónde guardar (diálogo)
→ ¡Archivo listo en tu ubicación elegida!
```

---

**Fecha:** Diciembre 13, 2025
**Estado:** ✅ DESCARGA CON SELECCIÓN IMPLEMENTADA
**Próximo:** Sistema completamente listo para usar
