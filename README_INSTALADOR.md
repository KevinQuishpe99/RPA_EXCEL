# Guía para crear el instalador de Demo

## Requisitos previos

1. **Python 3.8 o superior** instalado
2. **PyInstaller** (se instalará automáticamente)
3. **Inno Setup** (opcional, solo si quieres crear un instalador con interfaz)

## Pasos para crear el ejecutable

### Opción 1: Usando el script automático (Recomendado)

1. Abre una terminal en la carpeta del proyecto
2. Ejecuta:
   ```batch
   build.bat
   ```
3. El ejecutable se creará en la carpeta `dist\Demo.exe`

### Opción 2: Manualmente

1. Instala las dependencias:
   ```batch
   pip install -r requirements.txt
   ```

2. Ejecuta PyInstaller:
   ```batch
   pyinstaller build_exe.spec --clean --noconfirm
   ```

3. El ejecutable estará en `dist\Demo.exe`

## Crear el instalador (Opcional)

Si quieres crear un instalador con interfaz gráfica:

1. **Instala Inno Setup** desde: https://jrsoftware.org/isinfo.php

2. Ejecuta:
   ```batch
   crear_instalador.bat
   ```

3. El instalador se creará en `instalador\Demo_Instalador.exe`

## Notas importantes

- La plantilla.xlsx está empaquetada dentro del ejecutable
- El ejecutable es independiente, no requiere Python instalado
- El tamaño del ejecutable será aproximadamente 50-100 MB (incluye todas las librerías)

## Distribución

Para distribuir la aplicación, solo necesitas:
- **Opción 1**: Enviar el archivo `dist\Demo.exe` directamente
- **Opción 2**: Enviar el instalador `instalador\Demo_Instalador.exe` (más profesional)

## Solución de problemas

### Error: "No se encuentra plantilla.xlsx"
- Asegúrate de que `plantilla.xlsx` esté en la misma carpeta que `transformador_excel.py`
- Verifica que el archivo `.spec` incluya la plantilla en la sección `datas`

### El ejecutable es muy grande
- Esto es normal, incluye Python y todas las librerías necesarias
- Puedes usar UPX para comprimir (ya está habilitado en el .spec)

### El ejecutable no inicia
- Verifica que todas las dependencias estén en `requirements.txt`
- Revisa que `console=False` en el .spec para aplicaciones GUI

