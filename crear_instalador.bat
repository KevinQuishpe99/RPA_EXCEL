@echo off
echo ========================================
echo Creando instalador de Demo
echo ========================================
echo.

REM Verificar que el ejecutable existe
if not exist "dist\Demo.exe" (
    echo ERROR: No se encuentra el ejecutable dist\Demo.exe
    echo Por favor ejecuta primero build.bat para crear el ejecutable
    pause
    exit /b 1
)

REM Verificar que Inno Setup esté instalado
set INNO_SETUP="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %INNO_SETUP% (
    set INNO_SETUP="C:\Program Files\Inno Setup 6\ISCC.exe"
)

if not exist %INNO_SETUP% (
    echo ERROR: No se encuentra Inno Setup
    echo Por favor instala Inno Setup desde: https://jrsoftware.org/isinfo.php
    echo O modifica la ruta en este archivo
    pause
    exit /b 1
)

echo.
echo Compilando instalador...
%INNO_SETUP% crear_instalador.iss

echo.
echo ========================================
echo Instalador creado exitosamente!
echo Se encuentra en: instalador\Demo_Instalador.exe
echo ========================================
pause

