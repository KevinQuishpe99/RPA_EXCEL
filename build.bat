@echo off
echo ========================================
echo Construyendo ejecutable Demo
echo ========================================
echo.

REM Verificar que PyInstaller esté instalado
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo Limpiando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__

echo.
echo Construyendo ejecutable...
python -m PyInstaller build_exe.spec --clean --noconfirm

echo.
echo ========================================
echo Construccion completada!
echo El ejecutable se encuentra en: dist\Demo.exe
echo ========================================
pause

