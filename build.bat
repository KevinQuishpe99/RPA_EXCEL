@echo off
echo ========================================
echo Construyendo ejecutable TransformadorExcelRPA
set MODE=%1
if "%MODE%"=="" set MODE=onefile
echo Modo de build: %MODE%

echo.
echo Verificando UPX para compresion avanzada...
upx --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
	echo [AVISO] UPX no esta instalado. El ejecutable puede ser mas grande.
	echo Para instalarlo: descarga desde https://upx.github.io/ y agrega la carpeta al PATH.
) else (
	for /f "tokens=*" %%i in ('upx --version') do echo UPX detectado: %%i & goto :after_upx
)
:after_upx
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
for /r src %%d in (__pycache__) do if exist "%%d" rmdir /s /q "%%d"

echo.
echo Construyendo ejecutable...
if /I "%MODE%"=="onefile" (
	python -m PyInstaller build_exe.spec --clean --noconfirm
) else (
	python -m PyInstaller build_exe_onedir.spec --clean --noconfirm
)

echo.
echo ========================================
echo Construccion completada!
if /I "%MODE%"=="onefile" (
	echo El ejecutable se encuentra en: dist\TransformadorExcelRPA.exe
) else (
	echo La carpeta de distribución se encuentra en: dist\TransformadorExcelRPA\
)
echo ========================================
pause

