@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo.
echo  ============================================
echo    E500 3D Print - Control de Impresiones
echo  ============================================
echo.

if not exist "node_modules" (
    echo  Primera ejecucion detectada.
    echo  Instalando dependencias, espera un momento...
    echo.
    npm install
    echo.
    echo  Instalacion completada.
    echo.
)

echo  Iniciando la aplicacion...
echo.
npx electron .
