@echo off
chcp 65001 > nul
echo ========================================
echo    ACTUALIZACION MAPA DEMANDA ESTADO
echo ========================================

REM -- El script, el HTML y el repo de GitHub estan en la misma carpeta --
REM -- Solo modifica PYTHON si cambia la ruta de tu Python              --
set PYTHON="C:\Users\csegil\AppData\Local\Microsoft\WindowsApps\python.exe"
set CARPETA=%~dp0
set SCRIPT="%CARPETA%P.DEMANDA_ESTADO.py"

echo.
echo [1/4] Generando mapa y reporte Excel...
echo         (requiere VPN activa)
echo.
cd /d "%CARPETA%"
%PYTHON% %SCRIPT%

if errorlevel 1 (
    echo.
    echo ERROR al generar el mapa. Revisa:
    echo   - VPN conectada?
    echo   - SQL Server accesible?
    echo   - Archivos Excel de entrada existen?
    pause
    exit /b
)

echo.
echo [2/4] Registrando cambios en Git...
cd /d "%CARPETA%"
REM -- Esto crea una regla automatica para que Git ignore tu Excel --
echo resultados_demanda.xlsx > .gitignore
git add .
git commit -m "actualizacion %date% %time:~0,5%"

echo.
echo [3/4] Sincronizando con GitHub...
git pull --rebase

if errorlevel 1 (
    echo.
    echo ERROR en git pull. Puede haber conflictos.
    echo Solucion manual: abre Git Bash y ejecuta 'git rebase --abort'
    pause
    exit /b
)

echo.
echo [4/4] Subiendo mapa a GitHub...
git push

if errorlevel 1 (
    echo.
    echo ERROR en git push. Verifica:
    echo   - Conexion a internet activa?
    echo   - Credenciales de GitHub configuradas?
    pause
    exit /b
)

echo.
echo ========================================
echo  Listo. Streamlit actualiza en ~1 min.
echo  El Excel se ha guardado solo en tu PC.
echo ========================================
pause