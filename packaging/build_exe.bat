@echo off
REM Navegar para a raiz do projeto
cd /d "%~dp0\.."

echo ============================================
echo   Criando executavel NeuroTrace...
echo ============================================
echo.

python -m PyInstaller packaging\main.spec --clean --noconfirm

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao criar o executavel.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Executavel criado com sucesso!
echo   Local: dist\NeuroTrace.exe
echo ============================================
echo.
pause
