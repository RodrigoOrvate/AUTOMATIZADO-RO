@echo off
REM Navegar para a raiz do projeto
cd /d "%~dp0\.."

echo ============================================
echo   NeuroTrace - Build Completo
echo   (Executavel + Instalador)
echo ============================================
echo.

REM --- Etapa 1: Criar o executavel com PyInstaller ---
echo [1/2] Criando executavel com PyInstaller...
echo.

python -m PyInstaller packaging\main.spec --clean --noconfirm

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao criar o executavel.
    pause
    exit /b 1
)

echo.
echo [OK] Executavel criado em dist\NeuroTrace.exe
echo.

REM --- Etapa 2: Compilar o instalador com Inno Setup ---
echo [2/2] Compilando instalador com Inno Setup...
echo.

REM Tenta encontrar o Inno Setup no PATH ou nos caminhos padrão
where iscc >nul 2>nul
if %errorlevel% equ 0 (
    iscc packaging\installer.iss
) else if exist "C:\Program Files (x86)\Inno Setup 6\iscc.exe" (
    "C:\Program Files (x86)\Inno Setup 6\iscc.exe" packaging\installer.iss
) else if exist "C:\Program Files\Inno Setup 6\iscc.exe" (
    "C:\Program Files\Inno Setup 6\iscc.exe" packaging\installer.iss
) else (
    echo.
    echo [AVISO] Inno Setup 6 nao encontrado!
    echo Instale em: https://jrsoftware.org/isinfo.php
    echo Depois execute novamente este script.
    echo.
    echo O executavel foi criado com sucesso em dist\NeuroTrace.exe
    pause
    exit /b 1
)

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao compilar o instalador.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Build completo com sucesso!
echo.
echo   Executavel: dist\NeuroTrace.exe
echo   Instalador: installer_output\NeuroTrace_Setup_v2.0.0.exe
echo ============================================
echo.
pause
