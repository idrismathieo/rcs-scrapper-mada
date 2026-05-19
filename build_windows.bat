@echo off
REM Script de build pour Windows - produit RCS Madagascar.exe
REM Usage : double-cliquer sur ce fichier ou lancer depuis cmd

cd /d "%~dp0"

echo ==========================================
echo   Build RCS Madagascar pour Windows
echo ==========================================
echo.

REM Verifie que Python est installe
where py >nul 2>&1
if %errorlevel% neq 0 (
    echo ERREUR : Python n'est pas installe ou pas dans le PATH.
    echo.
    echo 1. Telecharger Python depuis https://www.python.org/downloads/
    echo 2. Lancer l'installateur en cochant "Add python.exe to PATH"
    echo 3. Relancer ce script
    echo.
    pause
    exit /b 1
)

py --version
echo.

REM Verifie que Tkinter marche
py -c "import tkinter; r = tkinter.Tk(); r.withdraw(); r.destroy()" 2>nul
if %errorlevel% neq 0 (
    echo ERREUR : Tkinter n'est pas disponible avec ce Python.
    echo Reinstaller Python depuis python.org en cochant "tcl/tk and IDLE".
    echo.
    pause
    exit /b 1
)

echo Tkinter : OK
echo.

echo Installation des dependances...
py -m pip install --upgrade pip --quiet
py -m pip install -r requirements.txt pyinstaller --quiet
echo   Dependances installees.
echo.

echo Nettoyage des builds precedents...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "RCS Madagascar.spec" del "RCS Madagascar.spec"
if exist "rcs_scraper.spec" del "rcs_scraper.spec"
echo.

echo Compilation en cours (peut prendre 1-2 minutes)...
py -m PyInstaller ^
    --windowed ^
    --onefile ^
    --name="RCS Madagascar" ^
    --noconfirm ^
    rcs_scraper.py

echo.
echo ==========================================
echo   Compilation terminee !
echo ==========================================
echo.
echo Executable produit :
echo   %cd%\dist\RCS Madagascar.exe
echo.
echo Ce .exe peut etre distribue a n'importe quel utilisateur Windows.
echo Il lui suffit de double-cliquer pour lancer l'interface.
echo.

explorer dist
pause
