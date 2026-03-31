@echo off
REM ============================================================
REM  FlowCheck - Build .exe con PyInstaller
REM  Eseguire dalla root del progetto:
REM    build_exe.bat
REM ============================================================

echo.
echo  ===== FlowCheck - Build EXE =====
echo.

REM Verifica che pyinstaller sia installato
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] PyInstaller non trovato. Installa con:
    echo   pip install pyinstaller
    pause
    exit /b 1
)

REM Pulizia build precedente
if exist dist\FlowCheck rmdir /s /q dist\FlowCheck
if exist build rmdir /s /q build

REM Build
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "FlowCheck" ^
    --hidden-import pandas ^
    --hidden-import pandas._libs.tslibs.base ^
    --hidden-import pandas._libs.tslibs.nattype ^
    --hidden-import pandas._libs.tslibs.np_datetime ^
    --hidden-import pandas._libs.tslibs.timestamps ^
    --hidden-import pandas._libs.tslibs.timedeltas ^
    --hidden-import pandas._libs.tslibs.timezones ^
    --hidden-import pandas._libs.tslibs.parsing ^
    --hidden-import pandas._libs.tslibs.offsets ^
    --hidden-import pandas._libs.tslibs.period ^
    --hidden-import pandas._libs.tslibs.vectorized ^
    --hidden-import pandas._libs.hashtable ^
    --hidden-import pandas._libs.lib ^
    --hidden-import pandas._libs.missing ^
    --hidden-import pandas._libs.writers ^
    --hidden-import pandas._libs.ops ^
    --hidden-import pandas._libs.interval ^
    --hidden-import pandas._libs.indexing ^
    --hidden-import pandas._libs.join ^
    --hidden-import pandas._libs.reduction ^
    --hidden-import pandas._libs.groupby ^
    --hidden-import pandas._libs.window.aggregations ^
    --hidden-import pandas._libs.window.indexers ^
    --hidden-import pandas.io.formats.excel ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --hidden-import openpyxl.utils ^
    --hidden-import openpyxl.chart ^
    --hidden-import openpyxl.drawing ^
    --hidden-import openpyxl.worksheet ^
    --hidden-import tkinter ^
    --hidden-import tkinter.ttk ^
    --hidden-import tkinter.filedialog ^
    --hidden-import tkinter.messagebox ^
    --hidden-import tkinter.scrolledtext ^
    --add-data "flowcheck_engine.py;." ^
    flowcheck_app.py

if errorlevel 1 (
    echo.
    echo [ERRORE] Build fallita. Controlla i messaggi sopra.
    pause
    exit /b 1
)

echo.
echo  ===== Build completata! =====
echo.
echo  EXE disponibile in: dist\FlowCheck.exe
echo.
echo  Puoi copiare dist\FlowCheck.exe su qualsiasi PC Windows
echo  senza necessita' di installare Python.
echo.
pause
