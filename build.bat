@echo off
REM ============================================================
REM  FAI Dimension Numberer - PyInstaller build script
REM  Run this from the folder containing both .py files.
REM  Produces:  dist\FAI_Dimension_Numberer.exe
REM ============================================================

echo.
echo ============================================
echo  FAI Dimension Numberer - Build Script
echo ============================================
echo.

REM Check that both source files exist
if not exist "fai_dimension_numberer.py" (
    echo ERROR: fai_dimension_numberer.py not found in current folder.
    echo        Run this script from the folder containing both .py files.
    pause
    exit /b 1
)
if not exist "dimension_detector.py" (
    echo ERROR: dimension_detector.py not found in current folder.
    echo        Both Python files must be in the same folder.
    pause
    exit /b 1
)

echo [1/3] Installing / upgrading dependencies...
python -m pip install -r requirements.txt -q
python -m pip install pyinstaller -q

echo.
echo [2/3] Building executable...
echo.

python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "FAI_Dimension_Numberer" ^
    --add-data "dimension_detector.py;." ^
    fai_dimension_numberer.py

REM Check if build succeeded
if not exist "dist\FAI_Dimension_Numberer.exe" (
    echo.
    echo BUILD FAILED - check the output above for errors.
    pause
    exit /b 1
)

echo.
echo [3/3] Copying measurement template into dist\...
copy /Y "MSI-VSA_Template_Dim_Check.xlsx" "dist\" >nul

echo.
echo ============================================
echo  Build complete!
echo.
echo  Deliverable folder:  dist\
echo    - FAI_Dimension_Numberer.exe
echo    - MSI-VSA_Template_Dim_Check.xlsx
echo.
echo  Zip the dist\ folder and send to coworkers.
echo ============================================
echo.
pause
