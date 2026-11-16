@echo off
setlocal
cd /d "%~dp0"
py -m pip install -r requirements-build.txt
if errorlevel 1 goto :error
py -m PyInstaller --noconfirm --clean --onefile --windowed --name ChangeImpactAnalyzer --collect-all fitz --exclude-module pandas.tests --exclude-module numpy.tests app.py
if errorlevel 1 goto :error
echo.
echo EXE created: dist\ChangeImpactAnalyzer.exe
pause
exit /b 0
:error
echo Build failed. Check Python installation and error messages above.
pause
exit /b 1
