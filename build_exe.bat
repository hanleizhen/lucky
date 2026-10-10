@echo off
setlocal
py -3.12 -m pip install -r requirements.txt || goto :error
py -3.12 -m pytest || goto :error
rem Do not let unrelated host DLL directories (for example image/PDF tools)
rem participate in PyInstaller's Qt dependency scan.  Mixing their ICU DLLs
rem with PySide6 makes a packaged program fail before its window is shown.
set "PATH=%SystemRoot%\System32;%SystemRoot%"
py -3.12 -m PyInstaller --noconfirm --clean --distpath build\dist --workpath build\work cnc_program_sheet.spec || goto :error
for /f %%v in ('py -3.12 -c "from cnc_program_sheet.version import __version__; print(__version__)"') do set "APP_VERSION=%%v"
echo.
echo EXE folder created: build\dist\CNCProgramSheet
echo To build the installer, install Inno Setup 6 and run:
echo "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DMyAppVersion=%APP_VERSION% installer.iss
exit /b 0
:error
echo Build failed. Read the error above.
exit /b 1
