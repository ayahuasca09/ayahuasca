@echo off
REM Setting up paths and variables
set SCRIPT_PATH="F:\pppppy\SP\func\comlib;comlib"
set SCRIPT_NAME="AILanguage_Import_v5.py"
set OUTPUT_NAME="AILanguage_Import"
set DIST_PATH="S:\chen.gong_DCC_Audio\Audio\Tool\Auto_Sound\AILanguage_Import"

REM Running PyInstaller with specified options
pyinstaller --onefile --add-data %SCRIPT_PATH% --name %OUTPUT_NAME% --distpath %DIST_PATH% %SCRIPT_NAME%

REM Notify user of completion
echo PyInstaller build completed.
pause