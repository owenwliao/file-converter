@echo off
echo Installing dependencies...
py -m pip install -r requirements.txt --quiet
py -m pip install pyinstaller --quiet

echo.
echo Building FileConverter.exe...
py -m PyInstaller --onefile --noconsole --name "FileConverter" app.py

echo.
if exist dist\FileConverter.exe (
    echo SUCCESS: dist\FileConverter.exe is ready.
    echo Size: ~30-50MB, launches in 1-2 seconds.
) else (
    echo FAILED. Check output above for errors.
)
pause
