@echo off
cd /d "%~dp0"
python -m unittest -v
if errorlevel 1 (
    echo تعذر اجتياز الاختبارات. لم يتم بناء البرنامج.
    pause
    exit /b 1
)
python -m PyInstaller --clean --noconfirm --onefile --windowed --name AlumetalManager app.py
if errorlevel 1 (
    echo تعذر بناء البرنامج. تأكد من تثبيت Python والحزم الموجودة في requirements.txt.
    pause
    exit /b 1
)
echo تم إنشاء النسخة المستقلة في dist\AlumetalManager.exe
pause
