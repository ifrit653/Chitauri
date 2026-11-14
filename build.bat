@echo off
REM Build ManuscritOCR.exe (run from this folder, Windows only)
REM Prerequisites: Python 3.10+ and Tesseract installed (UB Mannheim build)
REM with the French (fra) and Malagasy (mlg) language data.

pip install pytesseract pillow python-docx pyinstaller || goto :err

REM 1) Copy Tesseract next to the script (once)
if not exist tesseract (
  xcopy "C:\Program Files\Tesseract-OCR" tesseract\ /E /I /Q || goto :err
)

REM 2) Keep only the languages we need, to keep the exe small
for %%f in (tesseract\tessdata\*.traineddata) do (
  if /I not "%%~nf"=="fra" if /I not "%%~nf"=="mlg" if /I not "%%~nf"=="eng" del "%%f"
)

REM 3) Package. Use --onedir instead of --onefile for faster startup.
pyinstaller --noconsole --onefile --name ManuscritOCR ^
  --add-data "tesseract;tesseract" ocr_client.py || goto :err

echo.
echo Done: dist\ManuscritOCR.exe
exit /b 0

:err
echo Build failed.
exit /b 1
