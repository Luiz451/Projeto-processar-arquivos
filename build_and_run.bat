@echo off
echo Limpando builds antigos...
rmdir /s /q dist
rmdir /s /q build

echo.
echo Construindo o novo executavel...
pyinstaller --onefile --windowed --icon=extraction.ico --hidden-import=tkinterdnd2 main.py

echo.
echo Executando o programa...
cd dist
start Extrator de Arquivos.exe

echo Processo concluido!