@echo off
echo === Gerando executável com PyInstaller ===

REM Caminho base do MySQL Connector instalado via Microsoft Store
set BASE_PATH=%LOCALAPPDATA%\Packages\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\LocalCache\local-packages\Python311\site-packages\mysql\connector

REM Comando de geração do EXE
python -m PyInstaller ^
  --onefile ^
  --windowed ^
  --collect-submodules=mysql.connector ^
  --add-data "%BASE_PATH%\locales;mysql/connector/locales" ^
  --add-data "%BASE_PATH%\authentication_plugins;mysql/connector/authentication_plugins" ^
  gerador_sql.py

echo.
echo ✅ EXE gerado com sucesso em: dist\gerador_sql.exe
pause

