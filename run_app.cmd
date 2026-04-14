@echo off
setlocal
set "APP_DIR=%~dp0"

where python >nul 2>nul
if not errorlevel 1 (
	python "%APP_DIR%app.py"
	goto :eof
)

where py >nul 2>nul
if not errorlevel 1 (
	py -3 "%APP_DIR%app.py"
	goto :eof
)

echo Python nao encontrado no PATH.
echo Instale Python 3.11+ ou recrie a .venv do projeto.
exit /b 1
