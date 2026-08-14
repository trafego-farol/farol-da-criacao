@echo off
cd /d "%~dp0"
echo.
echo [1/2] extraindo do iClips...
py src\farol_api.py
if errorlevel 1 (
  echo.
  echo  Falhou a extracao. O dashboard NAO foi atualizado.
  pause
  exit /b 1
)
echo.
echo [2/2] gerando o dashboard...
py src\farol_pasta.py
echo.
echo Pronto.
pause
