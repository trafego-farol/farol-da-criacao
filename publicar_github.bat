@echo off
setlocal
cd /d "%~dp0"
echo ============================================================
echo   FAROL DA CRIACAO - publicar no GitHub
echo ============================================================
echo.
git --version >nul 2>&1
if errorlevel 1 (
  echo  O Git nao esta instalado.
  echo  Baixe em https://git-scm.com/download/win
  echo  Instale com as opcoes padrao e rode este arquivo de novo.
  echo.
  pause
  exit /b 1
)
echo  Git encontrado.
echo.
set /p URL=Cole a URL do repositorio que voce criou no GitHub: 
if "%URL%"=="" (
  echo  Nenhuma URL informada.
  pause
  exit /b 1
)
echo.
echo  Conferindo se nenhum arquivo sensivel entrou...
git ls-files | findstr /I /R "api_key\.txt farol_config\.json \.xls \.xlsx snapshot" >nul
if not errorlevel 1 (
  echo.
  echo  ATENCAO: ha arquivo sensivel na lista. NAO vou publicar.
  echo  Confira o .gitignore.
  pause
  exit /b 1
)
echo  Nenhum arquivo sensivel. Seguindo.
echo.
git remote remove origin >nul 2>&1
git remote add origin %URL%
echo  Enviando... (vai abrir o navegador para voce entrar no GitHub)
git push -u origin main
if errorlevel 1 (
  echo.
  echo  O envio falhou. Causas comuns:
  echo   - o repositorio no GitHub foi criado com README (crie vazio)
  echo   - login nao concluido no navegador
  echo   - URL errada
  echo.
  pause
  exit /b 1
)
echo.
echo  ============================================================
echo   PRONTO. Projeto publicado.
echo  ============================================================
pause
