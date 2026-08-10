@echo off
setlocal

echo ================================================
echo   Achats - Generateur de contrats fournisseurs
echo ================================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [ERREUR] Node.js n'est pas installe, ou n'est pas reconnu.
  echo Telecharge la version LTS ici puis relance ce fichier : https://nodejs.org/
  echo.
  pause
  exit /b 1
)

echo Node.js detecte :
node -v
echo.

cd /d "%~dp0server"
if not exist node_modules (
  echo Installation des dependances du serveur (premiere fois seulement)...
  call npm install
  if errorlevel 1 goto :error
  echo.
)

cd /d "%~dp0client"
if not exist node_modules (
  echo Installation des dependances du frontend (premiere fois seulement)...
  call npm install
  if errorlevel 1 goto :error
  echo.
)

echo Construction de l'interface...
call npm run build
if errorlevel 1 goto :error
echo.

cd /d "%~dp0server"
echo Demarrage de l'application sur http://localhost:4000 ...
echo (laisse cette fenetre ouverte pendant le test ; ferme-la pour arreter l'application)
echo.

start "" "http://localhost:4000"
call npm run dev

goto :eof

:error
echo.
echo [ERREUR] Une etape a echoue - voir le message ci-dessus.
pause
exit /b 1
