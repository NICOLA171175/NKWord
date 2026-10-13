@echo off
rem Avvia nikoffice con un doppio clic (Windows).
title nikoffice
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo Node.js non e' installato.
  echo Scaricalo da https://nodejs.org ^(versione LTS^), installalo e poi riapri questo file.
  echo.
  start "" https://nodejs.org
  pause
  exit /b 1
)

for /f "tokens=1 delims=." %%v in ('node -p "process.versions.node"') do set NODE_MAJOR=%%v
if %NODE_MAJOR% LSS 18 (
  echo.
  echo La versione di Node.js installata e' troppo vecchia: serve la 18 o piu' recente.
  echo Scarica l'ultima LTS da https://nodejs.org
  echo.
  start "" https://nodejs.org
  pause
  exit /b 1
)

if not exist node_modules (
  echo Prima installazione: scarico i pacchetti necessari, attendi...
  call npm install
  if errorlevel 1 (
    echo.
    echo Installazione non riuscita. Controlla la connessione a Internet e riprova.
    pause
    exit /b 1
  )
)

echo.
echo Avvio di nikoffice: il browser si aprira' da solo.
echo Per chiudere l'app chiudi questa finestra.
echo.
call npm run dev -- --open
pause
