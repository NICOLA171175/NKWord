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
if %NODE_MAJOR% LSS 20 (
  echo.
  echo La versione di Node.js installata ^(%NODE_MAJOR%^) e' troppo vecchia: serve la 20 o piu' recente.
  echo Scarica e installa la versione LTS da https://nodejs.org, poi riapri questo file.
  echo.
  start "" https://nodejs.org
  pause
  exit /b 1
)

if not exist node_modules\vite\package.json (
  echo Prima installazione: scarico i pacchetti necessari, attendi...
  call npm install
  if errorlevel 1 (
    if exist node_modules rmdir /s /q node_modules
    echo.
    echo Installazione non riuscita. Leggi i messaggi qui sopra.
    echo Cause piu' comuni: nessuna connessione a Internet, oppure Node.js da aggiornare.
    echo Per aggiornare Node.js scarica la versione LTS da https://nodejs.org, poi riapri questo file.
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
