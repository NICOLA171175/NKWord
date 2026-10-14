@echo off
rem Avvia NKWord con un doppio clic (Windows).
title NKWord
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

node -e "const [a,b]=process.versions.node.split('.').map(Number);process.exit(a>22||(a===22&&b>=12)?0:1)"
if errorlevel 1 (
  echo.
  for /f %%v in ('node -p "process.versions.node"') do echo La versione di Node.js installata ^(%%v^) e' troppo vecchia: serve la 22.12 o piu' recente.
  echo Scarica e installa la versione LTS da https://nodejs.org, poi riapri questo file.
  echo.
  start "" https://nodejs.org
  pause
  exit /b 1
)

node tools\pacchetti-aggiornati.cjs
if errorlevel 1 (
  echo Installo o aggiorno i pacchetti necessari, attendi...
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
echo Avvio di NKWord: il browser si aprira' da solo.
echo Per chiudere l'app chiudi questa finestra.
echo.
call npm run dev -- --open
pause
