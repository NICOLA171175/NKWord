#!/usr/bin/env bash
# Avvia NKoffice con un doppio clic (Linux e macOS).
cd "$(dirname "$0")" || exit 1

apri_link() {
  if command -v open >/dev/null 2>&1; then open "$1"
  elif command -v xdg-open >/dev/null 2>&1; then xdg-open "$1" >/dev/null 2>&1
  fi
}

attendi_e_esci() {
  echo
  read -r -p "Premi Invio per chiudere..." _
  exit 1
}

if ! command -v node >/dev/null 2>&1; then
  echo
  echo "Node.js non è installato."
  echo "Scaricalo da https://nodejs.org (versione LTS), installalo e poi riapri questo file."
  apri_link https://nodejs.org
  attendi_e_esci
fi

if ! node -e 'const [a,b]=process.versions.node.split(".").map(Number);process.exit(a>22||(a===22&&b>=12)?0:1)'; then
  echo
  echo "La versione di Node.js installata ($(node -p 'process.versions.node')) è troppo vecchia: serve la 22.12 o più recente."
  echo "Scarica e installa la versione LTS da https://nodejs.org, poi riapri questo file."
  apri_link https://nodejs.org
  attendi_e_esci
fi

if ! node tools/pacchetti-aggiornati.cjs; then
  echo "Installo o aggiorno i pacchetti necessari, attendi..."
  if ! npm install; then
    rm -rf node_modules
    echo
    echo "Installazione non riuscita. Leggi i messaggi qui sopra."
    echo "Cause più comuni: nessuna connessione a Internet, oppure Node.js da aggiornare."
    echo "Per aggiornare Node.js scarica la versione LTS da https://nodejs.org, poi riapri questo file."
    attendi_e_esci
  fi
fi

echo
echo "Avvio di NKoffice: il browser si aprirà da solo."
echo "Per chiudere l'app chiudi questa finestra (o premi Ctrl+C)."
echo
npm run dev -- --open
