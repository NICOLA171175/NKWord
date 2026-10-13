#!/usr/bin/env bash
# Avvia nikoffice con un doppio clic (Linux e macOS).
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

NODE_MAJOR=$(node -p 'process.versions.node.split(".")[0]')
if [ "$NODE_MAJOR" -lt 18 ]; then
  echo
  echo "La versione di Node.js installata è troppo vecchia: serve la 18 o più recente."
  echo "Scarica l'ultima LTS da https://nodejs.org"
  apri_link https://nodejs.org
  attendi_e_esci
fi

if [ ! -d node_modules ]; then
  echo "Prima installazione: scarico i pacchetti necessari, attendi..."
  if ! npm install; then
    echo
    echo "Installazione non riuscita. Controlla la connessione a Internet e riprova."
    attendi_e_esci
  fi
fi

echo
echo "Avvio di nikoffice: il browser si aprirà da solo."
echo "Per chiudere l'app chiudi questa finestra (o premi Ctrl+C)."
echo
npm run dev -- --open
