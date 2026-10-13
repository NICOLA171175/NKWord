# nikoffice

Clone di Microsoft Word in TypeScript.

## Programma da installare (senza Node.js)

Scarica dalla pagina [Release "ultima versione"](https://github.com/NICOLA171175/nikoffice/releases/latest):

- **Windows, installazione**: [`nikoffice-installa.exe`](https://github.com/NICOLA171175/nikoffice/releases/latest/download/nikoffice-installa.exe). Crea il collegamento sul desktop e nel menu Start; si disinstalla da *Impostazioni > App*.
- **Windows, senza installare**: [`nikoffice-portable.exe`](https://github.com/NICOLA171175/nikoffice/releases/latest/download/nikoffice-portable.exe). Doppio clic e parte, anche da chiavetta USB.
- **Mac**: [`nikoffice-mac.dmg`](https://github.com/NICOLA171175/nikoffice/releases/latest/download/nikoffice-mac.dmg). Aprilo e trascina nikoffice in *Applicazioni*; la prima volta clic destro sull'app, *Apri*.

Il programma non è firmato digitalmente: Windows può mostrare l'avviso SmartScreen, clicca *Ulteriori informazioni* e poi *Esegui comunque*.
I file vengono ricostruiti in automatico (GitHub Action "Programma desktop") a ogni modifica su `main`; il codice del pacchetto è nella cartella `desktop/` (Electron + electron-builder).

## Avvio con un doppio clic

Serve [Node.js](https://nodejs.org) 22.12 o più recente (versione LTS). Poi, nella cartella del progetto:

- **Windows**: doppio clic su `avvia.bat`.
- **macOS**: doppio clic su `avvia.command`. La prima volta macOS può bloccarlo: clic destro sul file, *Apri*, poi *Apri* di nuovo.
- **Linux**: esegui `./avvia.sh`.

La prima volta scarica i pacchetti (qualche minuto), poi apre l'app nel browser. Per chiuderla basta chiudere la finestra nera del terminale.

## Sviluppo

```bash
npm install
npm run typecheck
npm test
```

## Struttura

| Cartella | Contenuto |
| --- | --- |
| `src/core` | Base comune: modello del documento, unità, stato, comandi, cronologia, contratti fra moduli. Si modifica solo coordinandosi. |
| `src/editing` | Comandi di modifica del testo, formattazione di carattere e paragrafo, stili. |
| `src/structures` | Comandi per tabelle, elenchi puntati/numerati, immagini, interruzioni e sezioni. |
| `src/layout` | Impaginazione e rendering delle pagine. |
| `src/io` | Formati di file (DOCX, HTML, testo) e appunti. |
| `src/ui` | Interfaccia: barra multifunzione, tastiera, finestre di dialogo. |

### Regole del core

- Il modello (`src/core/types.ts`) è immutabile; le lunghezze sono in punti.
- Ogni azione è un comando con id `<area>.<azione>` registrato in `editor.commands`; i comandi sono funzioni pure `stato → transazione`.
- I moduli si agganciano con `editor.use(modulo)` e pubblicano i loro servizi in `editor.services` secondo le interfacce di `src/core/interfaces.ts`.
- Le posizioni sono `{ paragraphId, offset }`: ogni inline non testuale vale 1.
