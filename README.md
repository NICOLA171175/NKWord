# nikoffice

Clone di Microsoft Word in TypeScript.

## Avvio con un doppio clic

Serve [Node.js](https://nodejs.org) (versione LTS). Poi, nella cartella del progetto:

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
