/** Punto di ingresso dell'app nel browser (`npm run dev`). */

import { Editor } from '../core';
import { collectModules } from './modules';
import { createUiModule, UI_MODULE_ID } from './module';
import { createApp } from './shell';
import './styles.css';

// I moduli degli altri thread vengono caricati se esistono; quelli mancanti
// lasciano semplicemente i loro pulsanti disabilitati. Ogni cartella di
// src con un index.ts è un modulo: aggiungerne uno non richiede di toccare qui.
const found = import.meta.glob<Record<string, unknown>>(['../*/index.ts', '!../core/index.ts', '!../ui/index.ts'], { eager: true });

const editor = new Editor();
for (const module of collectModules(Object.values(found), [UI_MODULE_ID])) editor.use(module);
const ui = createUiModule();
editor.use(ui);

const root = document.getElementById('app');
if (!root) throw new Error('Elemento #app mancante');
const app = createApp(root, editor, { ui });

editor.events.on('error', ({ commandId, error }) => console.error(`Comando ${commandId} fallito`, error));
Object.assign(globalThis, { nikoffice: app });
