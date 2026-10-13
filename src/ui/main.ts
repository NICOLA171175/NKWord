/** Punto di ingresso dell'app nel browser (`npm run dev`). */

import { Editor } from '../core';
import { collectModules, collectRibbonItems } from './modules';
import { createUiModule, UI_MODULE_ID } from './module';
import { createApp } from './shell';
import './styles.css';

// I moduli degli altri thread vengono caricati se esistono; quelli mancanti
// lasciano semplicemente i loro pulsanti disabilitati. Ogni cartella di
// src con un index.ts è un modulo: aggiungerne uno non richiede di toccare qui.
const found = import.meta.glob<Record<string, unknown>>(['../*/index.ts', '!../core/index.ts', '!../ui/index.ts'], { eager: true });
const exportsList = Object.values(found);

const editor = new Editor();
const modules = collectModules(exportsList, [UI_MODULE_ID]);
for (const module of modules) editor.use(module);
const ui = createUiModule();
editor.use(ui);

const root = document.getElementById('app');
if (!root) throw new Error('Elemento #app mancante');
const app = createApp(root, editor, { ui, ribbonItems: collectRibbonItems(exportsList) });
// Un modulo con un metodo `mount(app)` aggiunge la propria interfaccia (es. il riquadro dei commenti).
for (const module of modules) (module as { mount?: (host: typeof app) => unknown }).mount?.(app);

editor.events.on('error', ({ commandId, error }) => console.error(`Comando ${commandId} fallito`, error));
Object.assign(globalThis, { nikoffice: app });
