/**
 * Shell dell'applicazione: barra del titolo, barra multifunzione, vista del
 * documento, barra di stato e input da tastiera, collegati a un `Editor`.
 */

import type { Editor } from '../core';
import { ContextMenu, contextMenuItems } from './context-menu';
import { DocumentView } from './document-view';
import { viewExtensions } from './extensions';
import { InputController, runFirstExecutable } from './input';
import { Keymap, defaultKeybindings } from './keymap';
import { createUiModule, type UiModule } from './module';
import { Ribbon } from './ribbon';
import { RIBBON_ITEMS, type UiRibbonItem } from './ribbon-items';
import { StatusBar } from './status-bar';

export interface AppOptions {
  /** Modulo UI già attivato sull'editor; se assente viene creato e attivato. */
  readonly ui?: UiModule;
  readonly keymap?: Keymap;
  readonly ask?: (message: string, initial?: string) => string | null;
  /** Voci aggiunte dai moduli in coda a quelle predefinite (vedi `collectRibbonItems`). */
  readonly ribbonItems?: readonly UiRibbonItem[];
}

export interface App {
  readonly editor: Editor;
  readonly ui: UiModule;
  readonly keymap: Keymap;
  readonly ribbon: Ribbon;
  readonly view: DocumentView;
  readonly input: InputController;
  readonly statusBar: StatusBar;
  /** Menu contestuale aperto, se c'è. */
  readonly contextMenu: ContextMenu | null;
  destroy(): void;
}

const TEXT_EXTENSIONS = new Set(['txt', 'html', 'htm', 'md', 'xml', 'rtf']);

/** Legge un file scelto dall'utente e lo passa a `file.open`. */
export async function openFile(editor: Editor, file: File): Promise<boolean> {
  const ext = file.name.split('.').pop()?.toLowerCase() ?? '';
  const format = editor.services.formats?.byExtension(ext);
  const formatId = format?.id ?? ext;
  const data = TEXT_EXTENSIONS.has(ext) ? await file.text() : new Uint8Array(await file.arrayBuffer());
  if (!editor.canExecute('file.open', { data, formatId })) return false;
  return editor.executeAsync('file.open', { data, formatId });
}

export function createApp(root: HTMLElement, editor: Editor, options: AppOptions = {}): App {
  const disposers: (() => void)[] = [];
  let ui = options.ui;
  if (!ui) {
    ui = createUiModule();
    disposers.push(editor.use(ui));
  }
  const keymap = options.keymap ?? new Keymap(defaultKeybindings());
  const uiModule = ui;
  // Estensioni pubblicate dagli altri moduli (es. controllo ortografico).
  const extensions = viewExtensions(editor);
  for (const ext of extensions) for (const kb of ext.keybindings ?? []) disposers.push(keymap.add(kb));

  const app = document.createElement('div');
  app.className = 'app';

  const title = document.createElement('header');
  title.className = 'title-bar';
  const updateTitle = () => {
    const name = editor.state.doc.metadata.title || 'Documento1';
    title.textContent = `${name} - NKoffice`;
    document.title = title.textContent;
  };

  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.hidden = true;
  fileInput.accept = '.docx,.html,.htm,.txt,.md,.pdf';
  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    fileInput.value = '';
    if (file) void openFile(editor, file).finally(() => input.focus());
  });

  const ribbon = new Ribbon(editor, uiModule, {
    ...(options.ribbonItems?.length ? { items: [...RIBBON_ITEMS, ...options.ribbonItems] } : {}),
    keymap,
    ask: options.ask,
    handlers: { 'file.open': () => fileInput.click() },
    afterAction: () => input.focus(),
  });

  // Barra contestuale di intestazione/piè di pagina, come in Word.
  const hfBar = document.createElement('div');
  hfBar.className = 'hf-bar';
  hfBar.hidden = true;
  const hfLabel = document.createElement('span');
  const hfClose = document.createElement('button');
  hfClose.type = 'button';
  hfClose.textContent = 'Chiudi intestazione e piè di pagina';
  hfClose.addEventListener('click', () => {
    const part = uiModule.view.state.headerFooter;
    if (part) runFirstExecutable(editor, [{ commandId: 'page.headerFooter', args: { part, edit: false } }]);
    input.focus();
  });
  hfBar.append(hfLabel, hfClose);

  const view = new DocumentView(editor, uiModule.view);
  const sink = document.createElement('textarea');
  sink.className = 'input-sink';
  view.element.append(sink);
  const input = new InputController(editor, sink, {
    keymap,
    onKey: (e) => {
      const part = uiModule.view.state.headerFooter;
      if (e.key === 'Escape' && part) {
        runFirstExecutable(editor, [{ commandId: 'page.headerFooter', args: { part, edit: false } }]);
        return true;
      }
      return false;
    },
  });
  const statusBar = new StatusBar(editor, uiModule.view);

  // Il textarea segue il cursore, così l'IME mostra l'anteprima al punto giusto.
  const placeInput = () => {
    const caret = view.caretClientRect();
    const host = view.element.getBoundingClientRect();
    if (!caret) return;
    sink.style.left = `${caret.left - host.left + view.element.scrollLeft}px`;
    sink.style.top = `${caret.top - host.top + view.element.scrollTop}px`;
    sink.style.height = `${Math.max(caret.height, 16)}px`;
  };
  const onViewChange = () => {
    const part = uiModule.view.state.headerFooter;
    hfBar.hidden = part === null;
    hfLabel.textContent = part === 'header' ? 'Intestazione' : part === 'footer' ? 'Piè di pagina' : '';
    placeInput();
  };
  view.element.addEventListener('mousedown', () => input.focus());
  let menu: ContextMenu | null = null;
  const onContextMenu = (e: MouseEvent) => {
    menu?.close();
    const pos = view.positionAt(e.clientX, e.clientY, e.target as Node | null);
    const items = pos ? contextMenuItems(editor, pos) : [];
    if (!items.length) return;
    e.preventDefault();
    menu = new ContextMenu(items, e.clientX, e.clientY, () => {
      menu = null;
      input.focus();
    });
  };
  view.element.addEventListener('contextmenu', onContextMenu);
  disposers.push(() => {
    view.element.removeEventListener('contextmenu', onContextMenu);
    menu?.close();
  });
  disposers.push(
    editor.events.on('change', () => {
      updateTitle();
      placeInput();
    }),
    uiModule.view.events.on('change', onViewChange),
  );

  app.append(title, ribbon.element, hfBar, view.element, statusBar.element, fileInput);
  root.append(app);
  updateTitle();
  onViewChange();
  input.focus();

  return {
    editor,
    ui: uiModule,
    keymap,
    ribbon,
    view,
    input,
    statusBar,
    get contextMenu() {
      return menu;
    },
    destroy() {
      input.destroy();
      ribbon.destroy();
      view.destroy();
      statusBar.destroy();
      for (const d of disposers.splice(0).reverse()) d();
      app.remove();
    },
  };
}
