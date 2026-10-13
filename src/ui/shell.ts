/**
 * Shell dell'applicazione: barra del titolo, barra multifunzione, vista del
 * documento, barra di stato e input da tastiera, collegati a un `Editor`.
 */

import type { Editor } from '../core';
import { ContextMenu, contextMenuItems } from './context-menu';
import { DocumentView } from './document-view';
import { viewExtensions } from './extensions';
import { InputController, runFirstExecutable } from './input';
import { Keymap, defaultKeybindings, isMacPlatform } from './keymap';
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
  /** Apre un file e avvisa chi ascolta `onFileOpened` (es. i documenti recenti). */
  openFile(file: File, from?: FileOrigin): Promise<boolean>;
  /** Mostra la finestra di scelta del file, come "Apri…". */
  pickFile(): void;
  /** Avvisa dopo ogni file aperto con successo. */
  onFileOpened(listener: (opened: OpenedFile) => void): () => void;
  destroy(): void;
}

/** Maniglia di un file del disco (File System Access API, Chrome ed Edge). */
export interface FileHandle {
  readonly kind: 'file';
  readonly name: string;
  getFile(): Promise<File>;
  isSameEntry?(other: FileHandle): Promise<boolean>;
  queryPermission?(options: { mode: 'read' | 'readwrite' }): Promise<PermissionState>;
  requestPermission?(options: { mode: 'read' | 'readwrite' }): Promise<PermissionState>;
}

/** Da dove viene un file aperto: la maniglia del browser o il percorso sul disco (programma desktop). */
export interface FileOrigin {
  readonly handle?: FileHandle;
  readonly path?: string;
}

export interface OpenedFile extends FileOrigin {
  readonly file: File;
}

type OpenFilePicker = (options: object) => Promise<FileHandle[]>;

/** Tipi di file proposti dalla finestra "Apri". */
const OPEN_TYPES = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
  'text/html': ['.html', '.htm'],
  'text/plain': ['.txt', '.md'],
  'application/pdf': ['.pdf'],
};

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
  fileInput.accept = Object.values(OPEN_TYPES).flat().join(',');
  const openedListeners = new Set<(opened: OpenedFile) => void>();
  const openAndNotify = async (file: File, from: FileOrigin = {}) => {
    const ok = await openFile(editor, file);
    if (ok) for (const l of openedListeners) l({ file, ...from });
    input.focus();
    return ok;
  };
  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    fileInput.value = '';
    if (file) void openAndNotify(file);
  });
  // Nel browser che lo permette si tiene la maniglia del file, così si può
  // riaprirlo dai documenti recenti; il programma desktop usa i percorsi.
  const win = root.ownerDocument.defaultView as (Window & { showOpenFilePicker?: OpenFilePicker; nkofficeDesktop?: unknown }) | null;
  const picker = win && !win.nkofficeDesktop ? win.showOpenFilePicker : undefined;
  const pickFile = () => {
    if (!picker) return fileInput.click();
    picker
      .call(win, { types: [{ description: 'Documenti', accept: OPEN_TYPES }] })
      .then(async ([handle]) => {
        if (handle) await openAndNotify(await handle.getFile(), { handle });
      })
      .catch((error: unknown) => {
        // Finestra chiusa senza scegliere: niente da fare.
        if ((error as { name?: string })?.name !== 'AbortError') console.error('Apertura non riuscita', error);
      });
  };

  const ribbon = new Ribbon(editor, uiModule, {
    ...(options.ribbonItems?.length ? { items: [...RIBBON_ITEMS, ...options.ribbonItems] } : {}),
    keymap,
    ask: options.ask,
    handlers: { 'file.open': pickFile },
    // Il focus torna al documento, ma non se il comando ha aperto una finestra
    // (didascalia, intestazione...): lì si deve poter scrivere e premere Esc.
    afterAction: () => {
      const active = root.ownerDocument.activeElement;
      if (!active || active === root.ownerDocument.body || ribbon.element.contains(active)) input.focus();
    },
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
  let revealOnChange = false;
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
      // PagSu/PagGiù: il cursore e la vista si spostano di una schermata, come in Word.
      if ((e.key === 'PageUp' || e.key === 'PageDown') && !e.ctrlKey && !e.metaKey && !e.altKey) {
        const height = view.visibleHeight();
        if (height > 0) {
          const direction = e.key === 'PageUp' ? 'pageUp' : 'pageDown';
          view.element.scrollTop += (direction === 'pageUp' ? -1 : 1) * view.element.clientHeight;
          runFirstExecutable(editor, [{ commandId: 'edit.moveCaret', args: { direction, extend: e.shiftKey, distance: height } }]);
          return true;
        }
        return false;
      }
      // Ctrl+S salva (in Word), Ctrl+O apre: il browser farebbe "Salva pagina" e "Apri file".
      const mod = (e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey;
      const key = e.key.toLowerCase();
      if (mod && key === 's') {
        runFirstExecutable(editor, [{ commandId: 'file.export', args: { formatId: 'docx' } }]);
        return true;
      }
      if (mod && key === 'o') {
        pickFile();
        return true;
      }
      // Ctrl+F e Ctrl+H aprono Trova e Sostituisci (non la ricerca del browser).
      const findKey = e.code === 'KeyF' ? 'f' : e.code === 'KeyH' ? 'h' : key;
      if (mod && (isMacPlatform() ? e.metaKey : e.ctrlKey) && (findKey === 'f' || findKey === 'h')) {
        const item = RIBBON_ITEMS.find((i) => i.tab === 'home' && i.commandId === (findKey === 'f' ? 'edit.find' : 'edit.replace'));
        if (item) ribbon.invoke(item);
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
  view.element.addEventListener('mousedown', () => {
    revealOnChange = false;
    input.focus();
  });
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
  // Dopo un tasto (frecce, PagGiù, testo digitato) la vista segue il cursore.
  const onKeyCapture = () => {
    revealOnChange = true;
  };
  view.element.addEventListener('keydown', onKeyCapture, true);
  view.element.addEventListener('input', onKeyCapture, true);
  disposers.push(() => {
    view.element.removeEventListener('keydown', onKeyCapture, true);
    view.element.removeEventListener('input', onKeyCapture, true);
  });
  disposers.push(
    editor.events.on('change', () => {
      updateTitle();
      if (revealOnChange) {
        revealOnChange = false;
        view.revealCaret();
      }
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
    openFile: openAndNotify,
    pickFile,
    onFileOpened(listener) {
      openedListeners.add(listener);
      return () => openedListeners.delete(listener);
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
