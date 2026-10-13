/**
 * Modulo collegamenti ipertestuali (Inserisci > Collegamento, Ctrl+K):
 * collegamenti a pagine web, indirizzi email e punti del documento, da
 * modificare o rimuovere col clic destro e da aprire con Ctrl+clic. Gli
 * indirizzi scritti nel testo diventano collegamenti da soli, come in Word.
 */

import { isCollapsed, orderedRange, findParagraph, paragraphLength, type Editor, type EditorModule, type EditorState, type Position, type RibbonItem } from '../core';
import { addViewExtension, type ContextMenuItem, type ViewExtension } from '../ui/extensions';
import { createLinkCommands, type LinkCommandOptions } from './commands';
import { openLinkDialog } from './dialog';
import { describeLink, findAutoLink, linkAt, linkKind } from './model';
import './styles.css';

export const LINKS_MODULE_ID = 'links';

/** Comandi dopo i quali un indirizzo appena scritto diventa collegamento (spazio, Tab, Invio). */
const AUTO_TRIGGERS = new Set(['edit.insertText', 'edit.insertTab', 'edit.splitParagraph', 'edit.insertLineBreak']);

export interface LinksModule extends EditorModule {
  /** Collega Ctrl+clic e lo scorrimento ai punti del documento (chiamato da `src/ui/main.ts`). */
  mount(app: LinksHost): () => void;
}

export interface LinksHost {
  readonly view: { readonly element: HTMLElement; caretClientRect(): DOMRect | null };
  readonly input?: { focus(): void };
}

export interface LinksModuleOptions extends LinkCommandOptions {
  /** Riconoscimento automatico degli indirizzi scritti (default: sì). */
  readonly autoLink?: boolean;
}

/** Posizione in cui finiva il testo prima di uno spazio, Tab o Invio. */
function typedEnd(previous: EditorState): Position {
  return isCollapsed(previous.selection) ? previous.selection.focus : orderedRange(previous.doc, previous.selection)[0];
}

function autoLinkAfter(editor: Editor, previous: EditorState, commandId: string, args: unknown): void {
  if (commandId === 'edit.insertText' && !/^[  ]$/.test((args as { text?: string } | undefined)?.text ?? '')) return;
  const end = typedEnd(previous);
  const p = findParagraph(editor.state.doc, end.paragraphId)?.paragraph;
  if (!p || end.offset > paragraphLength(p)) return;
  const found = findAutoLink(p, end.offset);
  if (found) editor.execute('link.auto', { paragraphId: p.id, start: found.start, end: found.end, url: found.url });
}

export function createLinksModule(options: LinksModuleOptions = {}): LinksModule {
  let host: LinksHost | null = null;
  let editorRef: Editor | null = null;

  /** Porta in vista il cursore dopo un salto a un punto del documento. */
  const reveal = () => {
    const view = host?.view;
    const rect = view?.caretClientRect();
    if (!view || !rect) return;
    const box = view.element.getBoundingClientRect();
    if (rect.top < box.top || rect.bottom > box.bottom) view.element.scrollTop += rect.top - box.top - box.height / 3;
  };

  return {
    id: LINKS_MODULE_ID,
    activate(editor) {
      editorRef = editor;
      const commands = createLinkCommands(editor, {
        openDialog: options.openDialog ?? (typeof document !== 'undefined' ? openLinkDialog : undefined),
        ...(options.openExternal ? { openExternal: options.openExternal } : {}),
        onNavigate: () => {
          options.onNavigate?.();
          reveal();
        },
      });
      const run = (id: string, args?: unknown) => {
        void editor.executeAsync(id, args).then(() => host?.input?.focus());
      };
      const extension: ViewExtension = {
        id: LINKS_MODULE_ID,
        keybindings: [{ key: 'Mod-k', commandId: 'link.dialog' }],
        contextMenu(pos, state) {
          const link = linkAt(state.doc, pos);
          if (!link) return [{ label: 'Collegamento…', run: () => run('link.dialog') }];
          const range = { paragraphId: link.paragraphId, start: link.start, end: link.end };
          const items: ContextMenuItem[] = [
            {
              label: 'Modifica collegamento…',
              run: () => {
                // Come in Word si seleziona il collegamento, così la finestra lo riconosce.
                const at = (offset: number) => ({ paragraphId: link.paragraphId, offset });
                editor.execute('edit.setSelection', { selection: { anchor: at(link.start), focus: at(link.end) } });
                run('link.dialog');
              },
            },
            { label: 'Apri collegamento', run: () => run('link.open', { url: link.url }), disabled: !editor.canExecute('link.open', { url: link.url }) },
          ];
          if (linkKind(link.url) !== 'place') {
            items.push({ label: 'Copia collegamento', run: () => void globalThis.navigator?.clipboard?.writeText(link.url).catch(() => {}) });
          }
          items.push({ label: 'Rimuovi collegamento', run: () => run('link.remove', { range }) });
          return items;
        },
        decoratePages(pages) {
          // Suggerimento al passaggio del mouse, come in Word.
          const doc = editor.state.doc;
          for (const page of pages) {
            for (const el of Array.from(page.querySelectorAll<HTMLElement>('[data-link]'))) {
              el.title = `${describeLink(doc, el.dataset.link!)}\nCtrl+clic per aprire il collegamento`;
            }
          }
        },
      };

      let previous: EditorState | null = null;
      const disposers = [
        ...commands.map((c) => editor.commands.register(c)),
        addViewExtension(editor, extension),
        editor.events.on('change', (e) => {
          previous = e.previous;
        }),
        editor.events.on('executed', ({ commandId, args }) => {
          if (options.autoLink === false || !previous || !AUTO_TRIGGERS.has(commandId)) return;
          const before = previous;
          previous = null;
          autoLinkAfter(editor, before, commandId, args);
        }),
      ];
      return () => {
        disposers.forEach((d) => d());
        if (editorRef === editor) editorRef = null;
      };
    },
    mount(app) {
      host = app;
      const el = app.view.element;
      // Ctrl+clic su un collegamento lo apre invece di spostare il cursore.
      const onDown = (e: MouseEvent) => {
        if (e.button !== 0 || !(e.ctrlKey || e.metaKey) || !editorRef) return;
        const url = (e.target as Element | null)?.closest?.<HTMLElement>('[data-link]')?.dataset.link;
        if (!url) return;
        e.preventDefault();
        e.stopPropagation();
        void editorRef.executeAsync('link.open', { url });
      };
      // Con Ctrl premuto il puntatore diventa la manina sopra i collegamenti.
      const onKey = (e: KeyboardEvent) => el.classList.toggle('links-follow', e.ctrlKey || e.metaKey);
      const onBlur = () => el.classList.remove('links-follow');
      el.addEventListener('mousedown', onDown, true);
      document.addEventListener('keydown', onKey);
      document.addEventListener('keyup', onKey);
      globalThis.addEventListener?.('blur', onBlur);
      return () => {
        el.removeEventListener('mousedown', onDown, true);
        document.removeEventListener('keydown', onKey);
        document.removeEventListener('keyup', onKey);
        globalThis.removeEventListener?.('blur', onBlur);
        if (host === app) host = null;
      };
    },
  };
}

type Item = RibbonItem & { readonly label: string };

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly Item[] = [
  { tab: 'insert', group: 'Collegamenti', commandId: 'link.dialog', label: 'Collegamento', icon: '🔗', kind: 'button' },
];

export { createLinkCommands, insertLink, removeLink, selectionText, type LinkCommandArgs, type LinkCommandId, type LinkDialogInput, type LinkDialogResult, type LinkInsertArgs, type LinkSpan } from './commands';
export { buildMailto, openLinkDialog, parseMailto } from './dialog';
export * from './model';
