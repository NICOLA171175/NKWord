/**
 * Modulo Appunti: i pulsanti Incolla, Taglia e Copia della scheda Home di
 * Word. Usano gli appunti del sistema (`navigator.clipboard`) con lo stesso
 * formato di Ctrl+C / Ctrl+V: HTML per mantenere la formattazione più il
 * testo semplice. Se il browser non dà accesso agli appunti si spiega di
 * usare la tastiera, come fa Word nel browser.
 */

import { isCollapsed, orderedRange, type AsyncCommand, type Editor, type EditorModule, type EditorState } from '../core';

export const CLIPBOARD_MODULE_ID = 'clipboard';

export interface ClipboardCommandArgs {
  'clipboard.cut': void;
  'clipboard.copy': void;
  /** `textOnly`: "Mantieni solo il testo". */
  'clipboard.paste': { textOnly?: boolean } | undefined;
}

/** Accesso agli appunti del sistema; iniettabile per i test. */
export interface SystemClipboard {
  write(data: { html: string; text: string }): Promise<void>;
  read(): Promise<{ html?: string; text?: string }>;
}

export interface ClipboardModuleOptions {
  readonly system?: SystemClipboard;
  /** Messaggio quando gli appunti non sono accessibili (default: `alert`). */
  readonly notify?: (message: string) => void;
}

const isMac = () => /mac|iphone|ipad/i.test(globalThis.navigator?.platform ?? '');

export function browserClipboard(): SystemClipboard {
  return {
    async write({ html, text }) {
      const clip = globalThis.navigator?.clipboard;
      if (!clip) throw new Error('Appunti non disponibili');
      if (typeof ClipboardItem !== 'undefined' && clip.write) {
        await clip.write([
          new ClipboardItem({
            'text/html': new Blob([html], { type: 'text/html' }),
            'text/plain': new Blob([text], { type: 'text/plain' }),
          }),
        ]);
      } else await clip.writeText(text);
    },
    async read() {
      const clip = globalThis.navigator?.clipboard;
      if (!clip) throw new Error('Appunti non disponibili');
      if (clip.read) {
        const out: { html?: string; text?: string } = {};
        // HTML originale, non quello "ripulito" da Chrome che aggiunge colori e
        // caratteri espliciti a ogni pezzo di testo (come fa già Ctrl+V).
        const read = clip.read as (options?: { unsanitized?: string[] }) => Promise<ClipboardItems>;
        const items = await read.call(clip, { unsanitized: ['text/html'] }).catch(() => clip.read());
        for (const item of items) {
          if (!out.html && item.types.includes('text/html')) out.html = await (await item.getType('text/html')).text();
          if (!out.text && item.types.includes('text/plain')) out.text = await (await item.getType('text/plain')).text();
        }
        return out;
      }
      return { text: await clip.readText() };
    },
  };
}

export function createClipboardModule(options: ClipboardModuleOptions = {}): EditorModule {
  const system = options.system ?? browserClipboard();
  const notify = options.notify ?? ((m: string) => globalThis.alert?.(m));
  const key = (k: string) => `${isMac() ? 'Cmd' : 'Ctrl'}+${k}`;
  let editor: Editor | undefined;

  const hasSelection = ({ state }: { state: EditorState }) => !isCollapsed(state.selection);

  const copySelection = async (state: EditorState) => {
    const clipboard = editor?.services.clipboard;
    if (!clipboard) return false;
    const [from, to] = orderedRange(state.doc, state.selection);
    await system.write(clipboard.serialize(state.doc, from, to));
    return true;
  };

  const copy: AsyncCommand = {
    id: 'clipboard.copy',
    label: 'Copia',
    async: true,
    isEnabled: hasSelection,
    run: async ({ state }) => {
      try {
        await copySelection(state);
      } catch {
        notify(`Il browser non permette di copiare da questo pulsante: usa ${key('C')}.`);
      }
      return null;
    },
  };

  const cut: AsyncCommand = {
    id: 'clipboard.cut',
    label: 'Taglia',
    async: true,
    isEnabled: hasSelection,
    run: async ({ state }) => {
      try {
        if (await copySelection(state)) editor?.execute('edit.deleteSelection');
      } catch {
        notify(`Il browser non permette di tagliare da questo pulsante: usa ${key('X')}.`);
      }
      return null;
    },
  };

  const paste: AsyncCommand<{ textOnly?: boolean } | undefined> = {
    id: 'clipboard.paste',
    label: 'Incolla',
    async: true,
    run: async (_ctx, args) => {
      let data: { html?: string; text?: string };
      try {
        data = await system.read();
      } catch {
        notify(`Il browser non permette di incollare da questo pulsante: usa ${key('V')}.`);
        return null;
      }
      if (!editor) return null;
      const clipboard = editor.services.clipboard;
      if (!args?.textOnly && clipboard && (data.html || data.text)) {
        if (editor.execute('edit.insertFragment', { fragment: clipboard.parse(data) })) return null;
      }
      const text = data.text ?? '';
      if (text) editor.execute('edit.insertText', { text });
      return null;
    },
  };

  return {
    id: CLIPBOARD_MODULE_ID,
    activate(ed: Editor) {
      editor = ed;
      const offs = [cut, copy, paste].map((c) => ed.commands.register(c));
      return () => {
        offs.forEach((off) => off());
        editor = undefined;
      };
    },
  };
}
