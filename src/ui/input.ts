/**
 * Input da tastiera. Un `<textarea>` nascosto, posizionato sul cursore,
 * riceve il focus: i tasti con una scorciatoia diventano comandi, il testo
 * digitato diventa `edit.insertText`. Durante una composizione IME (cinese,
 * giapponese, accenti con tasti morti) il testo resta nel textarea, dove il
 * sistema mostra l'anteprima, e viene inserito solo a `compositionend`.
 */

import { orderedRange, isCollapsed, type Editor } from '../core';
import { Keymap, type KeyLike } from './keymap';

/**
 * Prova le scorciatoie in ordine ed esegue la prima eseguibile. Restituisce
 * true se ha eseguito un comando.
 */
export function runFirstExecutable(editor: Editor, bindings: readonly { commandId: string; args?: unknown }[]): boolean {
  for (const b of bindings) {
    if (!editor.commands.has(b.commandId) || !editor.canExecute(b.commandId, b.args)) continue;
    void editor.executeAsync(b.commandId, b.args);
    return true;
  }
  return false;
}

export interface InputControllerOptions {
  readonly keymap: Keymap;
  /** Tasti gestiti prima della mappa (es. Esc per uscire da intestazione). Restituisce true se gestito. */
  readonly onKey?: (event: KeyLike) => boolean;
}

export class InputController {
  private composing = false;
  private readonly cleanup: (() => void)[] = [];

  constructor(
    private readonly editor: Editor,
    readonly element: HTMLTextAreaElement,
    private readonly options: InputControllerOptions,
  ) {
    element.setAttribute('autocomplete', 'off');
    element.setAttribute('autocorrect', 'off');
    element.setAttribute('autocapitalize', 'off');
    element.setAttribute('spellcheck', 'false');
    element.setAttribute('aria-label', 'Documento');
    this.listen('keydown', (e) => this.onKeyDown(e as KeyboardEvent));
    this.listen('input', (e) => this.onInput(e as InputEvent));
    this.listen('compositionstart', () => {
      this.composing = true;
      element.classList.add('composing');
    });
    this.listen('compositionend', (e) => this.onCompositionEnd(e as CompositionEvent));
    this.listen('paste', (e) => this.onPaste(e as ClipboardEvent));
    this.listen('copy', (e) => this.onCopy(e as ClipboardEvent, false));
    this.listen('cut', (e) => this.onCopy(e as ClipboardEvent, true));
  }

  get isComposing(): boolean {
    return this.composing;
  }

  focus(): void {
    this.element.focus({ preventScroll: true });
  }

  destroy(): void {
    for (const c of this.cleanup.splice(0)) c();
  }

  private listen(type: string, handler: (e: Event) => void): void {
    this.element.addEventListener(type, handler);
    this.cleanup.push(() => this.element.removeEventListener(type, handler));
  }

  private onKeyDown(e: KeyboardEvent): void {
    // Durante la composizione i tasti (Invio, frecce...) appartengono all'IME.
    if (this.composing || e.isComposing || e.keyCode === 229) return;
    if (this.options.onKey?.(e)) {
      e.preventDefault();
      return;
    }
    const bindings = this.options.keymap.lookup(e);
    if (bindings.length === 0) return;
    // Il tasto ha una scorciatoia: il comportamento del browser non serve
    // anche se il comando non è (ancora) disponibile.
    e.preventDefault();
    runFirstExecutable(this.editor, bindings);
  }

  private onInput(e: InputEvent): void {
    if (this.composing || e.isComposing) return;
    this.flush();
  }

  private onCompositionEnd(e: CompositionEvent): void {
    this.composing = false;
    this.element.classList.remove('composing');
    if (!this.flush() && e.data) this.insertText(e.data);
    this.element.value = '';
  }

  /** Inserisce il contenuto del textarea e lo svuota. */
  private flush(): boolean {
    const text = this.element.value;
    this.element.value = '';
    if (!text) return false;
    this.insertText(text);
    return true;
  }

  private insertText(text: string): void {
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    lines.forEach((line, i) => {
      if (i > 0) runFirstExecutable(this.editor, [{ commandId: 'edit.splitParagraph' }]);
      const parts = line.split('\t');
      parts.forEach((part, j) => {
        if (j > 0) runFirstExecutable(this.editor, [{ commandId: 'edit.insertTab' }]);
        if (part) runFirstExecutable(this.editor, [{ commandId: 'edit.insertText', args: { text: part } }]);
      });
    });
  }

  private onPaste(e: ClipboardEvent): void {
    const data = e.clipboardData;
    if (!data) return;
    e.preventDefault();
    const html = data.getData('text/html') || undefined;
    const text = data.getData('text/plain') || undefined;
    const clipboard = this.editor.services.clipboard;
    if (clipboard && (html || text) && this.editor.commands.has('edit.insertFragment')) {
      const fragment = clipboard.parse({ html, text });
      if (runFirstExecutable(this.editor, [{ commandId: 'edit.insertFragment', args: { fragment } }])) return;
    }
    if (text) this.insertText(text);
  }

  private onCopy(e: ClipboardEvent, cut: boolean): void {
    const { doc, selection } = this.editor.state;
    const clipboard = this.editor.services.clipboard;
    if (!clipboard || !e.clipboardData || isCollapsed(selection)) return;
    const [from, to] = orderedRange(doc, selection);
    const { html, text } = clipboard.serialize(doc, from, to);
    e.preventDefault();
    e.clipboardData.setData('text/html', html);
    e.clipboardData.setData('text/plain', text);
    if (cut) runFirstExecutable(this.editor, [{ commandId: 'edit.deleteSelection' }]);
  }
}
