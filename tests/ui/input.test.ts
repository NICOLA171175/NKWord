// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { firstParagraph, paragraphText } from '../../src/core';
import { InputController } from '../../src/ui/input';
import { Keymap, defaultKeybindings } from '../../src/ui/keymap';
import { recordingEditor } from './helpers';

function setup(ids: string[] = []) {
  const { editor, calls } = recordingEditor(ids);
  const textarea = document.createElement('textarea');
  document.body.append(textarea);
  const input = new InputController(editor, textarea, { keymap: new Keymap(defaultKeybindings(false), false) });
  const key = (key: string, init: KeyboardEventInit = {}) => {
    const e = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init });
    textarea.dispatchEvent(e);
    return e;
  };
  const type = (text: string) => {
    textarea.value += text;
    textarea.dispatchEvent(new InputEvent('input', { data: text, inputType: 'insertText' }));
  };
  const text = () => paragraphText(firstParagraph(editor.state.doc));
  return { editor, calls, textarea, input, key, type, text };
}

afterEach(() => document.body.replaceChildren());

describe('InputController', () => {
  it('trasforma la digitazione in edit.insertText e svuota il textarea', () => {
    const t = setup();
    t.type('c');
    t.type('iao');
    expect(t.text()).toBe('ciao');
    expect(t.textarea.value).toBe('');
  });

  it('esegue la scorciatoia e blocca il comportamento del browser', () => {
    const t = setup(['format.bold']);
    const e = t.key('b', { ctrlKey: true, code: 'KeyB' });
    expect(e.defaultPrevented).toBe(true);
    expect(t.calls.map((c) => c.id)).toEqual(['format.bold']);
  });

  it('blocca il tasto anche se il comando non è ancora registrato', () => {
    const t = setup();
    expect(t.key('i', { ctrlKey: true, code: 'KeyI' }).defaultPrevented).toBe(true);
    expect(t.key('x').defaultPrevented).toBe(false);
  });

  it('prova le alternative: Tab rientra l\'elenco solo se possibile', () => {
    const t = setup(['edit.insertTab']);
    t.key('Tab');
    expect(t.calls.map((c) => c.id)).toEqual(['edit.insertTab']);
  });

  it('passa gli argomenti della scorciatoia', () => {
    const t = setup(['edit.moveCaret']);
    t.key('ArrowLeft', { shiftKey: true, ctrlKey: true });
    expect(t.calls[0]).toEqual({ id: 'edit.moveCaret', args: { direction: 'wordLeft', extend: true } });
  });

  it('IME: inserisce solo il testo composto, una volta', () => {
    const t = setup(['edit.splitParagraph']);
    t.textarea.dispatchEvent(new CompositionEvent('compositionstart', { data: '' }));
    expect(t.input.isComposing).toBe(true);
    for (const partial of ['n', 'ni', 'にほ']) {
      t.textarea.value = partial;
      t.textarea.dispatchEvent(new InputEvent('input', { data: partial, isComposing: true }));
    }
    // Invio conferma la composizione: non deve dividere il paragrafo.
    t.key('Enter', { isComposing: true, keyCode: 229 } as KeyboardEventInit);
    t.textarea.value = '日本';
    t.textarea.dispatchEvent(new CompositionEvent('compositionend', { data: '日本' }));
    // Safari manda un input dopo compositionend.
    t.textarea.dispatchEvent(new InputEvent('input', { data: '日本' }));
    expect(t.text()).toBe('日本');
    expect(t.calls.filter((c) => c.id === 'edit.splitParagraph')).toHaveLength(0);
    expect(t.input.isComposing).toBe(false);
  });

  it('incolla testo semplice dividendo righe e tabulazioni', () => {
    const t = setup(['edit.splitParagraph', 'edit.insertTab']);
    const e = new Event('paste', { cancelable: true }) as ClipboardEvent;
    const data = new Map([['text/plain', 'a\tb\r\nc']]);
    Object.defineProperty(e, 'clipboardData', { value: { getData: (k: string) => data.get(k) ?? '' } });
    t.textarea.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(true);
    expect(t.calls.map((c) => c.id)).toEqual(['edit.insertText', 'edit.insertTab', 'edit.insertText', 'edit.splitParagraph', 'edit.insertText']);
  });
});
