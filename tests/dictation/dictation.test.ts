// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { caret, createState, Editor, firstParagraph, paragraph, paragraphText, textRun, type DocumentModel } from '../../src/core';
import { editingModule } from '../../src/editing';
import * as dictationExports from '../../src/dictation';
import {
  createDictationModule,
  detectEnvironment,
  dictationMode,
  dictationPieces,
  errorMessage,
  mountDictation,
  ribbonItems,
  type DictationEnvironment,
  type SpeechRecognitionLike,
  type SpeechResultEventLike,
} from '../../src/dictation';
import { createDictationModule as createCore } from '../../src/dictation/module';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { extensionIsActive } from '../../src/ui/extensions';
import { doc } from '../layout/helpers';

afterEach(() => document.body.replaceChildren());

const texts = (spoken: string, before = '') =>
  dictationPieces(spoken, before).map((p) => (p.kind === 'text' ? p.text : '¶')).join('');

describe('dal parlato al testo', () => {
  it('maiuscola a inizio frase e spazi fra le frasi', () => {
    expect(texts('ciao mondo')).toBe('Ciao mondo');
    expect(texts('come stai', 'Ciao.')).toBe(' Come stai');
    expect(texts('e poi', 'Ciao')).toBe(' e poi');
    expect(texts('poi', 'Ciao ')).toBe('poi');
  });

  it('comandi di punteggiatura', () => {
    expect(texts('ciao virgola come stai punto interrogativo')).toBe('Ciao, come stai?');
    expect(texts('fine punto inizio')).toBe('Fine. Inizio');
    expect(texts('attenzione punto esclamativo')).toBe('Attenzione!');
    expect(texts('uno punto e virgola due due punti tre')).toBe('Uno; due: tre');
    expect(texts('Punto')).toBe('.');
  });

  it('a capo crea un nuovo paragrafo e riparte con la maiuscola', () => {
    expect(dictationPieces('prima riga a capo seconda riga')).toEqual([
      { kind: 'text', text: 'Prima riga' },
      { kind: 'paragraph' },
      { kind: 'text', text: 'Seconda riga' },
    ]);
    expect(texts('nuovo paragrafo ciao', 'Testo')).toBe('¶Ciao');
  });

  it('parentesi e virgolette', () => {
    expect(texts('vedi apri parentesi pagina due chiudi parentesi')).toBe('Vedi (pagina due)');
    expect(texts('disse apri virgolette basta chiudi virgolette e uscì')).toBe('Disse "basta" e uscì');
  });

  it('attacca la punteggiatura messa già dal riconoscimento', () => {
    expect(texts('ciao , come va?', 'x')).toBe(' ciao, come va?');
  });
});

describe('ambiente', () => {
  const Fake = class {} as unknown as DictationEnvironment['recognition'];
  it('nel browser usa il riconoscimento del browser se c’è', () => {
    expect(dictationMode({ recognition: Fake, desktop: null, windows: false })).toBe('web');
    expect(dictationMode({ recognition: null, desktop: null, windows: true })).toBe('none');
  });
  it('nel programma desktop usa la dettatura del sistema', () => {
    const start = async () => true;
    expect(dictationMode({ recognition: Fake, desktop: { platform: 'win32', startSystemDictation: start }, windows: true })).toBe('windows');
    expect(dictationMode({ recognition: Fake, desktop: { platform: 'darwin' }, windows: false })).toBe('mac');
  });
  it('legge la finestra', () => {
    const w = { webkitSpeechRecognition: Fake, navigator: { userAgent: 'Mozilla/5.0 (Windows NT 10.0)' } } as unknown as Window;
    expect(detectEnvironment(w)).toEqual({ recognition: Fake, desktop: null, windows: true });
  });
  it('messaggi di errore comprensibili', () => {
    expect(errorMessage('no-speech', false)).toBeNull();
    expect(errorMessage('not-allowed', false)).toMatch(/microfono/);
    expect(errorMessage('network', true)).toMatch(/Win\+H/);
  });
});

/** Riconoscimento vocale finto, pilotato dal test. */
class FakeRecognition implements SpeechRecognitionLike {
  static last: FakeRecognition | null = null;
  lang = '';
  continuous = false;
  interimResults = false;
  onresult: ((event: SpeechResultEventLike) => void) | null = null;
  onerror: ((event: { readonly error: string }) => void) | null = null;
  onend: (() => void) | null = null;
  starts = 0;
  running = false;
  constructor() {
    FakeRecognition.last = this;
  }
  start() {
    this.starts += 1;
    this.running = true;
  }
  stop() {
    this.running = false;
    this.onend?.();
  }
  abort() {
    this.stop();
  }
  say(transcript: string, isFinal = true) {
    this.onresult?.({ resultIndex: 0, results: [{ isFinal, 0: { transcript } }] });
  }
  end() {
    this.running = false;
    this.onend?.();
  }
}

function setup(start: DocumentModel = doc([paragraph([textRun('')])])) {
  const editor = new Editor();
  editor.use(editingModule);
  const p = firstParagraph(start);
  editor.reset({ ...createState(start), selection: caret(p.id, paragraphText(p).length) });
  const module = createCore({ recognition: FakeRecognition, desktop: null, windows: false });
  editor.use(module);
  return { editor, module };
}
const allText = (editor: Editor) => editor.state.doc.sections[0]!.blocks.map((b) => paragraphText(b as never));

describe('dettatura con il riconoscimento del browser', () => {
  it('scrive nel punto del cursore con i comandi vocali', async () => {
    const { editor, module } = setup(doc([paragraph([textRun('Inizio.')])]));
    await editor.executeAsync('dictation.toggle');
    const rec = FakeRecognition.last!;
    expect(rec).toMatchObject({ lang: 'it-IT', continuous: true, interimResults: true, running: true });
    expect(module.listening).toBe(true);
    expect(extensionIsActive(editor, 'dictation.toggle')).toBe(true);

    rec.say('ciao a tutti', false);
    expect(module.interim).toBe('ciao a tutti');
    expect(allText(editor)).toEqual(['Inizio.']);

    rec.say('ciao a tutti virgola come state a capo bene punto');
    expect(module.interim).toBe('');
    expect(allText(editor)).toEqual(['Inizio. Ciao a tutti, come state', 'Bene.']);
  });

  it('continua ad ascoltare quando il browser si ferma da solo', async () => {
    const { editor, module } = setup();
    await editor.executeAsync('dictation.toggle');
    const rec = FakeRecognition.last!;
    rec.end();
    expect(rec.starts).toBe(2);
    expect(module.listening).toBe(true);
  });

  it('si ferma con un altro clic su Detta', async () => {
    const { editor, module } = setup();
    await editor.executeAsync('dictation.toggle');
    await editor.executeAsync('dictation.toggle');
    expect(module.listening).toBe(false);
    expect(FakeRecognition.last!.running).toBe(false);
    expect(extensionIsActive(editor, 'dictation.toggle')).toBe(false);
  });

  it('mostra un avviso se il microfono non è consentito', async () => {
    const { editor, module } = setup();
    await editor.executeAsync('dictation.toggle');
    FakeRecognition.last!.onerror?.({ error: 'not-allowed' });
    FakeRecognition.last!.end();
    expect(module.listening).toBe(false);
    expect(module.message).toMatchObject({ kind: 'error' });
    expect(FakeRecognition.last!.starts).toBe(1);
  });

  it('indicatore: frase in corso, Esc ferma', async () => {
    const { editor, module } = setup();
    const focus = vi.fn();
    const unmount = mountDictation({ input: { focus } }, module);
    const box = document.querySelector<HTMLElement>('.dictation')!;
    expect(box.hidden).toBe(true);
    await editor.executeAsync('dictation.toggle');
    expect(box.hidden).toBe(false);
    expect(box.classList.contains('dictation-listening')).toBe(true);
    FakeRecognition.last!.say('sto parlando', false);
    expect(box.textContent).toContain('sto parlando');
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(module.listening).toBe(false);
    expect(box.hidden).toBe(true);
    unmount();
    expect(document.querySelector('.dictation')).toBeNull();
  });
});

describe('dettatura nel programma desktop', () => {
  it('su Windows apre la dettatura di Windows (Win+H) e spiega come fermarla', async () => {
    const startSystemDictation = vi.fn(async () => true);
    const module = createCore({ recognition: FakeRecognition, desktop: { platform: 'win32', startSystemDictation }, windows: true });
    const editor = new Editor();
    editor.use(module);
    const focus = vi.fn();
    module.focusDocument = focus;
    await editor.executeAsync('dictation.toggle');
    expect(startSystemDictation).toHaveBeenCalledTimes(1);
    expect(focus).toHaveBeenCalled();
    expect(module.message?.text).toMatch(/Win\+H/);
    expect(module.listening).toBe(false);
  });

  it('sul Mac spiega la Dettatura del Mac', async () => {
    const module = createCore({ recognition: null, desktop: { platform: 'darwin' }, windows: false });
    await module.toggle();
    expect(module.message?.text).toMatch(/Fn/);
  });

  it('senza riconoscimento suggerisce Chrome o Edge', async () => {
    const module = createCore({ recognition: null, desktop: null, windows: true });
    await module.toggle();
    expect(module.message).toMatchObject({ kind: 'error' });
    expect(module.message?.text).toMatch(/Chrome o Edge.*Win\+H/);
  });
});

describe('barra multifunzione', () => {
  it('pulsante Detta nella scheda Home, gruppo Voce', () => {
    expect(ribbonItems).toEqual([expect.objectContaining({ tab: 'home', group: 'Voce', commandId: 'dictation.toggle', label: 'Detta', kind: 'toggle' })]);
    expect(collectRibbonItems([dictationExports])).toHaveLength(1);
  });
  it('la UI trova il modulo con la sua vista', () => {
    const modules = collectModules([dictationExports]);
    expect(modules).toHaveLength(1);
    expect(typeof (modules[0] as { mount?: unknown }).mount).toBe('function');
    expect(createDictationModule().id).toBe('dictation');
  });
});
