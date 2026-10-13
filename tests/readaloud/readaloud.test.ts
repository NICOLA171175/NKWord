// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { caret, createDocument, createState, Editor, paragraph, type Selection } from '../../src/core';
import { createApproxMeasurer, createLayoutModule } from '../../src/layout';
import {
  clampRate,
  createReadAloudCore,
  mountReadAloud,
  pickVoice,
  rateLabel,
  readingChunks,
  ribbonItems,
  sortVoices,
  spokenText,
  wordAt,
  wordStarts,
  type ReadAloudEngine,
  type UtteranceLike,
  type VoiceLike,
} from '../../src/readaloud';
import * as readExports from '../../src/readaloud';
import { collectMarks } from '../../src/ui/extensions';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';

afterEach(() => document.body.replaceChildren());

const VOICES: VoiceLike[] = [
  { name: 'English', lang: 'en-US', default: true, localService: true },
  { name: 'Google italiano', lang: 'it-IT', localService: false },
  { name: 'Microsoft Elsa - Italian (Italy)', lang: 'it-IT', localService: true },
];

/** Sintesi vocale finta: registra le frasi e lascia simulare parole e fine. */
function fakeEngine(voices = VOICES) {
  const spoken: UtteranceLike[] = [];
  let cancels = 0;
  const engine: ReadAloudEngine = {
    speech: {
      speak: (u) => spoken.push(u),
      cancel: () => {
        cancels += 1;
      },
      getVoices: () => voices,
    },
    createUtterance: (text) => ({ text, lang: '', rate: 1, voice: null, onboundary: null, onend: null, onerror: null }),
  };
  return {
    engine,
    spoken,
    get cancels() {
      return cancels;
    },
    last: () => spoken.at(-1)!,
  };
}

function memoryStorage(initial: Record<string, string> = {}) {
  const data = new Map(Object.entries(initial));
  return { getItem: (k: string) => data.get(k) ?? null, setItem: (k: string, v: string) => void data.set(k, v), data };
}

function setup(texts = ['Ciao mondo bello.', '', 'Secondo paragrafo qui.', 'Terzo.'], selection?: (ids: string[]) => Selection) {
  const doc = createDocument(texts.map((t) => paragraph(t)));
  const ids = doc.sections[0]!.blocks.map((b) => b.id);
  const state = createState(doc);
  const editor = new Editor(selection ? { ...state, selection: selection(ids) } : state);
  const fake = fakeEngine();
  const storage = memoryStorage();
  const read = createReadAloudCore({ engine: fake.engine, storage });
  editor.use(read);
  return { editor, read, fake, ids, storage };
}

describe('testo', () => {
  it('trova le parole', () => {
    expect(wordAt('Ciao, mondo!', 0)).toEqual([0, 4]);
    expect(wordAt('Ciao, mondo!', 2)).toEqual([0, 4]);
    expect(wordAt('Ciao, mondo!', 4)).toEqual([6, 11]);
    expect(wordAt('Ciao, mondo!', 11)).toBeNull();
    expect(wordAt("l'albero è verde", 0)).toEqual([0, 8]);
    expect(wordStarts('Ciao, bel mondo')).toEqual([0, 6, 10]);
  });

  it('mantiene le posizioni del testo pronunciato', () => {
    expect(spokenText('a\tb\nc￼d')).toBe('a b c d');
  });

  it('legge dal cursore fino alla fine, saltando i paragrafi vuoti', () => {
    const doc = createDocument([paragraph('Uno due tre'), paragraph(''), paragraph('Quattro')]);
    const [a, , c] = doc.sections[0]!.blocks.map((b) => b.id);
    const chunks = readingChunks(doc, caret(a!, 5));
    expect(chunks.map((ch) => [ch.paragraphId, ch.start, ch.end])).toEqual([
      [a, 4, 11],
      [c, 0, 7],
    ]);
  });

  it('con una selezione legge solo quella', () => {
    const doc = createDocument([paragraph('Uno due tre'), paragraph('Quattro cinque')]);
    const [a, b] = doc.sections[0]!.blocks.map((x) => x.id);
    const chunks = readingChunks(doc, { anchor: { paragraphId: b!, offset: 7 }, focus: { paragraphId: a!, offset: 4 } });
    expect(chunks.map((ch) => [ch.start, ch.end])).toEqual([
      [4, 11],
      [0, 7],
    ]);
  });

  it('sceglie una voce italiana installata sul computer', () => {
    expect(pickVoice(VOICES)?.name).toBe('Microsoft Elsa - Italian (Italy)');
    expect(pickVoice(VOICES, 'English')?.name).toBe('English');
    expect(pickVoice(VOICES, 'sparita')?.name).toBe('Microsoft Elsa - Italian (Italy)');
    expect(pickVoice([VOICES[0]!])?.name).toBe('English');
    expect(pickVoice([])).toBeNull();
    expect(sortVoices(VOICES).map((v) => v.lang)).toEqual(['it-IT', 'it-IT', 'en-US']);
  });

  it('limita la velocità', () => {
    expect(clampRate(0.1)).toBe(0.5);
    expect(clampRate(5)).toBe(2);
    expect(clampRate(1.26)).toBe(1.3);
    expect(clampRate(Number.NaN)).toBe(1);
    expect(rateLabel(1.5)).toBe('1,5×');
  });
});

describe('modulo', () => {
  it('si carica da solo con il pulsante in Revisione > Voce', () => {
    expect(collectModules([readExports]).map((m) => m.id)).toEqual(['readaloud']);
    expect(collectRibbonItems([readExports])).toEqual(ribbonItems);
    expect(ribbonItems[0]).toMatchObject({ tab: 'review', group: 'Voce', commandId: 'readaloud.toggle' });
  });

  it('legge un paragrafo alla volta evidenziando la parola', async () => {
    const { editor, read, fake, ids } = setup();
    await editor.executeAsync('readaloud.toggle');
    expect(read.status).toBe('playing');
    const u = fake.last();
    expect(u.text).toBe('Ciao mondo bello.');
    expect(u.lang).toBe('it-IT');
    expect(u.voice?.name).toBe('Microsoft Elsa - Italian (Italy)');
    expect(read.word).toEqual({ paragraphId: ids[0], start: 0, end: 4 });
    u.onboundary!({ name: 'word', charIndex: 5 });
    expect(read.word).toEqual({ paragraphId: ids[0], start: 5, end: 10 });
    expect(collectMarks(editor)).toEqual([expect.objectContaining({ paragraphId: ids[0], start: 5, end: 10, kind: 'highlight' })]);
    u.onend!();
    expect(fake.last().text).toBe('Secondo paragrafo qui.');
    fake.last().onend!();
    fake.last().onend!();
    expect(read.status).toBe('stopped');
    expect(collectMarks(editor)).toEqual([]);
    // Il cursore resta sull'ultima parola letta e Annulla non cambia.
    expect(editor.state.selection.focus).toEqual({ paragraphId: ids[3], offset: 0 });
    expect(editor.history.canUndo).toBe(false);
  });

  it('pausa e ripresa dalla parola dove si era fermata', async () => {
    const { editor, read, fake } = setup();
    await editor.executeAsync('readaloud.toggle');
    fake.last().onboundary!({ name: 'word', charIndex: 11 });
    await editor.executeAsync('readaloud.playPause');
    expect(read.status).toBe('paused');
    const count = fake.spoken.length;
    await editor.executeAsync('readaloud.playPause');
    expect(read.status).toBe('playing');
    expect(fake.spoken.length).toBe(count + 1);
    expect(fake.last().text).toBe('bello.');
    // Gli eventi della frase annullata non contano più.
    fake.spoken[0]!.onend!();
    expect(fake.last().text).toBe('bello.');
  });

  it('Avanti e Indietro saltano di paragrafo', async () => {
    const { editor, fake } = setup();
    await editor.executeAsync('readaloud.toggle');
    await editor.executeAsync('readaloud.next');
    expect(fake.last().text).toBe('Secondo paragrafo qui.');
    fake.last().onboundary!({ name: 'word', charIndex: 8 });
    await editor.executeAsync('readaloud.previous');
    expect(fake.last().text).toBe('Secondo paragrafo qui.');
    await editor.executeAsync('readaloud.previous');
    expect(fake.last().text).toBe('Ciao mondo bello.');
  });

  it('cambiare velocità o voce riparte dalla parola attuale e viene ricordato', async () => {
    const { editor, read, fake, storage } = setup();
    await editor.executeAsync('readaloud.toggle');
    fake.last().onboundary!({ name: 'word', charIndex: 5 });
    await editor.executeAsync('readaloud.rate', { rate: 1.5 });
    expect(fake.last()).toMatchObject({ text: 'mondo bello.', rate: 1.5 });
    await editor.executeAsync('readaloud.voice', { name: 'Google italiano' });
    expect(fake.last().voice?.name).toBe('Google italiano');
    expect(JSON.parse(storage.data.get('nkoffice.readaloud')!)).toEqual({ rate: 1.5, voice: 'Google italiano' });
    expect(read.voiceName).toBe('Google italiano');
    const again = createReadAloudCore({ engine: fakeEngine().engine, storage });
    expect(again.rate).toBe(1.5);
    expect(again.voiceName).toBe('Google italiano');
  });

  it('con una selezione legge solo quella', async () => {
    const { editor, fake } = setup(['Uno due tre', 'Quattro cinque'], ([a, b]) => ({
      anchor: { paragraphId: a!, offset: 4 },
      focus: { paragraphId: b!, offset: 7 },
    }));
    await editor.executeAsync('readaloud.toggle');
    expect(fake.last().text).toBe('due tre');
    fake.last().onend!();
    expect(fake.last().text).toBe('Quattro');
  });

  it('si ferma se il testo cambia e Esc la interrompe', async () => {
    const { editor, read, ids } = setup();
    await editor.executeAsync('readaloud.toggle');
    const doc = createDocument([paragraph('Altro')]);
    editor.apply({ state: { ...editor.state, doc, selection: caret(doc.sections[0]!.blocks[0]!.id, 0) } });
    expect(read.status).toBe('stopped');
    expect(editor.canExecute('readaloud.stop')).toBe(false);
    expect(ids.length).toBe(4);
  });

  it('un clic altrove durante la lettura riparte da lì', async () => {
    const { editor, fake, ids } = setup();
    await editor.executeAsync('readaloud.toggle');
    editor.apply({ state: { ...editor.state, selection: caret(ids[2]!, 8) }, meta: { addToHistory: false } });
    expect(fake.last().text).toBe('paragrafo qui.');
  });

  it('avvisa se il browser non sa leggere', async () => {
    const editor = new Editor();
    const read = createReadAloudCore({ engine: null, storage: null });
    editor.use(read);
    await editor.executeAsync('readaloud.toggle');
    expect(read.status).toBe('stopped');
    expect(read.message).toMatch(/non sa leggere/);
  });

  it('avvisa se la voce manca', async () => {
    const { editor, read, fake } = setup();
    await editor.executeAsync('readaloud.toggle');
    fake.last().onerror!({ error: 'synthesis-unavailable' });
    expect(read.status).toBe('stopped');
    expect(read.message).toMatch(/Nessuna voce/);
  });
});

describe('barra', () => {
  it('compare durante la lettura con pausa, velocità e voci', async () => {
    const { editor, read } = setup();
    editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
    const off = mountReadAloud({}, read);
    const bar = document.querySelector<HTMLElement>('.readaloud-bar')!;
    expect(bar.hidden).toBe(true);
    await editor.executeAsync('readaloud.toggle');
    expect(bar.hidden).toBe(false);
    const play = bar.querySelector<HTMLButtonElement>('.readaloud-play')!;
    expect(play.title).toBe('Pausa');
    play.click();
    expect(read.status).toBe('paused');
    expect(play.title).toBe('Riproduci');
    const options = Array.from(bar.querySelectorAll('.readaloud-voice option')).map((o) => o.textContent);
    expect(options).toEqual(['Automatica (italiana)', 'Google italiano (it-IT)', 'Microsoft Elsa (it-IT)', 'English (en-US)']);
    const rate = bar.querySelector<HTMLInputElement>('.readaloud-rate')!;
    rate.value = '1.7';
    rate.dispatchEvent(new Event('change'));
    expect(read.rate).toBe(1.7);
    bar.querySelector<HTMLButtonElement>('.readaloud-close')!.click();
    expect(read.status).toBe('stopped');
    expect(bar.hidden).toBe(true);
    off();
    expect(document.querySelector('.readaloud-bar')).toBeNull();
  });
});
