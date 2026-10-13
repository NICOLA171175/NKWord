import { describe, expect, it } from 'vitest';
import { paragraph, createDocument, textRun } from '../../src/core';
import { LANGUAGES, SpellChecker } from '../../src/spellcheck';
import { memoryStorage, miniLanguages } from './helpers';

describe('SpellChecker', () => {
  it('non segnala nulla finché il dizionario non è caricato', async () => {
    const c = new SpellChecker({ languages: miniLanguages, storage: null });
    const p = paragraph('gatto gatot');
    expect(c.settings).toEqual({ enabled: true, language: 'it', ignoreUppercase: true });
    expect(c.ready).toBe(false);
    expect(c.misspellingsIn(p)).toEqual([]);
    let changes = 0;
    c.events.on('change', () => changes++);
    expect(await c.load()).toBe(true);
    expect(changes).toBe(1);
    expect(c.misspellingsIn(p)).toEqual([{ paragraphId: p.id, start: 6, end: 11, word: 'gatot' }]);
  });

  it('trova gli errori in tutto il documento, tabelle comprese, e usa la cache per paragrafo', async () => {
    const c = new SpellChecker({ languages: miniLanguages, storage: null });
    await c.load();
    const a = paragraph([textRun('ciao '), textRun('gatot', { bold: true })]);
    const b = paragraph("'ciaoo' e 'ciao'");
    const doc = createDocument([a, b]);
    expect(c.misspellings(doc).map((m) => m.word)).toEqual(['gatot', 'ciaoo', 'e']);
    expect(c.misspellingsIn(a)).toBe(c.misspellingsIn(a));
  });

  it('Ignora tutto vale per la sessione, Aggiungi al dizionario viene salvato', async () => {
    const storage = memoryStorage();
    const c = new SpellChecker({ languages: miniLanguages, storage });
    await c.load();
    const p = paragraph('gatot nikoffice');
    expect(c.misspellingsIn(p)).toHaveLength(2);
    c.ignore('gatot');
    c.addWord('nikoffice');
    expect(c.misspellingsIn(p)).toEqual([]);
    expect(c.isCorrect('Nikoffice')).toBe(true);
    expect(JSON.parse(storage.value!)).toMatchObject({ words: ['nikoffice'], language: 'it', enabled: true });

    const again = new SpellChecker({ languages: miniLanguages, storage });
    await again.load();
    expect(again.personalWords).toEqual(['nikoffice']);
    expect(again.misspellingsIn(paragraph('gatot nikoffice')).map((m) => m.word)).toEqual(['gatot']);
    again.removeWord('nikoffice');
    expect(JSON.parse(storage.value!).words).toEqual([]);
  });

  it('cambia lingua, spegne il controllo e ricorda le impostazioni', async () => {
    const storage = memoryStorage();
    const c = new SpellChecker({ languages: miniLanguages, storage });
    await c.load();
    c.update({ language: 'xx' });
    expect(c.ready).toBe(false);
    await c.load();
    expect(c.misspellingsIn(paragraph('hello gatto')).map((m) => m.word)).toEqual(['gatto']);
    c.update({ enabled: false });
    expect(c.misspellingsIn(paragraph('hello gatto'))).toEqual([]);
    c.update({ language: 'sconosciuta' });
    expect(c.settings.language).toBe('xx');
    expect(new SpellChecker({ languages: miniLanguages, storage }).settings).toMatchObject({ language: 'xx', enabled: false });
  });

  it('sopravvive a un salvataggio rovinato o a una lingua che non si carica', async () => {
    const c = new SpellChecker({ languages: [{ id: 'it', label: 'Rotto', load: () => Promise.reject(new Error('rete')) }], storage: memoryStorage('{non json') });
    expect(await c.load()).toBe(false);
    expect(c.ready).toBe(false);
  });

  it("carica davvero i dizionari inclusi (italiano predefinito)", async () => {
    expect(LANGUAGES.map((l) => l.id)).toEqual(['it', 'en']);
    const c = new SpellChecker({ storage: null });
    expect(c.settings.language).toBe('it');
    expect(await c.load()).toBe(true);
    expect(c.misspellingsIn(paragraph("Perche l'albero è cresciuto così?")).map((m) => m.word)).toEqual(['Perche']);
    expect(c.suggest('Perche')[0]).toBe('Perché');
  });
});
