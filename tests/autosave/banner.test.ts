// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';
import { Editor, createDocument, paragraph } from '../../src/core';
import { createAutosaveModule, LocalStorageRecoveryStore, MemoryRecoveryStore, type Presence } from '../../src/autosave';

const nobody: Presence = { others: async () => new Set(), dispose() {} };

describe('avviso Documenti recuperati', () => {
  it('mostra le copie e le ripristina o le elimina', async () => {
    const store = new MemoryRecoveryStore();
    const doc = JSON.stringify(createDocument([paragraph('perso')]));
    await store.put({ id: 'x', title: 'Lettera', savedAt: '2026-10-13T08:00:00.000Z', doc });
    await store.put({ id: 'y', title: '', savedAt: '2026-10-12T08:00:00.000Z', doc });
    const editor = new Editor();
    const autosave = createAutosaveModule({ store, presence: nobody, sessionId: 'io' });
    editor.use(autosave);
    const ribbon = document.createElement('div');
    const statusBar = document.createElement('div');
    document.body.append(ribbon, statusBar);
    autosave.mount({ editor, ribbon: { element: ribbon }, statusBar: { element: statusBar } });
    await autosave.ready;

    const banner = document.querySelector<HTMLElement>('.autosave-banner')!;
    expect(banner.hidden).toBe(false);
    expect(banner.textContent).toContain('Documenti recuperati');
    expect(Array.from(banner.querySelectorAll('.autosave-title')).map((e) => e.textContent)).toEqual(['Lettera', 'Documento senza titolo']);

    banner.querySelector<HTMLButtonElement>('.autosave-restore')!.click();
    expect(editor.state.doc.sections[0]!.blocks[0]).toMatchObject({ children: [{ text: 'perso' }] });
    expect(banner.querySelectorAll('.autosave-item')).toHaveLength(1);
    expect(statusBar.textContent).toContain('Modifiche non salvate');

    banner.querySelector<HTMLButtonElement>('.autosave-close')!.click();
    expect(banner.hidden).toBe(true);
    editor.execute('autosave.show');
    expect(banner.hidden).toBe(false);

    banner.querySelector<HTMLButtonElement>('.autosave-discard')!.click();
    await new Promise((r) => setTimeout(r, 0));
    expect(banner.hidden).toBe(true);
  });

  it('archivio su localStorage', async () => {
    const store = new LocalStorageRecoveryStore(localStorage);
    await store.put({ id: 'a', title: 't', savedAt: '2026-10-13T08:00:00.000Z', doc: '{}' });
    localStorage.setItem('altro', 'x');
    expect((await store.list()).map((s) => s.id)).toEqual(['a']);
    await store.remove('a');
    expect(await store.list()).toEqual([]);
  });
});
