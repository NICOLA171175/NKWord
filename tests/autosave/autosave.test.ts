import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Editor, createDocument, paragraph, paragraphText, type Paragraph } from '../../src/core';
import { createIoModule } from '../../src/io';
import { createAutosaveModule, MemoryRecoveryStore, type Presence, type RecoverySnapshot } from '../../src/autosave';
import { editingModule } from '../../src/editing';

const nobody: Presence = { others: async () => new Set(), dispose() {} };
const firstText = (e: Editor) => paragraphText(e.state.doc.sections[0]!.blocks[0] as Paragraph);

function setup(store = new MemoryRecoveryStore(), presence = nobody, sessionId = 'io') {
  const editor = new Editor();
  editor.use(createIoModule({ save: () => {} }));
  editor.use(editingModule);
  const autosave = createAutosaveModule({ store, presence, sessionId, delayMs: 1000, maxWaitMs: 5000 });
  editor.use(autosave);
  return { editor, autosave, store };
}

const snapshot = (id: string, text: string, savedAt: string, title = ''): RecoverySnapshot => ({
  id,
  title,
  savedAt,
  doc: JSON.stringify(createDocument([paragraph(text)])),
});

describe('salvataggio automatico', () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  it('scrive una copia di recupero poco dopo una modifica', async () => {
    const { editor, autosave, store } = setup();
    editor.execute('edit.insertText', { text: 'Ciao' });
    expect(autosave.dirty).toBe(true);
    expect(await store.list()).toHaveLength(0);
    await vi.advanceTimersByTimeAsync(1000);
    const [copy] = await store.list();
    expect(copy?.id).toBe('io');
    expect(JSON.parse(copy!.doc).sections[0].blocks[0].children[0].text).toBe('Ciao');
    expect(autosave.lastSavedAt).not.toBeNull();
  });

  it('scrivendo senza pause aggiorna comunque la copia', async () => {
    const { editor, store } = setup();
    for (let i = 0; i < 12; i++) {
      editor.execute('edit.insertText', { text: 'a' });
      await vi.advanceTimersByTimeAsync(500);
    }
    expect(await store.list()).toHaveLength(1);
  });

  it('salvando in un formato riapribile la copia si cancella', async () => {
    const { editor, autosave, store } = setup();
    editor.execute('edit.insertText', { text: 'Ciao' });
    await autosave.flush();
    expect(await store.list()).toHaveLength(1);
    await editor.executeAsync('file.export', { formatId: 'txt' });
    await vi.runAllTimersAsync();
    expect(autosave.dirty).toBe(false);
    expect(await store.list()).toHaveLength(0);
  });

  it('un documento nuovo non ha nulla da recuperare', async () => {
    const { editor, autosave, store } = setup();
    editor.execute('edit.insertText', { text: 'Ciao' });
    await autosave.flush();
    editor.execute('file.new');
    await vi.runAllTimersAsync();
    expect(await store.list()).toHaveLength(0);
  });
});

describe('recupero dopo una chiusura improvvisa', () => {
  it('propone le copie delle finestre chiuse, dalla più recente', async () => {
    const store = new MemoryRecoveryStore();
    await store.put(snapshot('vecchia', 'prima', '2026-10-12T08:00:00.000Z'));
    await store.put(snapshot('nuova', 'seconda', '2026-10-13T08:00:00.000Z', 'Relazione'));
    await store.put(snapshot('aperta', 'altra scheda', '2026-10-13T09:00:00.000Z'));
    const presence: Presence = { others: async () => new Set(['aperta']), dispose() {} };
    const { editor, autosave } = setup(store, presence);
    await autosave.ready;
    expect(autosave.recovered.map((r) => r.id)).toEqual(['nuova', 'vecchia']);
    expect(editor.canExecute('autosave.show')).toBe(true);

    expect(editor.execute('autosave.restore')).toBe(true);
    expect(firstText(editor)).toBe('seconda');
    expect(editor.history.canUndo).toBe(false);
    // La copia ripristinata ora appartiene a questa finestra.
    expect(autosave.recovered.map((r) => r.id)).toEqual(['vecchia']);
    expect(autosave.dirty).toBe(true);
    await autosave.flush();
    const ids = (await store.list()).map((s) => s.id).sort();
    expect(ids).toEqual(['aperta', 'io', 'vecchia']);
  });

  it('elimina le copie che non servono', async () => {
    const store = new MemoryRecoveryStore();
    await store.put(snapshot('a', 'uno', '2026-10-12T08:00:00.000Z'));
    await store.put(snapshot('b', 'due', '2026-10-12T09:00:00.000Z'));
    const { editor, autosave } = setup(store);
    await autosave.ready;
    await editor.executeAsync('autosave.discard', { id: 'a' });
    expect(autosave.recovered.map((r) => r.id)).toEqual(['b']);
    await editor.executeAsync('autosave.discard');
    expect(autosave.recovered).toHaveLength(0);
    expect(await store.list()).toHaveLength(0);
    expect(editor.canExecute('autosave.restore')).toBe(false);
  });
});
