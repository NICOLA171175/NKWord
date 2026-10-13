// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { Editor, createState, createDocument, paragraph, paragraphText, firstParagraph, type Command } from '../../src/core';
import { createAutosaveModule, MemoryRecoveryStore, type Presence } from '../../src/autosave';

// Controllo dei comandi (File > Nuovo / Apri): con modifiche non salvate si chiede conferma.
const nobody: Presence = { others: async () => new Set(), dispose() {} };

afterEach(() => vi.restoreAllMocks());

/** happy-dom non ha window.confirm: lo si simula. */
function mockConfirm(answer: boolean) {
  const confirm = vi.fn((_message?: string) => answer);
  Object.assign(window, { confirm });
  return confirm;
}

function setup() {
  const editor = new Editor();
  const replace = (label: string): Command<void> => ({ id: label, label, run: () => ({ state: createState(createDocument([paragraph('nuovo')])), meta: { addToHistory: false, label } }) });
  editor.commands.register(replace('file.new'));
  editor.commands.register({ id: 'scrivi', label: 'scrivi', run: () => ({ state: createState(createDocument([paragraph('importante')])) }) });
  const autosave = createAutosaveModule({ store: new MemoryRecoveryStore(), presence: nobody, sessionId: 'io' });
  editor.use(autosave);
  const ribbon = document.createElement('div');
  document.body.append(ribbon);
  const unmount = autosave.mount({ editor, ribbon: { element: ribbon } });
  return { editor, unmount };
}

const text = (e: Editor) => paragraphText(firstParagraph(e.state.doc));

describe('File > Nuovo con modifiche non salvate', () => {
  it('senza modifiche non chiede niente', () => {
    const confirm = mockConfirm(true);
    const { editor } = setup();
    editor.execute('file.new');
    expect(confirm).not.toHaveBeenCalled();
    expect(text(editor)).toBe('nuovo');
  });

  it('chiede conferma: Annulla tiene il documento, OK lo sostituisce', () => {
    const confirm = mockConfirm(false);
    const { editor, unmount } = setup();
    editor.execute('scrivi');
    editor.execute('file.new');
    expect(confirm).toHaveBeenCalledOnce();
    expect(confirm.mock.calls[0]![0]).toContain('modifiche non salvate');
    expect(text(editor)).toBe('importante');

    confirm.mockImplementation(() => true);
    editor.execute('file.new');
    expect(text(editor)).toBe('nuovo');

    // Smontato l'avviso, il controllo non resta attivo.
    unmount();
    editor.execute('scrivi');
    confirm.mockClear();
    editor.execute('file.new');
    expect(confirm).not.toHaveBeenCalled();
  });
});
