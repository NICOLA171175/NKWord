import { Editor, findParagraph, textRun, updateParagraph, type Command } from '../../src/core';

/** Editor con comandi finti che registrano le chiamate e inseriscono testo davvero. */
export function recordingEditor(ids: readonly string[] = []) {
  const editor = new Editor();
  const calls: { id: string; args: unknown }[] = [];
  const insertText: Command<{ text: string }> = {
    id: 'edit.insertText',
    label: 'Inserisci testo',
    run: ({ state }, { text }) => {
      calls.push({ id: 'edit.insertText', args: { text } });
      const p = findParagraph(state.doc, state.selection.focus.paragraphId)!.paragraph;
      const doc = updateParagraph(state.doc, { ...p, children: [...p.children, textRun(text)] });
      return { state: { ...state, doc } };
    },
  };
  editor.commands.register(insertText);
  for (const id of ids) {
    editor.commands.register({
      id,
      label: id,
      run: ({ state }, args) => {
        calls.push({ id, args });
        return { state: { ...state }, meta: { addToHistory: false } };
      },
    });
  }
  return { editor, calls };
}
