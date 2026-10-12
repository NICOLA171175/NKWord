import { describe, expect, it } from 'vitest';
import { COMMAND_IDS, COMMAND_OWNERS, Editor } from '../../src/core';
import { editingModule } from '../../src/editing';

describe('editingModule', () => {
  it('registra esattamente i comandi di proprietà del modulo editing', () => {
    const editor = new Editor();
    const before = new Set(editor.commands.list().map((c) => c.id));
    const dispose = editor.use(editingModule);
    const added = editor.commands.list().map((c) => c.id).filter((id) => !before.has(id));
    const owned = COMMAND_IDS.filter((id) => COMMAND_OWNERS[id] === 'editing');
    expect(added.sort()).toEqual([...owned].sort());

    dispose();
    for (const id of owned) expect(editor.commands.has(id)).toBe(false);
    expect(editor.commands.has('edit.undo')).toBe(true);
  });

  it('ogni comando ha un\'etichetta in italiano', () => {
    const editor = new Editor();
    editor.use(editingModule);
    for (const c of editor.commands.list()) expect(c.label.length).toBeGreaterThan(0);
  });
});
