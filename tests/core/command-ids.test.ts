import { describe, expect, it } from 'vitest';
import { COMMAND_IDS, COMMAND_OWNERS, Editor } from '../../src/core';

describe('catalogo comandi', () => {
  it('usa id nel formato area.azione', () => {
    for (const id of COMMAND_IDS) expect(id).toMatch(/^[a-z]+\.[a-zA-Z]+$/);
  });
  it('i comandi del core sono registrati dall\'editor', () => {
    const editor = new Editor();
    for (const id of COMMAND_IDS.filter((id) => COMMAND_OWNERS[id] === 'core')) {
      expect(editor.commands.has(id)).toBe(true);
    }
  });
  it('un comando non ancora registrato non è eseguibile', () => {
    expect(new Editor().canExecute('format.bold')).toBe(false);
  });
});
