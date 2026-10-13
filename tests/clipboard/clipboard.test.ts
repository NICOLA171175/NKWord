import { describe, expect, it } from 'vitest';
import { createDocument, createState, Editor, paragraph, paragraphText, textRun, type Paragraph } from '../../src/core';
import { createClipboardModule, ribbonItems, type SystemClipboard } from '../../src/clipboard';
import { editingModule } from '../../src/editing';
import { ioModule } from '../../src/io';

function setup(system: SystemClipboard, notes: string[] = []) {
  const p = paragraph([textRun('uno '), textRun('due', { bold: true })]);
  const editor = new Editor(createState(createDocument([p])));
  editor.use(editingModule);
  editor.use(ioModule);
  editor.use(createClipboardModule({ system, notify: (m) => notes.push(m) }));
  const select = (a: number, b: number) =>
    editor.execute('edit.setSelection', {
      selection: { anchor: { paragraphId: p.id, offset: a }, focus: { paragraphId: p.id, offset: b } },
    });
  return { editor, select };
}
const text = (editor: Editor) => paragraphText(editor.state.doc.sections[0]!.blocks[0] as Paragraph);

function memory() {
  let data: { html?: string; text?: string } = {};
  const system: SystemClipboard = {
    write: async (d) => void (data = d),
    read: async () => data,
  };
  return { system, get data() { return data; } };
}

describe('pulsanti Incolla, Taglia e Copia', () => {
  it('Copia mette HTML e testo negli appunti; Incolla mantiene la formattazione', async () => {
    const clip = memory();
    const { editor, select } = setup(clip.system);
    expect(editor.canExecute('clipboard.copy')).toBe(false);
    select(4, 7);
    await editor.executeAsync('clipboard.copy');
    expect(clip.data.text).toBe('due');
    expect(clip.data.html).toContain('due');
    select(7, 7);
    await editor.executeAsync('clipboard.paste');
    expect(text(editor)).toBe('uno duedue');
    const p = editor.state.doc.sections[0]!.blocks[0] as Paragraph;
    expect(p.children.at(-1)).toMatchObject({ format: { bold: true } });
  });

  it('Taglia copia e cancella; Incolla solo testo non porta il formato', async () => {
    const clip = memory();
    const { editor, select } = setup(clip.system);
    select(4, 7);
    await editor.executeAsync('clipboard.cut');
    expect(text(editor)).toBe('uno ');
    select(0, 0);
    await editor.executeAsync('clipboard.paste', { textOnly: true });
    expect(text(editor)).toBe('dueuno ');
    const p = editor.state.doc.sections[0]!.blocks[0] as Paragraph;
    expect(p.children[0]).toMatchObject({ text: 'dueuno ', format: {} });
  });

  it('se il browser nega gli appunti spiega di usare la tastiera', async () => {
    const notes: string[] = [];
    const denied: SystemClipboard = {
      write: () => Promise.reject(new Error('no')),
      read: () => Promise.reject(new Error('no')),
    };
    const { editor } = setup(denied, notes);
    await editor.executeAsync('clipboard.paste');
    expect(notes[0]).toMatch(/\+V/);
    expect(text(editor)).toBe('uno due');
  });

  it('i pulsanti stanno nel gruppo Appunti della scheda Home', () => {
    expect(ribbonItems.map((i) => i.label)).toEqual(['Incolla', 'Incolla solo testo', 'Taglia', 'Copia']);
    expect(ribbonItems.every((i) => i.tab === 'home' && i.group === 'Appunti')).toBe(true);
  });
});
