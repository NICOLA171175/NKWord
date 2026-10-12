// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { Editor, firstParagraph, paragraphText } from '../../src/core';
import { createApp } from '../../src/ui/shell';
import { countWords } from '../../src/ui/status-bar';
import { recordingEditor } from './helpers';

afterEach(() => document.body.replaceChildren());

describe('shell', () => {
  it('monta barra multifunzione, documento e barra di stato', () => {
    const root = document.createElement('div');
    document.body.append(root);
    const app = createApp(root, new Editor());
    expect(root.querySelector('.ribbon')).not.toBeNull();
    expect(root.querySelector('.page')).not.toBeNull();
    expect(root.querySelector('.status-bar')!.textContent).toContain('0 parole');
    expect(document.activeElement).toBe(app.input.element);
    expect(document.title).toBe('Documento1 - nikoffice');
    app.destroy();
    expect(root.children).toHaveLength(0);
    expect(app.editor.commands.has('view.zoom')).toBe(false);
  });

  it('digitare aggiorna documento e conteggio parole', () => {
    const { editor } = recordingEditor();
    const root = document.createElement('div');
    document.body.append(root);
    const app = createApp(root, editor);
    const sink = app.input.element;
    sink.value = 'due parole';
    sink.dispatchEvent(new InputEvent('input'));
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('due parole');
    expect(root.querySelector('.status-words')!.textContent).toBe('2 parole');
  });

  it('Esc chiude la modifica di intestazione', async () => {
    const root = document.createElement('div');
    document.body.append(root);
    const app = createApp(root, new Editor());
    await app.editor.executeAsync('page.headerFooter', { part: 'header', edit: true });
    expect(root.querySelector<HTMLElement>('.hf-bar')!.hidden).toBe(false);
    app.input.element.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', cancelable: true }));
    await Promise.resolve();
    expect(app.ui.view.state.headerFooter).toBeNull();
    expect(root.querySelector<HTMLElement>('.hf-bar')!.hidden).toBe(true);
  });

  it('conta le parole ignorando la punteggiatura isolata', () => {
    const { editor } = recordingEditor();
    editor.execute('edit.insertText', { text: 'Ciao, mondo - è un test 2026' });
    expect(countWords(editor.state.doc)).toBe(6);
  });
});
