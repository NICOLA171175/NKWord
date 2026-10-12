import { describe, expect, it, vi } from 'vitest';
import { COMMAND_IDS, COMMAND_OWNERS, Editor } from '../../src/core';
import { createUiModule } from '../../src/ui/module';
import { clampZoom } from '../../src/ui/view-state';

describe('modulo ui', () => {
  it('registra esattamente i comandi di proprietà ui', () => {
    const editor = new Editor();
    const before = new Set(editor.commands.list().map((c) => c.id));
    const dispose = editor.use(createUiModule({ print: () => {} }));
    const added = editor.commands.list().map((c) => c.id).filter((id) => !before.has(id)).sort();
    expect(added).toEqual(COMMAND_IDS.filter((id) => COMMAND_OWNERS[id] === 'ui').sort());
    dispose();
    expect(editor.commands.has('view.zoom')).toBe(false);
  });

  it('cambia la vista senza toccare documento e cronologia', async () => {
    const editor = new Editor();
    const ui = createUiModule({ print: () => {} });
    editor.use(ui);
    const doc = editor.state.doc;
    let viewChanges = 0;
    ui.view.events.on('change', () => viewChanges++);

    await editor.executeAsync('view.zoom', { zoom: 1.5 });
    expect(ui.view.state.zoom).toBe(1.5);
    expect(ui.isActive('view.zoom', { zoom: 1.5 })).toBe(true);
    await editor.executeAsync('view.zoom', { zoom: 99 });
    expect(ui.view.state.zoom).toBe(5);
    expect(editor.canExecute('view.zoom', { zoom: Number.NaN })).toBe(false);

    await editor.executeAsync('view.toggleFormattingMarks');
    expect(ui.isActive('view.toggleFormattingMarks')).toBe(true);
    await editor.executeAsync('view.toggleFormattingMarks');
    expect(ui.view.state.formattingMarks).toBe(false);

    expect(editor.state.doc).toBe(doc);
    expect(editor.history.canUndo).toBe(false);
    expect(viewChanges).toBe(4);
  });

  it('entra ed esce da intestazione e piè di pagina', async () => {
    const editor = new Editor();
    const ui = createUiModule();
    editor.use(ui);
    await editor.executeAsync('page.headerFooter', { part: 'header', edit: true });
    expect(ui.view.state.headerFooter).toBe('header');
    expect(ui.isActive('page.headerFooter', { part: 'header', edit: true })).toBe(true);
    // Chiudere il piè di pagina non chiude l'intestazione.
    await editor.executeAsync('page.headerFooter', { part: 'footer', edit: false });
    expect(ui.view.state.headerFooter).toBe('header');
    await editor.executeAsync('page.headerFooter', { part: 'footer', edit: true });
    expect(ui.view.state.headerFooter).toBe('footer');
    await editor.executeAsync('page.headerFooter', { part: 'footer', edit: false });
    expect(ui.view.state.headerFooter).toBeNull();
    expect(editor.canExecute('page.headerFooter', { part: 'body', edit: true })).toBe(false);
  });

  it('stampa con la funzione iniettata', async () => {
    const print = vi.fn();
    const editor = new Editor();
    editor.use(createUiModule({ print }));
    await editor.executeAsync('view.print');
    expect(print).toHaveBeenCalledOnce();
  });

  it('limita lo zoom', () => {
    expect(clampZoom(0.01)).toBe(0.1);
    expect(clampZoom(1.234)).toBe(1.23);
    expect(clampZoom(Number.POSITIVE_INFINITY)).toBe(1);
  });
});
