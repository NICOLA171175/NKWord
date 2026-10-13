// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';
import { Editor, caret, createDocument, createState, paragraph, type GlyphRun } from '../../src/core';
import { createApproxMeasurer, createLayoutModule } from '../../src/layout';
import * as toc from '../../src/toc';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { createApp } from '../../src/ui/shell';
import { TabLeaders } from '../../src/ui/tab-leaders';

describe('sommario nella UI', () => {
  it('le voci dei moduli finiscono nella barra multifunzione', () => {
    expect(collectRibbonItems([{ ribbonItems: toc.ribbonItems }, { ribbonItems: [{ commandId: 'x' }] }, {}])).toHaveLength(3);
    expect(collectModules([toc]).map((m) => m.id)).toEqual(['toc']);

    const doc = createDocument([paragraph(''), paragraph('Titolo', {}, 'Heading1')]);
    const editor = new Editor({ ...createState(doc), selection: caret(doc.sections[0]!.blocks[0]!.id, 0) });
    editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
    editor.use(toc.tocModule);
    const root = document.createElement('div');
    const app = createApp(root, editor, { ribbonItems: collectRibbonItems([toc]) });
    app.ribbon.selectTab('references');
    const button = Array.from(root.querySelectorAll<HTMLButtonElement>('.ribbon-button')).find((b) => b.textContent?.includes('Sommario'));
    expect(button).toBeDefined();
    button!.click();
    expect(toc.findToc(editor.state.doc)).toBeDefined();

    // La voce del sommario ha i puntini fino al numero di pagina.
    expect(root.querySelector('.run-tab-leader')?.textContent).toMatch(/^\.+$/);
    app.destroy();
  });

  it('sceglie il riempitivo della tabulazione successiva', () => {
    const p = paragraph('a', { tabs: [{ position: 50, align: 'left' }, { position: 200, align: 'right', leader: 'hyphen' }] });
    const leaders = new TabLeaders(createDocument([p]));
    const run = (kind: GlyphRun['kind']) => ({ kind, paragraphId: p.id, start: 0, end: 1, text: '\t', format: {}, x: 0, width: 10 });
    expect(leaders.leader(run('tab'), 10)).toBeUndefined();
    expect(leaders.leader(run('tab'), 60)).toBe('-');
    expect(leaders.leader(run('text'), 60)).toBeUndefined();
  });
});
