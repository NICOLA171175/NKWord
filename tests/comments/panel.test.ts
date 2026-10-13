// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, createState, Editor, paragraph, textRun } from '../../src/core';
import { commentThreads, createCommentsModule } from '../../src/comments';
import { editingModule } from '../../src/editing';
import { createApp } from '../../src/ui/shell';

afterEach(() => document.body.replaceChildren());

function setup() {
  const p = paragraph([textRun('Ciao bel mondo')]);
  const editor = new Editor({ ...createState(createDocument([p])), selection: { anchor: { paragraphId: p.id, offset: 5 }, focus: { paragraphId: p.id, offset: 8 } } });
  editor.use(editingModule);
  const comments = createCommentsModule({ author: 'Nicola' });
  editor.use(comments);
  const root = document.createElement('div');
  document.body.append(root);
  const app = createApp(root, editor);
  comments.mount(app);
  const pane = root.querySelector<HTMLElement>('.comments-pane')!;
  const click = (label: string, scope: ParentNode = pane) =>
    Array.from(scope.querySelectorAll<HTMLButtonElement>('button')).find((b) => b.textContent === label)!.click();
  return { editor, app, root, pane, p, click };
}

const tick = () => new Promise((r) => setTimeout(r, 0));

describe('riquadro dei commenti', () => {
  it('compare accanto al documento solo quando ci sono commenti', () => {
    const { root, pane } = setup();
    expect(root.querySelector('.comments-workspace > .doc-scroll + .comments-pane')).not.toBeNull();
    expect(pane.hidden).toBe(true);
  });

  it('il pulsante della barra crea un commento e apre subito la casella di testo', async () => {
    const { editor, app, root, pane, click } = setup();
    app.ribbon.selectTab('review');
    const button = Array.from(root.querySelectorAll<HTMLButtonElement>('.ribbon button')).find((b) => b.textContent?.includes('Nuovo commento'))!;
    expect(button.disabled).toBe(false);
    button.click();
    await tick();
    expect(pane.hidden).toBe(false);
    const box = pane.querySelector<HTMLTextAreaElement>('textarea')!;
    expect(document.activeElement).toBe(box);
    expect(pane.querySelector('.comment-quote')!.textContent).toBe('bel');
    box.value = 'Parola da cambiare';
    box.dispatchEvent(new Event('input'));
    click('Salva');
    expect(commentThreads(editor.state.doc)[0]!.comment.text).toBe('Parola da cambiare');
    expect(pane.querySelector('.comment-text')!.textContent).toBe('Parola da cambiare');
    expect(pane.querySelector('.comment-author')!.textContent).toBe('Nicola');
    // Il testo commentato è evidenziato nella pagina.
    expect(root.querySelector('.run-comment')?.textContent).toBe('bel');
  });

  it('risponde, risolve ed elimina dal riquadro', () => {
    const { editor, pane, click } = setup();
    editor.execute('comment.insert', { text: 'Primo' });
    click('Rispondi');
    const box = pane.querySelector<HTMLTextAreaElement>('textarea')!;
    box.value = 'Risposta';
    box.dispatchEvent(new Event('input'));
    box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true }));
    expect(commentThreads(editor.state.doc)[0]!.replies.map((r) => r.text)).toEqual(['Risposta']);
    expect(pane.querySelectorAll('.comment-reply')).toHaveLength(1);

    click('Risolvi');
    expect(pane.querySelector('.comment-card')!.classList.contains('resolved')).toBe(true);
    click('Riapri');
    expect(pane.querySelector('.comment-card')!.classList.contains('resolved')).toBe(false);

    click('Elimina', pane.querySelector('.comment-reply')!);
    expect(pane.querySelectorAll('.comment-reply')).toHaveLength(0);
    click('Elimina discussione');
    expect(editor.state.doc.comments).toBeUndefined();
    expect(pane.hidden).toBe(true);
    editor.undo();
    expect(pane.hidden).toBe(false);
  });

  it('un clic sulla scheda seleziona il testo commentato e la scheda è attiva', () => {
    const { editor, pane, p } = setup();
    editor.execute('comment.insert', { text: 'x' });
    editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: p.id, offset: 0 }, focus: { paragraphId: p.id, offset: 0 } } });
    expect(pane.querySelector('.comment-card.active')).toBeNull();
    pane.querySelector<HTMLElement>('.comment-quote')!.click();
    expect(editor.state.selection).toEqual({ anchor: { paragraphId: p.id, offset: 5 }, focus: { paragraphId: p.id, offset: 8 } });
    expect(pane.querySelector('.comment-card.active')).not.toBeNull();
  });
});
