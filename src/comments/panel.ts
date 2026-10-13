/**
 * Riquadro laterale dei commenti, come il riquadro "Commenti" di Word:
 * una scheda per discussione nell'ordine del documento, con risposte,
 * modifica, risoluzione ed eliminazione. Tutto passa dai comandi
 * `comment.*`, quindi ogni azione si può annullare.
 */

import { paragraphs, paragraphText, type Comment, type Editor, type Position } from '../core';
import { commentsAt, commentThreads, type CommentThread } from './model';
import type { CommentsModule } from './module';
import './styles.css';

type Draft = { readonly kind: 'reply'; readonly id: string; text: string } | { readonly kind: 'edit'; readonly id: string; text: string };

const QUOTE_MAX = 80;

/** Testo commentato, abbreviato per l'intestazione della scheda. */
export function quotedText(editor: Editor, from: Position, to: Position): string {
  let out = '';
  let inside = false;
  for (const { paragraph: p } of paragraphs(editor.state.doc)) {
    const isFrom = p.id === from.paragraphId;
    const isTo = p.id === to.paragraphId;
    if (isFrom) inside = true;
    if (inside) {
      const text = paragraphText(p);
      out += (out ? ' ' : '') + text.slice(isFrom ? from.offset : 0, isTo ? to.offset : text.length);
    }
    if (isTo || out.length > QUOTE_MAX) break;
  }
  out = out.replace(/[\t\n￼]/g, ' ').replace(/\s+/g, ' ').trim();
  return out.length > QUOTE_MAX ? `${out.slice(0, QUOTE_MAX - 1)}…` : out;
}

export function formatCommentDate(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  return d.toLocaleString('it-IT', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}

/** Ciò che serve dell'app per montare il riquadro (vedi `createApp` in src/ui). */
export interface PanelHost {
  readonly editor: Editor;
  readonly view: { readonly element: HTMLElement };
}

/**
 * Affianca il riquadro alla vista del documento: la vista e il riquadro
 * finiscono in una riga `.comments-workspace` al posto della vista.
 */
export function mountCommentsPanel(host: PanelHost, module: CommentsModule): CommentsPanel {
  const view = host.view.element;
  const panel = new CommentsPanel(host.editor, module);
  const row = document.createElement('div');
  row.className = 'comments-workspace';
  view.replaceWith(row);
  row.append(view, panel.element);
  return panel;
}

export class CommentsPanel {
  readonly element = document.createElement('aside');
  private readonly list = document.createElement('div');
  private readonly count = document.createElement('span');
  private readonly activeStyle = document.createElement('style');
  private draft: Draft | null = null;
  private knownIds = new Set<string>();
  private readonly dispose: () => void;

  constructor(
    private readonly editor: Editor,
    private readonly module: CommentsModule,
  ) {
    this.element.className = 'comments-pane';
    this.element.setAttribute('aria-label', 'Commenti');
    const header = document.createElement('header');
    header.className = 'comments-header';
    const title = document.createElement('span');
    title.textContent = 'Commenti';
    this.count.className = 'comments-count';
    header.append(title, this.count);
    this.list.className = 'comments-list';
    this.element.append(header, this.list, this.activeStyle);
    this.knownIds = new Set(Object.keys(editor.state.doc.comments ?? {}));
    this.dispose = editor.events.on('change', () => this.render());
    this.render();
  }

  render(): void {
    const doc = this.editor.state.doc;
    const threads = commentThreads(doc);
    const ids = Object.keys(doc.comments ?? {});
    // Un commento appena creato vuoto si apre subito in modifica, come in Word.
    const created = ids.filter((id) => !this.knownIds.has(id));
    this.knownIds = new Set(ids);
    const fresh = created.map((id) => doc.comments![id]!).find((c) => !c.parentId && c.text === '');
    if (fresh) this.draft = { kind: 'edit', id: fresh.id, text: '' };
    if (this.draft && !doc.comments?.[this.draft.id]) this.draft = null;

    const active = new Set(commentsAt(doc, this.editor.state.selection.focus));
    this.element.hidden = threads.length === 0;
    this.count.textContent = threads.length ? String(threads.length) : '';
    this.activeStyle.textContent = [...active]
      .map((id) => `.run-comment[data-comments~="${id}"]{background-color:var(--comment-active,#f5c518)}`)
      .join('\n');
    this.list.replaceChildren(...threads.map((t) => this.renderThread(t, active.has(t.comment.id))));

    if (fresh) setTimeout(() => this.focusDraft(), 0);
  }

  /** Mette il cursore nella casella di testo aperta (risposta o modifica). */
  focusDraft(): void {
    const box = this.list.querySelector<HTMLTextAreaElement>('textarea');
    box?.focus();
  }

  private renderThread(t: CommentThread, active: boolean): HTMLElement {
    const card = document.createElement('article');
    card.className = 'comment-card';
    card.dataset.id = t.comment.id;
    if (active) card.classList.add('active');
    if (t.comment.resolved) card.classList.add('resolved');
    card.addEventListener('click', (e) => {
      if ((e.target as HTMLElement).closest('button, textarea')) return;
      const selection = { anchor: t.range.from, focus: t.range.to };
      if (this.editor.canExecute('edit.setSelection', { selection })) this.editor.execute('edit.setSelection', { selection });
    });

    const quote = document.createElement('div');
    quote.className = 'comment-quote';
    quote.textContent = quotedText(this.editor, t.range.from, t.range.to);
    card.append(quote, this.renderEntry(t.comment, true));
    for (const r of t.replies) card.append(this.renderEntry(r, false));

    if (this.draft?.kind === 'reply' && this.draft.id === t.comment.id) {
      card.append(this.renderForm(this.draft, 'Rispondi'));
    } else {
      const actions = document.createElement('div');
      actions.className = 'comment-actions';
      actions.append(
        this.button('Rispondi', () => this.open({ kind: 'reply', id: t.comment.id, text: '' })),
        this.button(t.comment.resolved ? 'Riapri' : 'Risolvi', () =>
          this.editor.execute('comment.resolve', { id: t.comment.id, resolved: !t.comment.resolved }),
        ),
        this.button('Elimina discussione', () => this.editor.execute('comment.delete', { id: t.comment.id })),
      );
      card.append(actions);
    }
    return card;
  }

  private renderEntry(c: Comment, isRoot: boolean): HTMLElement {
    const entry = document.createElement('div');
    entry.className = isRoot ? 'comment-entry' : 'comment-entry comment-reply';
    entry.dataset.id = c.id;
    const head = document.createElement('div');
    head.className = 'comment-head';
    const avatar = document.createElement('span');
    avatar.className = 'comment-avatar';
    avatar.textContent = c.initials || c.author.slice(0, 1).toUpperCase();
    avatar.style.background = colorFor(c.author);
    const who = document.createElement('span');
    who.className = 'comment-author';
    who.textContent = c.author;
    const when = document.createElement('time');
    when.className = 'comment-date';
    when.dateTime = c.date;
    when.textContent = formatCommentDate(c.date);
    head.append(avatar, who, when);
    entry.append(head);

    if (this.draft?.kind === 'edit' && this.draft.id === c.id) {
      entry.append(this.renderForm(this.draft, 'Salva'));
      return entry;
    }
    const text = document.createElement('div');
    text.className = 'comment-text';
    text.textContent = c.text;
    entry.append(text);
    const tools = document.createElement('div');
    tools.className = 'comment-tools';
    tools.append(this.button('Modifica', () => this.open({ kind: 'edit', id: c.id, text: c.text })));
    if (!isRoot) tools.append(this.button('Elimina', () => this.editor.execute('comment.delete', { id: c.id })));
    entry.append(tools);
    return entry;
  }

  private renderForm(draft: Draft, submitLabel: string): HTMLElement {
    const form = document.createElement('div');
    form.className = 'comment-form';
    const box = document.createElement('textarea');
    box.className = 'comment-input';
    box.rows = 2;
    box.placeholder = draft.kind === 'reply' ? 'Rispondi…' : 'Scrivi un commento…';
    box.value = draft.text;
    box.addEventListener('input', () => {
      draft.text = box.value;
    });
    const submit = () => this.submit(draft);
    box.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        submit();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.close();
      }
    });
    const buttons = document.createElement('div');
    buttons.className = 'comment-actions';
    buttons.append(this.button(submitLabel, submit, true), this.button('Annulla', () => this.close()));
    form.append(box, buttons);
    return form;
  }

  private submit(draft: Draft): void {
    const text = draft.text.trim();
    this.draft = null;
    if (draft.kind === 'reply') {
      if (text) this.editor.execute('comment.reply', { parentId: draft.id, text, author: this.module.author });
    } else this.editor.execute('comment.edit', { id: draft.id, text });
    this.render();
  }

  private open(draft: Draft): void {
    this.draft = draft;
    this.render();
    this.focusDraft();
  }

  private close(): void {
    this.draft = null;
    this.render();
  }

  private button(label: string, onClick: () => void, primary = false): HTMLButtonElement {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = primary ? 'comment-button primary' : 'comment-button';
    b.textContent = label;
    b.addEventListener('click', (e) => {
      e.stopPropagation();
      onClick();
    });
    return b;
  }

  destroy(): void {
    this.dispose();
    this.element.remove();
  }
}

const AVATAR_COLORS = ['#2b579a', '#a4373a', '#107c10', '#8764b8', '#c239b3', '#ca5010', '#038387'];

function colorFor(name: string): string {
  let h = 0;
  for (const ch of name) h = (h * 31 + ch.codePointAt(0)!) >>> 0;
  return AVATAR_COLORS[h % AVATAR_COLORS.length]!;
}
