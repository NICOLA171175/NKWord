/**
 * Modulo commenti: registra i comandi `comment.*` (proprietario `comments`
 * in COMMAND_OWNERS) e tiene il nome dell'autore dei nuovi commenti.
 */

import {
  isCollapsed,
  newId,
  orderedRange,
  findParagraph,
  paragraphs,
  type Command,
  type CommandArgs,
  type CommandId,
  type Editor,
  type EditorModule,
  type EditorState,
  type Selection,
} from '../core';
import { mountCommentsPanel, type PanelHost } from './panel';
import { addComment, addReply, commentsAt, commentThreads, deleteComment, initialsOf, updateComment, wordAt } from './model';

type CommentCommand<K extends CommandId> = Command<CommandArgs[K]> & { readonly id: K };

export interface CommentsModuleOptions {
  /** Autore dei nuovi commenti (default "Utente"). */
  readonly author?: string;
  /** Data dei nuovi commenti; iniettabile per i test. */
  readonly now?: () => Date;
}

export interface CommentsModule extends EditorModule {
  /** Autore usato quando il comando non ne indica uno. */
  author: string;
  /** Monta il riquadro dei commenti accanto alla vista del documento (chiamato da src/ui/main.ts). */
  mount(host: PanelHost): () => void;
}

export const COMMENTS_MODULE_ID = 'comments';

/** Intervallo da commentare: la selezione, oppure la parola al cursore. */
export function commentTarget(state: EditorState): Selection | null {
  const sel = state.selection;
  if (!isCollapsed(sel)) {
    const [from, to] = orderedRange(state.doc, sel);
    return { anchor: from, focus: to };
  }
  const p = findParagraph(state.doc, sel.focus.paragraphId)?.paragraph;
  const word = p && wordAt(p, sel.focus.offset);
  if (!p || !word) return null;
  return { anchor: { paragraphId: p.id, offset: word[0] }, focus: { paragraphId: p.id, offset: word[1] } };
}

export function createCommentsModule(options: CommentsModuleOptions = {}): CommentsModule {
  const now = options.now ?? (() => new Date());
  const date = () => now().toISOString().replace(/\.\d{3}Z$/, 'Z');

  const module: CommentsModule = {
    id: COMMENTS_MODULE_ID,
    author: options.author ?? 'Utente',
    mount(host) {
      const panel = mountCommentsPanel(host, module);
      return () => panel.destroy();
    },
    activate(editor: Editor) {
      const disposers = commands().map((c) => editor.commands.register(c));
      return () => disposers.forEach((d) => d());
    },
  };

  const author = (a?: string) => (a?.trim() ? a.trim() : module.author);

  const commands = () => {
    const insert: CommentCommand<'comment.insert'> = {
      id: 'comment.insert',
      label: 'Nuovo commento',
      isEnabled: ({ state }) => commentTarget(state) !== null,
      run: ({ state }, args) => {
        const target = commentTarget(state);
        if (!target) return null;
        const name = author(args?.author);
        const id = newId();
        const doc = addComment(state.doc, target.anchor, target.focus, {
          id,
          author: name,
          initials: initialsOf(name),
          date: date(),
          text: args?.text ?? '',
        });
        return doc && { state: { ...state, doc, selection: target }, meta: { label: 'Nuovo commento' } };
      },
    };

    const reply: CommentCommand<'comment.reply'> = {
      id: 'comment.reply',
      label: 'Rispondi',
      isEnabled: ({ state }, args) => !!args && !!state.doc.comments?.[args.parentId],
      run: ({ state }, { parentId, text, author: a }) => {
        // Si risponde sempre al commento principale della discussione.
        const root = state.doc.comments?.[parentId]?.parentId ?? parentId;
        const name = author(a);
        const doc = addReply(state.doc, { id: newId(), author: name, initials: initialsOf(name), date: date(), text, parentId: root });
        return doc && { state: { ...state, doc }, meta: { label: 'Rispondi al commento' } };
      },
    };

    const edit: CommentCommand<'comment.edit'> = {
      id: 'comment.edit',
      label: 'Modifica commento',
      isEnabled: ({ state }, args) => !!args && !!state.doc.comments?.[args.id],
      run: ({ state }, { id, text }) => {
        if (state.doc.comments?.[id]?.text === text) return null;
        const doc = updateComment(state.doc, id, { text });
        return doc && { state: { ...state, doc }, meta: { label: 'Modifica commento' } };
      },
    };

    const target = (state: EditorState, id?: string) => id ?? commentsAt(state.doc, state.selection.focus)[0];

    const del: CommentCommand<'comment.delete'> = {
      id: 'comment.delete',
      label: 'Elimina commento',
      isEnabled: ({ state }, args) => {
        const id = target(state, args?.id);
        return !!id && !!state.doc.comments?.[id];
      },
      run: ({ state }, args) => {
        const id = target(state, args?.id);
        const doc = id ? deleteComment(state.doc, id) : null;
        return doc && { state: { ...state, doc }, meta: { label: 'Elimina commento' } };
      },
    };

    const resolve: CommentCommand<'comment.resolve'> = {
      id: 'comment.resolve',
      label: 'Risolvi commento',
      isEnabled: ({ state }, args) => !!args && !!state.doc.comments?.[args.id],
      run: ({ state }, { id, resolved }) => {
        if (!!state.doc.comments?.[id]?.resolved === resolved) return null;
        const doc = updateComment(state.doc, id, { resolved });
        return doc && { state: { ...state, doc }, meta: { label: resolved ? 'Risolvi commento' : 'Riapri commento' } };
      },
    };

    /** Sposta la selezione sul commento successivo o precedente rispetto al cursore. */
    const go = (state: EditorState, dir: 1 | -1) => {
      const threads = commentThreads(state.doc);
      if (!threads.length) return null;
      const order = new Map<string, number>();
      let i = 0;
      for (const { paragraph } of paragraphs(state.doc)) order.set(paragraph.id, i++);
      const key = (p: { paragraphId: string; offset: number }) => (order.get(p.paragraphId) ?? 0) * 1e9 + p.offset;
      const at = key(orderedRange(state.doc, state.selection)[0]);
      // Come in Word, alla fine si ricomincia dall'altro capo del documento.
      const pick =
        dir === 1
          ? threads.find((t) => key(t.range.from) > at) ?? threads[0]!
          : [...threads].reverse().find((t) => key(t.range.from) < at) ?? threads[threads.length - 1]!;
      const selection = { anchor: pick.range.from, focus: pick.range.to };
      return { state: { ...state, selection, pendingFormat: undefined }, meta: { addToHistory: false } };
    };

    const next: CommentCommand<'comment.next'> = {
      id: 'comment.next',
      label: 'Commento successivo',
      isEnabled: ({ state }) => commentThreads(state.doc).length > 0,
      run: ({ state }) => go(state, 1),
    };

    const previous: CommentCommand<'comment.previous'> = {
      id: 'comment.previous',
      label: 'Commento precedente',
      isEnabled: ({ state }) => commentThreads(state.doc).length > 0,
      run: ({ state }) => go(state, -1),
    };

    return [insert, reply, edit, del, resolve, next, previous];
  };

  return module;
}
