/**
 * Modulo revisioni (rilevamento modifiche, scheda Revisione di Word).
 *
 * - `revision.toggleTracking` accende/spegne il rilevamento: da acceso, un
 *   filtro di transazione riscrive ogni modifica del documento con le
 *   revisioni (vedi `track.ts`), qualunque modulo l'abbia fatta.
 * - `revision.accept/reject(All)` e `revision.next/previous` lavorano sulle
 *   revisioni presenti (vedi `review.ts`), anche quelle lette da un DOCX.
 * - Avvolge `editor.services.layout` per mostrare inserimenti ed
 *   eliminazioni colorati (vedi `markup.ts`), in qualunque ordine vengano
 *   attivati i moduli.
 */

import {
  type Command,
  type Editor,
  type EditorModule,
  type EditorState,
  type LayoutEngine,
  type Transaction,
} from '../core';
import { withRevisionMarkup } from './markup';
import { adjacentRevision, hasRevisions, resolveRevisions, scopeForSelection, type Scope } from './review';
import { trackChanges } from './track';

export const REVISIONS_MODULE_ID = 'revisions';

export interface RevisionsModuleOptions {
  /** Autore delle nuove revisioni (default 'Autore'). */
  readonly author?: string;
  /** Rilevamento acceso all'avvio (default no). */
  readonly tracking?: boolean;
  /** Orologio per la data delle revisioni (iniettabile nei test). */
  readonly now?: () => Date;
}

/** Stato del modulo, leggibile da UI e test. */
export interface RevisionsController {
  readonly tracking: boolean;
  readonly author: string;
}

/**
 * Comandi del modulo. Non fanno parte del catalogo di `src/core`
 * (`CommandArgs`): sono registrati per id e la barra multifunzione li trova
 * tramite `ribbonItems` (vedi `index.ts`).
 */
export interface RevisionCommandArgs {
  /** Attiva o disattiva il rilevamento modifiche. */
  'revision.toggleTracking': void;
  /** Nome con cui vengono firmate le nuove revisioni. */
  'revision.setAuthor': { author: string };
  /** Accetta le revisioni nella selezione (o quella al cursore). */
  'revision.accept': void;
  /** Rifiuta le revisioni nella selezione (o quella al cursore). */
  'revision.reject': void;
  'revision.acceptAll': void;
  'revision.rejectAll': void;
  /** Seleziona la revisione successiva / precedente. */
  'revision.next': void;
  'revision.previous': void;
}

export type RevisionCommandId = keyof RevisionCommandArgs;

type Cmd<K extends RevisionCommandId> = Command<RevisionCommandArgs[K]> & { readonly id: K };

export function createRevisionsModule(options: RevisionsModuleOptions = {}): EditorModule & { readonly controller: RevisionsController } {
  const settings = { tracking: options.tracking ?? false, author: options.author ?? 'Autore' };
  const now = options.now ?? (() => new Date());

  return {
    id: REVISIONS_MODULE_ID,
    controller: {
      get tracking() {
        return settings.tracking;
      },
      get author() {
        return settings.author;
      },
    },
    activate(editor: Editor) {
      // Transazioni dei comandi di revisione: non vanno rilevate.
      const own = new WeakSet<Transaction>();
      const ownTr = (state: EditorState, label: string, addToHistory = true): Transaction => {
        const tr: Transaction = { state, meta: { label, addToHistory } };
        own.add(tr);
        return tr;
      };
      /** Cambia solo le impostazioni: uno stato "nuovo" fa aggiornare la UI. */
      const touch = (state: EditorState, label: string) => ownTr({ ...state }, label, false);

      const resolve = (action: 'accept' | 'reject', all: boolean) => (state: EditorState) => {
        const scope: Scope | null = all ? { kind: 'all' } : scopeForSelection(state.doc, state.selection);
        if (!scope) return null;
        const r = resolveRevisions(state.doc, state.selection, action, scope);
        if (!r) return null;
        const label = action === 'accept' ? 'Accetta revisioni' : 'Rifiuta revisioni';
        return ownTr({ ...state, doc: r.doc, selection: r.selection }, label);
      };
      const go = (direction: 'next' | 'previous') => (state: EditorState) => {
        const span = adjacentRevision(state.doc, state.selection, direction);
        return span ? touch({ ...state, selection: { anchor: span.from, focus: span.to } }, 'revision.select') : null;
      };
      const any = ({ state }: { state: EditorState }) => hasRevisions(state.doc);

      const commands: Cmd<RevisionCommandId>[] = [
        {
          id: 'revision.toggleTracking',
          label: 'Rileva modifiche',
          isActive: () => settings.tracking,
          run: ({ state }) => {
            settings.tracking = !settings.tracking;
            return touch(state, 'revision.toggleTracking');
          },
        } satisfies Cmd<'revision.toggleTracking'>,
        {
          id: 'revision.setAuthor',
          label: 'Nome autore',
          isEnabled: (_ctx, args) => !!args?.author?.trim(),
          run: ({ state }, { author }) => {
            settings.author = author.trim();
            return touch(state, 'revision.setAuthor');
          },
        } satisfies Cmd<'revision.setAuthor'>,
        {
          id: 'revision.accept',
          label: 'Accetta',
          isEnabled: ({ state }) => !!scopeForSelection(state.doc, state.selection),
          run: ({ state }) => resolve('accept', false)(state),
        } satisfies Cmd<'revision.accept'>,
        {
          id: 'revision.reject',
          label: 'Rifiuta',
          isEnabled: ({ state }) => !!scopeForSelection(state.doc, state.selection),
          run: ({ state }) => resolve('reject', false)(state),
        } satisfies Cmd<'revision.reject'>,
        { id: 'revision.acceptAll', label: 'Accetta tutte le modifiche', isEnabled: any, run: ({ state }) => resolve('accept', true)(state) },
        { id: 'revision.rejectAll', label: 'Rifiuta tutte le modifiche', isEnabled: any, run: ({ state }) => resolve('reject', true)(state) },
        { id: 'revision.next', label: 'Successiva', isEnabled: any, run: ({ state }) => go('next')(state) },
        { id: 'revision.previous', label: 'Precedente', isEnabled: any, run: ({ state }) => go('previous')(state) },
      ] as Cmd<RevisionCommandId>[];
      const disposers = commands.map((c) => editor.commands.register(c));

      disposers.push(
        editor.addTransactionFilter((tr, state) => {
          const meta = tr.meta ?? {};
          if (!settings.tracking || own.has(tr) || tr.state.doc === state.doc) return tr;
          // Annulla/Ripeti, cursore, documento nuovo o aperto: niente da rilevare.
          if (meta.addToHistory === false || meta.label?.startsWith('file.')) return tr;
          const next = trackChanges(state, tr.state, {
            author: settings.author,
            date: now().toISOString(),
            backward: meta.mergeKey === 'deleteBackward',
          });
          return next === tr.state ? tr : { ...tr, state: next };
        }),
      );

      // Il motore di impaginazione può arrivare prima o dopo questo modulo:
      // la proprietà lo avvolge ogni volta che viene impostata.
      const services = editor.services;
      let inner: LayoutEngine | undefined = services.layout;
      let wrapped = inner && withRevisionMarkup(inner);
      Object.defineProperty(services, 'layout', {
        configurable: true,
        enumerable: true,
        get: () => wrapped,
        set: (engine: LayoutEngine | undefined) => {
          inner = engine;
          wrapped = engine && withRevisionMarkup(engine);
        },
      });
      disposers.push(() => {
        Object.defineProperty(services, 'layout', { configurable: true, enumerable: true, writable: true, value: inner });
        if (inner === undefined) delete services.layout;
      });

      return () => {
        for (const d of disposers.reverse()) d();
      };
    },
  };
}

export const revisionsModule = createRevisionsModule();
