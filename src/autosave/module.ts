/**
 * Salvataggio automatico e recupero, come "Salvataggio automatico delle
 * informazioni di ripristino" di Word.
 *
 * - Dopo ogni modifica, trascorsa una breve pausa (o al massimo ogni
 *   `maxWaitMs` mentre si continua a scrivere), il documento viene copiato
 *   nell'archivio di recupero (IndexedDB nel browser e nel programma desktop).
 * - Quando il documento viene salvato (`file.export` in un formato
 *   riapribile), creato o aperto, la copia si cancella: non serve più.
 * - All'avvio, le copie lasciate da finestre chiuse senza salvare (o da una
 *   chiusura improvvisa) sono proposte per il ripristino.
 */

import { createState, type AsyncCommand, type Command, type DocumentModel, type Editor, type EditorModule } from '../core';
import { createPresence, type Presence } from './presence';
import { defaultRecoveryStore, type RecoverySnapshot, type RecoveryStore } from './store';

export const AUTOSAVE_MODULE_ID = 'autosave';

/** Transazioni dopo le quali il documento non ha modifiche da salvare. */
const CLEAN_LABELS: readonly string[] = ['file.new', 'file.open'];
export const RESTORE_LABEL = 'autosave.restore';

export interface AutosaveCommandArgs {
  /** Ripristina una copia (per default la più recente). */
  'autosave.restore': { id?: string } | void;
  /** Mostra di nuovo l'avviso con i documenti da recuperare. */
  'autosave.show': void;
  /** Elimina una copia (per default tutte quelle da recuperare). */
  'autosave.discard': { id?: string } | void;
}

export interface AutosaveOptions {
  readonly store?: RecoveryStore;
  readonly presence?: Presence;
  /** Pausa dopo l'ultima modifica prima di salvare la copia (default 2 s). */
  readonly delayMs?: number;
  /** Scrivendo senza pause, la copia si aggiorna comunque ogni tanto (default 10 s). */
  readonly maxWaitMs?: number;
  readonly now?: () => Date;
  /** Id di questa finestra; per default uno casuale. */
  readonly sessionId?: string;
}

export interface AutosaveModule extends EditorModule {
  /** Copie da recuperare, dalla più recente. */
  readonly recovered: readonly RecoverySnapshot[];
  /** Il documento ha modifiche non salvate. */
  readonly dirty: boolean;
  /** Data dell'ultima copia di recupero scritta. */
  readonly lastSavedAt: Date | null;
  /** Si risolve quando le copie da recuperare sono state lette. */
  readonly ready: Promise<void>;
  /** Scrive subito la copia di recupero, se ci sono modifiche. */
  flush(): Promise<void>;
  /** Avvisa quando cambiano `recovered`, `dirty` o `lastSavedAt`. */
  onChange(listener: () => void): () => void;
}

export function parseSnapshot(snapshot: RecoverySnapshot): DocumentModel {
  return JSON.parse(snapshot.doc) as DocumentModel;
}

function randomId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

export function createAutosaveModule(options: AutosaveOptions = {}): AutosaveModule {
  const delayMs = options.delayMs ?? 2000;
  const maxWaitMs = options.maxWaitMs ?? 10000;
  const now = options.now ?? (() => new Date());
  const selfId = options.sessionId ?? randomId();
  let store: RecoveryStore | undefined = options.store;
  let presence: Presence | undefined = options.presence;

  let editor: Editor | undefined;
  let recovered: RecoverySnapshot[] = [];
  let dirty = false;
  let lastSavedAt: Date | null = null;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let pendingSince: number | undefined;
  let queue: Promise<void> = Promise.resolve();
  let resolveReady!: () => void;
  const ready = new Promise<void>((r) => (resolveReady = r));
  const listeners = new Set<() => void>();
  const notify = () => listeners.forEach((l) => l());

  /** Le operazioni sull'archivio vanno in fila, così una cancellazione non supera una scrittura. */
  const enqueue = (op: (s: RecoveryStore) => Promise<void>) => {
    const s = store!;
    queue = queue.then(() => op(s)).catch((error) => console.error('Salvataggio automatico non riuscito', error));
    return queue;
  };

  const cancelTimer = () => {
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
    pendingSince = undefined;
  };

  const writeSnapshot = (): Promise<void> => {
    cancelTimer();
    if (!editor || !dirty) return queue;
    const doc = editor.state.doc;
    const savedAt = now();
    const snapshot: RecoverySnapshot = { id: selfId, title: doc.metadata.title ?? '', savedAt: savedAt.toISOString(), doc: JSON.stringify(doc) };
    return enqueue(async (s) => {
      await s.put(snapshot);
      lastSavedAt = savedAt;
      notify();
    });
  };

  const schedule = () => {
    const t = Date.now();
    pendingSince ??= t;
    if (timer !== undefined) clearTimeout(timer);
    if (t - pendingSince >= maxWaitMs) void writeSnapshot();
    else timer = setTimeout(() => void writeSnapshot(), delayMs);
  };

  const markClean = () => {
    cancelTimer();
    const wasDirty = dirty;
    dirty = false;
    void enqueue((s) => s.remove(selfId));
    if (wasDirty) notify();
  };

  // Documenti già letti dalle copie: riconoscono quale copia è stata ripristinata.
  const parsed = new Map<string, DocumentModel>();
  const docOf = (snapshot: RecoverySnapshot) => {
    let doc = parsed.get(snapshot.id);
    if (!doc) parsed.set(snapshot.id, (doc = parseSnapshot(snapshot)));
    return doc;
  };

  const forget = (id: string) => {
    parsed.delete(id);
    recovered = recovered.filter((r) => r.id !== id);
    void enqueue((s) => s.remove(id));
    notify();
  };

  const restoreCommand: Command<AutosaveCommandArgs['autosave.restore']> = {
    id: 'autosave.restore',
    label: 'Recupera documento non salvato',
    isEnabled: (_ctx, args) => (args?.id ? recovered.some((r) => r.id === args.id) : recovered.length > 0),
    run: (_ctx, args) => {
      const snapshot = args?.id ? recovered.find((r) => r.id === args.id) : recovered[0];
      if (!snapshot) return null;
      return { state: createState(docOf(snapshot)), meta: { addToHistory: false, label: RESTORE_LABEL } };
    },
  };

  const discardCommand: AsyncCommand<AutosaveCommandArgs['autosave.discard']> = {
    id: 'autosave.discard',
    label: 'Elimina documenti recuperati',
    async: true,
    isEnabled: (_ctx, args) => (args?.id ? recovered.some((r) => r.id === args.id) : recovered.length > 0),
    async run(_ctx, args) {
      for (const r of args?.id ? recovered.filter((x) => x.id === args.id) : [...recovered]) forget(r.id);
      await queue;
      return null;
    },
  };

  const showCommand: Command<void> = {
    id: 'autosave.show',
    label: 'Documenti recuperati',
    isEnabled: () => recovered.length > 0,
    run: () => null,
  };

  const module: AutosaveModule = {
    id: AUTOSAVE_MODULE_ID,
    get recovered() {
      return recovered;
    },
    get dirty() {
      return dirty;
    },
    get lastSavedAt() {
      return lastSavedAt;
    },
    ready,
    flush: () => writeSnapshot(),
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(ed) {
      editor = ed;
      store ??= defaultRecoveryStore();
      presence ??= createPresence(selfId);
      const disposers = [ed.commands.register(restoreCommand), ed.commands.register(discardCommand), ed.commands.register(showCommand)];

      disposers.push(
        ed.events.on('change', ({ state, previous, meta }) => {
          if (state.doc === previous.doc) return;
          if (meta.label && CLEAN_LABELS.includes(meta.label)) return markClean();
          if (meta.label === RESTORE_LABEL) {
            // La copia ripristinata diventa il documento di questa finestra.
            for (const [id, doc] of parsed) if (doc === state.doc) forget(id);
          }
          const wasDirty = dirty;
          dirty = true;
          if (!wasDirty) notify();
          schedule();
        }),
        // Salvato in un formato che si può riaprire: la copia di recupero non serve più.
        ed.events.on('executed', ({ commandId, args }) => {
          if (commandId !== 'file.export') return;
          const formatId = (args as { formatId?: string } | undefined)?.formatId;
          if (formatId && ed.services.formats?.get(formatId)?.import) markClean();
        }),
      );

      const p = presence;
      void (async () => {
        try {
          const [all, alive] = await Promise.all([store!.list(), p.others()]);
          recovered = all.filter((s) => s.id !== selfId && !alive.has(s.id)).sort((a, b) => b.savedAt.localeCompare(a.savedAt));
          notify();
        } catch (error) {
          console.error('Lettura dei documenti da recuperare non riuscita', error);
        } finally {
          resolveReady();
        }
      })();

      return () => {
        cancelTimer();
        for (const d of disposers) d();
        p.dispose();
        editor = undefined;
      };
    },
  };
  return module;
}
