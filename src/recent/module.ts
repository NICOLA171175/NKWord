/**
 * Documenti recenti, come in Word (File > Apri > Recenti): l'elenco degli
 * ultimi file aperti o salvati, riapribili con un clic.
 *
 * - Programma desktop: si ricorda il percorso del file sul disco e lo si
 *   rilegge direttamente.
 * - Browser con File System Access (Chrome, Edge): si ricorda la maniglia del
 *   file; riaprendolo il browser può chiedere il permesso di leggerlo.
 * - Altri browser, o file solo scaricati: si ricorda il nome e, cliccandolo,
 *   si apre la finestra per sceglierlo di nuovo.
 */

import type { AsyncCommand, Command, Editor, EditorModule } from '../core';
import { suggestedFileName } from '../io';
import type { FileHandle, FileOrigin, OpenedFile } from '../ui/shell';
import { baseName, desktopBridge, type DesktopBridge } from './desktop';
import { defaultRecentStore, type RecentEntry, type RecentStore } from './store';

export const RECENT_MODULE_ID = 'recent';
/** Quanti documenti restano nell'elenco. */
export const MAX_RECENT = 20;

export interface RecentCommandArgs {
  /** Mostra l'elenco dei documenti recenti. */
  'recent.show': void;
  /** Riapre un documento dell'elenco. */
  'recent.open': { id: string };
  /** Toglie un documento dall'elenco (il file resta dov'è). */
  'recent.remove': { id: string };
  /** Svuota l'elenco. */
  'recent.clear': void;
}

/** Chi apre davvero i file: la shell dell'app (`App`). */
export interface RecentOpener {
  openFile(file: File, from?: FileOrigin): Promise<boolean>;
  pickFile(): void;
}

/** Esito di una riapertura: aperto, finestra di scelta mostrata, file sparito o permesso negato. */
export type RecentOpenResult = 'opened' | 'picked' | 'missing' | 'denied' | 'failed';

export interface RecentOptions {
  readonly store?: RecentStore;
  /** Ponte con il programma desktop; per default `window.nkofficeDesktop`. */
  readonly desktop?: DesktopBridge | null;
  readonly now?: () => Date;
}

export interface RecentModule extends EditorModule {
  /** Documenti dal più recente. */
  readonly entries: readonly RecentEntry[];
  /** Si risolve quando l'elenco salvato è stato letto. */
  readonly ready: Promise<void>;
  /** Aggiunge (o porta in cima) un file appena aperto. */
  record(opened: OpenedFile): Promise<void>;
  open(id: string): Promise<RecentOpenResult>;
  remove(id: string): Promise<void>;
  clear(): Promise<void>;
  /** Collega chi apre i file (la shell). */
  attach(opener: RecentOpener): () => void;
  onChange(listener: () => void): () => void;
}

function randomId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

/** Su Windows i percorsi non distinguono maiuscole e minuscole. */
const samePath = (a: string, b: string) => (/^[a-z]:\\/i.test(a) ? a.toLowerCase() === b.toLowerCase() : a === b);

const extensionOf = (name: string) => (name.includes('.') ? name.split('.').pop()!.toLowerCase() : '');

export function createRecentModule(options: RecentOptions = {}): RecentModule {
  const now = options.now ?? (() => new Date());
  let store = options.store;
  const desktop = options.desktop === undefined ? desktopBridge() : (options.desktop ?? undefined);
  let entries: RecentEntry[] = [];
  let opener: RecentOpener | undefined;
  let editor: Editor | undefined;
  let queue: Promise<unknown> = Promise.resolve();
  let resolveReady!: () => void;
  const ready = new Promise<void>((r) => (resolveReady = r));
  const listeners = new Set<() => void>();
  const notify = () => listeners.forEach((l) => l());

  /** Le operazioni vanno in fila: due aperture ravvicinate non si sovrappongono. */
  const serial = <T>(op: () => Promise<T>): Promise<T> => {
    const next = queue.then(op);
    queue = next.catch((error) => console.error('Documenti recenti: operazione non riuscita', error));
    return next;
  };

  const save = async (list: RecentEntry[]) => {
    const dropped = entries.filter((e) => !list.some((x) => x.id === e.id));
    entries = list;
    notify();
    for (const e of dropped) await store!.remove(e.id);
  };

  const find = async (origin: { name: string; path?: string; handle?: FileHandle; savedOnly?: boolean }) => {
    if (origin.path) return entries.find((e) => e.path && samePath(e.path, origin.path!));
    if (origin.handle) {
      for (const e of entries) {
        if (e.handle && (await e.handle.isSameEntry?.(origin.handle).catch(() => false))) return e;
      }
      return undefined;
    }
    // Solo il nome: un file salvato con lo stesso nome di uno dell'elenco è lui.
    return entries.find((e) => e.name === origin.name && (origin.savedOnly || (!e.path && !e.handle)));
  };

  const add = (origin: { name: string; path?: string; handle?: FileHandle; savedOnly?: boolean }) =>
    serial(async () => {
      if (!store) return;
      const old = await find(origin);
      const entry: RecentEntry = {
        ...old,
        id: old?.id ?? randomId(),
        name: origin.path || origin.handle ? origin.name : (old?.name ?? origin.name),
        usedAt: now().toISOString(),
        ...(origin.path ? { path: origin.path } : {}),
        ...(origin.handle ? { handle: origin.handle } : {}),
      };
      await store.put(entry);
      await save([entry, ...entries.filter((e) => e.id !== entry.id)].slice(0, MAX_RECENT));
    });

  const remove = (id: string) => serial(() => save(entries.filter((e) => e.id !== id)));
  const clear = () => serial(() => save([]));

  const readEntry = async (entry: RecentEntry): Promise<File | Exclude<RecentOpenResult, 'opened'>> => {
    if (entry.path && desktop) {
      try {
        const bytes = await desktop.readFile(entry.path);
        return new File([bytes as Uint8Array<ArrayBuffer>], baseName(entry.path));
      } catch {
        return 'missing';
      }
    }
    if (entry.handle) {
      const h = entry.handle;
      try {
        if (h.queryPermission && (await h.queryPermission({ mode: 'read' })) !== 'granted') {
          if ((await h.requestPermission?.({ mode: 'read' })) !== 'granted') return 'denied';
        }
        return await h.getFile();
      } catch (error) {
        return (error as { name?: string })?.name === 'NotAllowedError' ? 'denied' : 'missing';
      }
    }
    opener!.pickFile();
    return 'picked';
  };

  const open = async (id: string): Promise<RecentOpenResult> => {
    const entry = entries.find((e) => e.id === id);
    if (!entry || !opener) return 'failed';
    const file = await readEntry(entry);
    if (typeof file === 'string') return file;
    const ok = await opener.openFile(file, { ...(entry.path ? { path: entry.path } : {}), ...(entry.handle ? { handle: entry.handle } : {}) });
    return ok ? 'opened' : 'failed';
  };

  const has = (args: { id: string } | undefined) => !!args && entries.some((e) => e.id === args.id);

  const commands: (Command<never> | AsyncCommand<never>)[] = [
    { id: 'recent.show', label: 'Documenti recenti', run: () => null } satisfies Command<void>,
    {
      id: 'recent.open',
      label: 'Apri documento recente',
      async: true,
      isEnabled: (_ctx, args) => !!opener && (!args || has(args)),
      async run(_ctx, args) {
        await open(args.id);
        return null;
      },
    } satisfies AsyncCommand<RecentCommandArgs['recent.open']>,
    {
      id: 'recent.remove',
      label: "Rimuovi dall'elenco",
      async: true,
      isEnabled: (_ctx, args) => !args || has(args),
      async run(_ctx, args) {
        await remove(args.id);
        return null;
      },
    } satisfies AsyncCommand<RecentCommandArgs['recent.remove']>,
    {
      id: 'recent.clear',
      label: 'Cancella documenti recenti',
      async: true,
      isEnabled: () => entries.length > 0,
      async run() {
        await clear();
        return null;
      },
    } satisfies AsyncCommand<void>,
  ] as never;

  const module: RecentModule = {
    id: RECENT_MODULE_ID,
    get entries() {
      return entries;
    },
    ready,
    record({ file, handle, path }) {
      const p = path || desktop?.pathForFile(file) || undefined;
      return add({ name: p ? baseName(p) : (handle?.name ?? file.name), ...(p ? { path: p } : {}), ...(handle ? { handle } : {}) });
    },
    open,
    remove,
    clear,
    attach(o) {
      opener = o;
      return () => {
        if (opener === o) opener = undefined;
      };
    },
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(ed) {
      editor = ed;
      store ??= defaultRecentStore();
      const disposers = commands.map((c) => ed.commands.register(c));

      // Un file riapribile appena salvato entra nell'elenco.
      const reopenable = (ext: string) => !!editor?.services.formats?.byExtension(ext)?.import;
      if (desktop) {
        // Il programma desktop conosce il percorso scelto nella finestra "Salva".
        disposers.push(desktop.onFileSaved((path) => void (reopenable(extensionOf(path)) && add({ name: baseName(path), path }))));
      } else {
        disposers.push(
          ed.events.on('executed', ({ commandId, args }) => {
            if (commandId !== 'file.export') return;
            const format = ed.services.formats?.get((args as { formatId?: string } | undefined)?.formatId ?? '');
            if (!format?.import) return;
            void add({ name: suggestedFileName(ed.state.doc, format.extensions[0] ?? format.id), savedOnly: true });
          }),
        );
      }

      const s = store;
      void serial(async () => {
        try {
          const saved = (await s.list()).sort((a, b) => b.usedAt.localeCompare(a.usedAt));
          // Quelli aggiunti nel frattempo restano in cima.
          entries = [...entries, ...saved.filter((x) => !entries.some((e) => e.id === x.id))].slice(0, MAX_RECENT);
          notify();
        } finally {
          resolveReady();
        }
      });

      return () => {
        for (const d of disposers) d();
        editor = undefined;
      };
    },
  };
  return module;
}
