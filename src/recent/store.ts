/**
 * Archivio dei documenti recenti. Si usa IndexedDB perché, nel browser,
 * conserva anche la "maniglia" del file (File System Access API), che
 * permette di riaprirlo senza cercarlo di nuovo; se manca, la memoria.
 */

import type { FileHandle } from '../ui/shell';

export interface RecentEntry {
  readonly id: string;
  /** Nome del file, es. 'Relazione.docx'. */
  readonly name: string;
  /** Data ISO 8601 dell'ultima apertura o dell'ultimo salvataggio. */
  readonly usedAt: string;
  /** Percorso sul disco (programma desktop). */
  readonly path?: string;
  /** Maniglia del file (browser che la supportano). */
  readonly handle?: FileHandle;
}

export interface RecentStore {
  list(): Promise<RecentEntry[]>;
  put(entry: RecentEntry): Promise<void>;
  remove(id: string): Promise<void>;
}

export class MemoryRecentStore implements RecentStore {
  private readonly items = new Map<string, RecentEntry>();
  async list() {
    return [...this.items.values()];
  }
  async put(entry: RecentEntry) {
    this.items.set(entry.id, entry);
  }
  async remove(id: string) {
    this.items.delete(id);
  }
}

const DB_NAME = 'nkoffice-recenti';
const STORE = 'documenti';

export class IndexedDbRecentStore implements RecentStore {
  private db: Promise<IDBDatabase> | undefined;

  constructor(private readonly factory: IDBFactory) {}

  private open(): Promise<IDBDatabase> {
    this.db ??= new Promise((resolve, reject) => {
      const req = this.factory.open(DB_NAME, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return this.db;
  }

  private async run<T>(mode: IDBTransactionMode, op: (store: IDBObjectStore) => IDBRequest): Promise<T> {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const req = op(tx.objectStore(STORE));
      tx.oncomplete = () => resolve(req.result as T);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }

  list() {
    return this.run<RecentEntry[]>('readonly', (s) => s.getAll());
  }
  async put(entry: RecentEntry) {
    await this.run('readwrite', (s) => s.put(entry));
  }
  async remove(id: string) {
    await this.run('readwrite', (s) => s.delete(id));
  }
}

/** L'archivio migliore disponibile in questo ambiente. */
export function defaultRecentStore(): RecentStore {
  try {
    if (typeof indexedDB !== 'undefined') return new IndexedDbRecentStore(indexedDB);
  } catch {
    // IndexedDB bloccato (es. navigazione privata).
  }
  return new MemoryRecentStore();
}
