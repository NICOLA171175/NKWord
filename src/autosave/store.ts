/**
 * Archivio delle copie di recupero. Nel browser e nella versione desktop
 * (Electron) si usa IndexedDB, che regge anche documenti con immagini; se
 * manca si ripiega su localStorage e, come ultima scelta, sulla memoria
 * (utile nei test).
 */

/** Copia di recupero di un documento non salvato. */
export interface RecoverySnapshot {
  /** Id della finestra che l'ha scritta: una copia per finestra aperta. */
  readonly id: string;
  /** Titolo del documento al momento della copia ('' se non ha titolo). */
  readonly title: string;
  /** Data ISO 8601 dell'ultima copia. */
  readonly savedAt: string;
  /** Il `DocumentModel` serializzato in JSON. */
  readonly doc: string;
}

export interface RecoveryStore {
  list(): Promise<RecoverySnapshot[]>;
  put(snapshot: RecoverySnapshot): Promise<void>;
  remove(id: string): Promise<void>;
}

export class MemoryRecoveryStore implements RecoveryStore {
  private readonly items = new Map<string, RecoverySnapshot>();
  async list() {
    return [...this.items.values()];
  }
  async put(snapshot: RecoverySnapshot) {
    this.items.set(snapshot.id, snapshot);
  }
  async remove(id: string) {
    this.items.delete(id);
  }
}

const DB_NAME = 'nkoffice-recupero';
const STORE = 'documenti';

export class IndexedDbRecoveryStore implements RecoveryStore {
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
    return this.run<RecoverySnapshot[]>('readonly', (s) => s.getAll());
  }
  async put(snapshot: RecoverySnapshot) {
    await this.run('readwrite', (s) => s.put(snapshot));
  }
  async remove(id: string) {
    await this.run('readwrite', (s) => s.delete(id));
  }
}

const LS_PREFIX = 'nkoffice.recupero.';

export class LocalStorageRecoveryStore implements RecoveryStore {
  constructor(private readonly storage: Storage) {}

  async list() {
    const out: RecoverySnapshot[] = [];
    for (let i = 0; i < this.storage.length; i++) {
      const key = this.storage.key(i);
      if (!key?.startsWith(LS_PREFIX)) continue;
      try {
        out.push(JSON.parse(this.storage.getItem(key) ?? '') as RecoverySnapshot);
      } catch {
        // Copia illeggibile: la si ignora.
      }
    }
    return out;
  }
  async put(snapshot: RecoverySnapshot) {
    this.storage.setItem(LS_PREFIX + snapshot.id, JSON.stringify(snapshot));
  }
  async remove(id: string) {
    this.storage.removeItem(LS_PREFIX + id);
  }
}

/** L'archivio migliore disponibile in questo ambiente. */
export function defaultRecoveryStore(): RecoveryStore {
  try {
    if (typeof indexedDB !== 'undefined') return new IndexedDbRecoveryStore(indexedDB);
  } catch {
    // IndexedDB bloccato (es. navigazione privata): si prova localStorage.
  }
  try {
    if (typeof localStorage !== 'undefined') return new LocalStorageRecoveryStore(localStorage);
  } catch {
    // localStorage bloccato.
  }
  return new MemoryRecoveryStore();
}
