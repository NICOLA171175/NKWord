/**
 * `Editor`: il punto di raccordo fra i moduli. Tiene lo stato corrente, la
 * cronologia e il registro dei comandi; notifica i cambiamenti. Non conosce
 * DOM, impaginazione o formati di file: quelli arrivano come moduli
 * (`EditorModule`) che si registrano qui.
 */

import { CommandRegistry, isAsyncCommand, type AnyCommand, type CommandContext } from './commands';
import { Emitter } from './events';
import { History } from './history';
import type { EditorModule, EditorServices } from './interfaces';
import { createState, type EditorState, type Transaction, type TransactionMeta } from './state';

export interface EditorEvents extends Record<string, unknown> {
  /** Dopo ogni cambiamento di stato (documento o selezione). */
  change: { readonly state: EditorState; readonly previous: EditorState; readonly meta: TransactionMeta };
  /** Solo quando cambia il documento (non la sola selezione). */
  docChange: { readonly state: EditorState; readonly previous: EditorState };
  /** Un comando è stato eseguito senza errori (anche se non ha cambiato lo stato, es. `file.export`). */
  executed: { readonly commandId: string; readonly args: unknown };
  /** Un comando è fallito con un'eccezione. */
  error: { readonly commandId: string; readonly error: unknown };
}

/**
 * Filtro di transazione: riceve ogni transazione prima che venga applicata
 * e può restituirne una diversa (es. il rilevamento modifiche trasforma le
 * cancellazioni in testo barrato). Deve essere puro come un comando.
 */
export type TransactionFilter = (tr: Transaction, state: EditorState) => Transaction;

export class Editor {
  readonly commands = new CommandRegistry();
  readonly events = new Emitter<EditorEvents>();
  readonly history = new History();
  /** Servizi forniti dai moduli (impaginazione, formati di file...). */
  readonly services: EditorServices = {};

  private _state: EditorState;
  private readonly disposers: (() => void)[] = [];
  private readonly filters: TransactionFilter[] = [];

  constructor(state: EditorState = createState()) {
    this._state = state;
    // Annulla/Ripeti sono comandi come gli altri, così UI e tastiera li
    // invocano per id. Sono gli unici che toccano la cronologia.
    this.commands.register({
      id: 'edit.undo',
      label: 'Annulla',
      isEnabled: () => this.history.canUndo,
      run: ({ state }) => {
        const prev = this.history.undo(state);
        return prev && { state: prev, meta: { addToHistory: false, label: 'undo' } };
      },
    });
    this.commands.register({
      id: 'edit.redo',
      label: 'Ripeti',
      isEnabled: () => this.history.canRedo,
      run: ({ state }) => {
        const next = this.history.redo(state);
        return next && { state: next, meta: { addToHistory: false, label: 'redo' } };
      },
    });
  }

  get state(): EditorState {
    return this._state;
  }

  /** Attiva un modulo; restituisce la funzione per disattivarlo. */
  use(module: EditorModule): () => void {
    const dispose = module.activate(this) ?? (() => {});
    this.disposers.push(dispose);
    return dispose;
  }

  /** Registra un filtro di transazione; restituisce la funzione per toglierlo. */
  addTransactionFilter(filter: TransactionFilter): () => void {
    this.filters.push(filter);
    return () => {
      const i = this.filters.indexOf(filter);
      if (i >= 0) this.filters.splice(i, 1);
    };
  }

  /** Applica una transazione: aggiorna lo stato, la cronologia e notifica. */
  apply(tr: Transaction): void {
    const previous = this._state;
    for (const filter of this.filters) tr = filter(tr, previous);
    if (tr.state === previous) return;
    const meta = tr.meta ?? {};
    if (tr.state.doc !== previous.doc) this.history.record(previous, meta);
    this._state = tr.state;
    this.events.emit('change', { state: tr.state, previous, meta });
    if (tr.state.doc !== previous.doc) this.events.emit('docChange', { state: tr.state, previous });
  }

  /** Sostituisce l'intero stato (es. apertura di un file) e azzera la cronologia. */
  reset(state: EditorState): void {
    const previous = this._state;
    this._state = state;
    this.history.clear();
    const meta: TransactionMeta = { addToHistory: false, label: 'reset' };
    this.events.emit('change', { state, previous, meta });
    this.events.emit('docChange', { state, previous });
  }

  private context(): CommandContext {
    return { state: this._state };
  }

  private require(id: string): AnyCommand {
    const cmd = this.commands.get(id);
    if (!cmd) throw new Error(`Comando sconosciuto: ${id}`);
    return cmd;
  }

  canExecute(id: string, args?: unknown): boolean {
    const cmd = this.commands.get(id);
    if (!cmd) return false;
    return cmd.isEnabled ? cmd.isEnabled(this.context(), args) : true;
  }

  isActive(id: string): boolean {
    const cmd = this.commands.get(id);
    return !!cmd && !isAsyncCommand(cmd) && !!cmd.isActive?.(this.context());
  }

  /** Esegue un comando sincrono. Restituisce true se ha avuto effetto. */
  execute(id: string, args?: unknown): boolean {
    const cmd = this.require(id);
    if (isAsyncCommand(cmd)) throw new Error(`Il comando ${id} è asincrono: usare executeAsync`);
    if (!this.canExecute(id, args)) return false;
    try {
      const tr = cmd.run(this.context(), args);
      if (tr) this.apply(tr);
      this.events.emit('executed', { commandId: id, args });
      return !!tr;
    } catch (error) {
      this.events.emit('error', { commandId: id, error });
      return false;
    }
  }

  /** Esegue un comando sincrono o asincrono. */
  async executeAsync(id: string, args?: unknown): Promise<boolean> {
    const cmd = this.require(id);
    if (!isAsyncCommand(cmd)) return this.execute(id, args);
    if (!this.canExecute(id, args)) return false;
    try {
      const tr = await cmd.run(this.context(), args);
      if (tr) this.apply(tr);
      this.events.emit('executed', { commandId: id, args });
      return !!tr;
    } catch (error) {
      this.events.emit('error', { commandId: id, error });
      return false;
    }
  }

  undo(): boolean {
    return this.execute('edit.undo');
  }

  redo(): boolean {
    return this.execute('edit.redo');
  }

  destroy(): void {
    for (const d of this.disposers.splice(0).reverse()) d();
  }
}
