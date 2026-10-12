import type { EditorState, TransactionMeta } from './state';

interface Entry {
  readonly before: EditorState;
  readonly label?: string;
  readonly mergeKey?: string;
  readonly time: number;
}

/**
 * Cronologia annulla/ripeti basata su istantanee. Grazie all'immutabilità
 * del modello le istantanee condividono quasi tutta la struttura.
 */
export class History {
  private undoStack: Entry[] = [];
  private redoStack: EditorState[] = [];

  constructor(
    private readonly limit = 500,
    private readonly mergeWindowMs = 1000,
  ) {}

  /** Registra il passaggio `before` → (stato nuovo). */
  record(before: EditorState, meta: TransactionMeta = {}, now = Date.now()): void {
    if (meta.addToHistory === false) return;
    this.redoStack = [];
    const last = this.undoStack[this.undoStack.length - 1];
    if (meta.mergeKey && last?.mergeKey === meta.mergeKey && now - last.time <= this.mergeWindowMs) {
      this.undoStack[this.undoStack.length - 1] = { ...last, time: now };
      return;
    }
    this.undoStack.push({ before, label: meta.label, mergeKey: meta.mergeKey, time: now });
    if (this.undoStack.length > this.limit) this.undoStack.shift();
  }

  /** Interrompe la fusione: il prossimo passo sarà separato. */
  seal(): void {
    const last = this.undoStack[this.undoStack.length - 1];
    if (last) this.undoStack[this.undoStack.length - 1] = { ...last, mergeKey: undefined };
  }

  undo(current: EditorState): EditorState | null {
    const entry = this.undoStack.pop();
    if (!entry) return null;
    this.redoStack.push(current);
    return entry.before;
  }

  redo(current: EditorState): EditorState | null {
    const next = this.redoStack.pop();
    if (!next) return null;
    this.undoStack.push({ before: current, time: 0 });
    return next;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  get undoLabel(): string | undefined {
    return this.undoStack[this.undoStack.length - 1]?.label;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}
