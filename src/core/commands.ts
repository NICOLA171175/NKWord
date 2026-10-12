/**
 * Sistema di comandi. Ogni azione dell'utente (digitare, grassetto, inserire
 * tabella, salvare...) è un comando registrato con un id stabile. UI,
 * scorciatoie da tastiera e macro invocano i comandi solo tramite id.
 *
 * Convenzione per gli id: `<area>.<azione>`, es. `format.bold`,
 * `edit.insertText`, `table.insert`, `file.save`.
 */

import type { EditorState, Transaction } from './state';

export interface CommandContext {
  readonly state: EditorState;
}

export interface Command<Args = void> {
  readonly id: string;
  /** Nome mostrato nella UI (italiano). */
  readonly label: string;
  /** Il comando è eseguibile nello stato attuale? (default: sì) */
  isEnabled?(ctx: CommandContext, args: Args): boolean;
  /** Per i pulsanti a interruttore (es. grassetto attivo sulla selezione). */
  isActive?(ctx: CommandContext): boolean;
  /**
   * Calcola la transazione. Deve essere pura: nessun effetto collaterale,
   * nessuna modifica dello stato ricevuto. `null` significa "nessun effetto".
   */
  run(ctx: CommandContext, args: Args): Transaction | null;
}

/** Comando con effetti esterni (apri file, stampa...), non puro. */
export interface AsyncCommand<Args = void> {
  readonly id: string;
  readonly label: string;
  readonly async: true;
  isEnabled?(ctx: CommandContext, args: Args): boolean;
  run(ctx: CommandContext, args: Args): Promise<Transaction | null>;
}

export type AnyCommand = Command<any> | AsyncCommand<any>;

export function isAsyncCommand(c: AnyCommand): c is AsyncCommand<unknown> {
  return 'async' in c && c.async === true;
}

export class CommandRegistry {
  private readonly commands = new Map<string, AnyCommand>();

  register(command: AnyCommand): () => void {
    if (this.commands.has(command.id)) throw new Error(`Comando già registrato: ${command.id}`);
    this.commands.set(command.id, command);
    return () => {
      if (this.commands.get(command.id) === command) this.commands.delete(command.id);
    };
  }

  get(id: string): AnyCommand | undefined {
    return this.commands.get(id);
  }

  has(id: string): boolean {
    return this.commands.has(id);
  }

  list(): AnyCommand[] {
    return [...this.commands.values()];
  }
}
