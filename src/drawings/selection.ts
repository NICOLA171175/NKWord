/** Disegno selezionato: stato della vista, non del documento (non entra in Annulla). */

import type { Id } from '../core';

export class DrawingSelection {
  private current: Id | null = null;
  private readonly listeners = new Set<() => void>();

  get id(): Id | null {
    return this.current;
  }

  select(id: Id | null): void {
    if (id === this.current) return;
    this.current = id;
    for (const l of [...this.listeners]) l();
  }

  onChange(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
}
