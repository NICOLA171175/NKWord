import type { DocumentFormat, FormatRegistry } from '../core';

/** Estensione da un nome di file o da un'estensione (con o senza punto). */
function normalizeExtension(nameOrExt: string): string {
  const s = nameOrExt.trim().toLowerCase();
  const dot = s.lastIndexOf('.');
  return dot >= 0 ? s.slice(dot + 1) : s;
}

export class DefaultFormatRegistry implements FormatRegistry {
  private readonly formats = new Map<string, DocumentFormat>();

  constructor(formats: readonly DocumentFormat[] = []) {
    for (const f of formats) this.register(f);
  }

  register(format: DocumentFormat): void {
    if (this.formats.has(format.id)) throw new Error(`Formato già registrato: ${format.id}`);
    this.formats.set(format.id, format);
  }

  /** Rimuove un formato (es. alla disattivazione di un modulo che l'aveva aggiunto). */
  unregister(id: string): void {
    this.formats.delete(id);
  }

  get(id: string): DocumentFormat | undefined {
    return this.formats.get(id);
  }

  /** Accetta 'docx', '.docx' o un nome di file come 'relazione.docx'. */
  byExtension(ext: string): DocumentFormat | undefined {
    const e = normalizeExtension(ext);
    for (const f of this.formats.values()) if (f.extensions.includes(e)) return f;
    return undefined;
  }

  list(): readonly DocumentFormat[] {
    return [...this.formats.values()];
  }

  /** Formati apribili, per il filtro della finestra "Apri". */
  importable(): readonly DocumentFormat[] {
    return this.list().filter((f) => f.import);
  }

  /** Formati salvabili, per "Salva con nome". */
  exportable(): readonly DocumentFormat[] {
    return this.list().filter((f) => f.export);
  }
}
