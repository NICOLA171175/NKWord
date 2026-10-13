/**
 * Scrittura a basso livello di un file PDF 1.4: oggetti numerati, flussi
 * binari, tabella xref e trailer. Non sa nulla del documento di nikoffice.
 */

const latin1 = (s: string): Uint8Array => {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
};

/** Numero compatto per il PDF: al più 3 decimali, senza zeri finali. */
export function num(n: number): string {
  if (!Number.isFinite(n)) return '0';
  const s = (Math.round(n * 1000) / 1000).toFixed(3).replace(/\.?0+$/, '');
  return s === '-0' ? '0' : s;
}

/** Riferimento indiretto, es. `5 0 R`. */
export const ref = (id: number): string => `${id} 0 R`;

/** Nome PDF con i caratteri speciali codificati (#xx). */
export function pdfName(name: string): string {
  return `/${name.replace(/[^!-~]|[#()<>[\]{}/%]/g, (c) => `#${c.charCodeAt(0).toString(16).padStart(2, '0')}`)}`;
}

/** Stringa letterale ASCII, con le parentesi e le barre protette. */
export function pdfString(s: string): string {
  if (/^[\x20-\x7e]*$/.test(s)) return `(${s.replace(/[\\()]/g, '\\$&')})`;
  // Testo Unicode: UTF-16BE con BOM, in esadecimale.
  let hex = 'FEFF';
  for (let i = 0; i < s.length; i++) hex += s.charCodeAt(i).toString(16).padStart(4, '0').toUpperCase();
  return `<${hex}>`;
}

export function hexString(bytes: Uint8Array): string {
  let hex = '';
  for (const b of bytes) hex += b.toString(16).padStart(2, '0');
  return `<${hex}>`;
}

export class PdfWriter {
  private readonly objects: (Uint8Array | undefined)[] = [];

  /** Riserva un numero di oggetto (per i riferimenti in avanti). */
  alloc(): number {
    this.objects.push(undefined);
    return this.objects.length;
  }

  /** Oggetto semplice: dizionario, array, numero... */
  set(id: number, body: string): void {
    this.objects[id - 1] = latin1(`${id} 0 obj\n${body}\nendobj\n`);
  }

  /** Flusso: `dict` è il contenuto del dizionario senza `/Length`. */
  setStream(id: number, dict: string, data: Uint8Array | string): void {
    const bytes = typeof data === 'string' ? latin1(data) : data;
    const head = latin1(`${id} 0 obj\n<< ${dict} /Length ${bytes.length} >>\nstream\n`);
    const tail = latin1('\nendstream\nendobj\n');
    const out = new Uint8Array(head.length + bytes.length + tail.length);
    out.set(head, 0);
    out.set(bytes, head.length);
    out.set(tail, head.length + bytes.length);
    this.objects[id - 1] = out;
  }

  add(body: string): number {
    const id = this.alloc();
    this.set(id, body);
    return id;
  }

  addStream(dict: string, data: Uint8Array | string): number {
    const id = this.alloc();
    this.setStream(id, dict, data);
    return id;
  }

  finish(rootId: number, infoId?: number): Uint8Array {
    // Il commento binario segnala ai programmi di trasferimento che il file non è testo.
    const header = latin1('%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');
    const chunks: Uint8Array[] = [header];
    const offsets: number[] = [];
    let pos = header.length;
    this.objects.forEach((obj, i) => {
      if (!obj) throw new Error(`Oggetto PDF ${i + 1} non scritto`);
      offsets.push(pos);
      chunks.push(obj);
      pos += obj.length;
    });
    const size = this.objects.length + 1;
    let xref = `xref\n0 ${size}\n0000000000 65535 f \n`;
    for (const o of offsets) xref += `${String(o).padStart(10, '0')} 00000 n \n`;
    xref += `trailer\n<< /Size ${size} /Root ${ref(rootId)}${infoId ? ` /Info ${ref(infoId)}` : ''} >>\nstartxref\n${pos}\n%%EOF\n`;
    chunks.push(latin1(xref));
    const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let at = 0;
    for (const c of chunks) {
      out.set(c, at);
      at += c.length;
    }
    return out;
  }
}
