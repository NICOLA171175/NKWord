/**
 * Calcolo dei marcatori degli elenchi (punti e numeri) per tutto il
 * documento: la numerazione dipende dai paragrafi precedenti dello stesso
 * elenco, quindi va calcolata in un unico passaggio in ordine di lettura.
 */

import type { DocumentModel, Id, NumberFormat } from '../core';
import { paragraphs, resolveCharFormat, resolveParaFormat } from '../core';
import type { MarkerInfo } from './paragraph';

function toRoman(n: number): string {
  const table: [number, string][] = [
    [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
    [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
  ];
  let out = '';
  for (const [v, s] of table) while (n >= v) (out += s), (n -= v);
  return out;
}

function toLetters(n: number): string {
  // a..z, poi aa, bb... come Word.
  if (n <= 0) return '';
  const letter = String.fromCharCode(97 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

export function formatNumber(n: number, format: NumberFormat): string {
  switch (format) {
    case 'decimal':
      return String(n);
    case 'lowerLetter':
      return toLetters(n);
    case 'upperLetter':
      return toLetters(n).toUpperCase();
    case 'lowerRoman':
      return toRoman(n);
    case 'upperRoman':
      return toRoman(n).toUpperCase();
    case 'bullet':
    case 'none':
      return '';
  }
}

/** Marcatore di ogni paragrafo di elenco, per id del paragrafo. */
export function computeMarkers(doc: DocumentModel): Map<Id, MarkerInfo> {
  const out = new Map<Id, MarkerInfo>();
  const counters = new Map<Id, number[]>();
  for (const { paragraph: p } of paragraphs(doc)) {
    const ref = resolveParaFormat(doc.styles, p).list;
    if (!ref) continue;
    const def = doc.lists[ref.listId];
    const level = def?.levels[ref.level];
    if (!def || !level) continue;
    let c = counters.get(ref.listId);
    if (!c) counters.set(ref.listId, (c = []));
    // Avanza il livello corrente e azzera quelli più profondi.
    c[ref.level] = (c[ref.level] ?? level.start - 1) + 1;
    c.length = ref.level + 1;
    let text: string;
    if (level.format === 'bullet') text = level.text;
    else if (level.format === 'none') text = level.text.replace(/%\d/g, '');
    else {
      text = level.text.replace(/%(\d)/g, (_, d: string) => {
        const l = Number(d) - 1;
        const lv = def.levels[l];
        const value = c![l] ?? lv?.start ?? 1;
        return formatNumber(value, lv?.format === 'bullet' || !lv ? 'decimal' : lv.format);
      });
    }
    const base = resolveCharFormat(doc.styles, p);
    out.set(p.id, {
      text,
      // Il marcatore prende il carattere del segno di paragrafo, senza sottolineatura.
      format: { ...base, underline: null, ...level.charFormat },
      indentLeft: level.indentLeft,
      hanging: level.hanging,
    });
  }
  return out;
}
