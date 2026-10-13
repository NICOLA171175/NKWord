import type { Id, ListDefinition, ListLevel, NumberFormat } from '../../core';

const BULLETS = ['•', '◦', '▪'];
const NUMBERED: NumberFormat[] = ['decimal', 'lowerLetter', 'lowerRoman'];

export function defaultListLevel(format: NumberFormat, level: number, start = 1): ListLevel {
  return {
    format,
    text: format === 'bullet' ? BULLETS[level % BULLETS.length]! : format === 'none' ? '' : `%${level + 1}.`,
    start,
    indentLeft: 36 * (level + 1),
    hanging: 18,
  };
}

/** Definizione di elenco a 9 livelli: puntato o numerato con formati alternati. */
export function defaultListDefinition(id: Id, kind: 'bullet' | 'numbered'): ListDefinition {
  return {
    id,
    levels: Array.from({ length: 9 }, (_, l) =>
      defaultListLevel(kind === 'bullet' ? 'bullet' : NUMBERED[l % NUMBERED.length]!, l),
    ),
  };
}
