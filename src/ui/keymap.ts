/**
 * Scorciatoie da tastiera. Una scorciatoia è un `Keybinding` del core:
 * `{ key: 'Mod-b', commandId: 'format.bold' }`. 'Mod' è Ctrl su
 * Windows/Linux e Cmd su macOS.
 *
 * Più scorciatoie possono condividere un tasto: vince la prima il cui
 * comando è eseguibile (es. Tab aumenta il livello dell'elenco se il cursore
 * è in un elenco, altrimenti inserisce una tabulazione).
 */

import type { Keybinding } from '../core';

const MODIFIER_ORDER = ['Alt', 'Ctrl', 'Meta', 'Shift'] as const;
type Modifier = (typeof MODIFIER_ORDER)[number];

export interface KeyLike {
  readonly key: string;
  readonly code?: string;
  readonly altKey: boolean;
  readonly ctrlKey: boolean;
  readonly metaKey: boolean;
  readonly shiftKey: boolean;
}

export function isMacPlatform(): boolean {
  const nav = globalThis.navigator as (Navigator & { userAgentData?: { platform?: string } }) | undefined;
  const platform = nav?.userAgentData?.platform ?? nav?.platform ?? '';
  return /mac|iphone|ipad/i.test(platform);
}

function normalizeKeyName(name: string): string {
  if (name === ' ' || name.toLowerCase() === 'space') return 'Space';
  if (name.length === 1) return name.toLowerCase();
  if (name.toLowerCase() === 'esc') return 'Escape';
  return name[0]!.toUpperCase() + name.slice(1);
}

/**
 * Forma canonica di una scorciatoia: modificatori in ordine fisso, 'Mod'
 * risolto per la piattaforma, tasto normalizzato. Es. 'shift-Mod-B' → 'Ctrl-Shift-b'.
 */
export function normalizeKey(spec: string, mac = isMacPlatform()): string {
  // Il tasto '-' si scrive come ultimo segmento vuoto: 'Mod--'.
  const parts = spec.endsWith('--') ? [...spec.slice(0, -2).split('-'), '-'] : spec.split('-');
  const key = parts.pop();
  if (!key) throw new Error(`Scorciatoia non valida: ${spec}`);
  const mods = new Set<Modifier>();
  for (const raw of parts) {
    const m = raw.toLowerCase();
    if (m === 'mod') mods.add(mac ? 'Meta' : 'Ctrl');
    else if (m === 'ctrl' || m === 'control') mods.add('Ctrl');
    else if (m === 'cmd' || m === 'meta') mods.add('Meta');
    else if (m === 'alt' || m === 'option') mods.add('Alt');
    else if (m === 'shift') mods.add('Shift');
    else throw new Error(`Modificatore sconosciuto in ${spec}: ${raw}`);
  }
  return [...MODIFIER_ORDER.filter((m) => mods.has(m)), normalizeKeyName(key)].join('-');
}

const CODE_KEYS: Record<string, string> = {
  Minus: '-',
  Equal: '=',
  BracketLeft: '[',
  BracketRight: ']',
  Semicolon: ';',
  Quote: "'",
  Backquote: '`',
  Backslash: '\\',
  Comma: ',',
  Period: '.',
  Slash: '/',
  Space: 'Space',
};

function keyFromCode(code: string | undefined): string | undefined {
  if (!code) return undefined;
  if (/^Key[A-Z]$/.test(code)) return code.slice(3).toLowerCase();
  if (/^Digit[0-9]$/.test(code)) return code.slice(5);
  if (/^Numpad[0-9]$/.test(code)) return code.slice(6);
  return CODE_KEYS[code];
}

/**
 * Nomi canonici di un evento, dal più al meno specifico: prima il carattere
 * prodotto (`key`), poi il tasto fisico (`code`), utile con Maiusc o con
 * layout non americani (es. Ctrl+Maiusc+8 produce '*' ma il tasto è '8').
 */
export function eventKeyNames(event: KeyLike): string[] {
  const mods = MODIFIER_ORDER.filter((m) =>
    m === 'Alt' ? event.altKey : m === 'Ctrl' ? event.ctrlKey : m === 'Meta' ? event.metaKey : event.shiftKey,
  );
  const names = new Set<string>();
  if (event.key && event.key !== 'Unidentified' && event.key !== 'Dead') {
    names.add([...mods, normalizeKeyName(event.key)].join('-'));
  }
  const fromCode = keyFromCode(event.code);
  if (fromCode) names.add([...mods, normalizeKeyName(fromCode)].join('-'));
  return [...names];
}

export class Keymap {
  private readonly byKey = new Map<string, Keybinding[]>();

  constructor(
    bindings: readonly Keybinding[] = [],
    private readonly mac = isMacPlatform(),
  ) {
    for (const b of bindings) this.add(b);
  }

  add(binding: Keybinding): () => void {
    const key = normalizeKey(binding.key, this.mac);
    let list = this.byKey.get(key);
    if (!list) this.byKey.set(key, (list = []));
    list.push(binding);
    return () => {
      const i = list.indexOf(binding);
      if (i >= 0) list.splice(i, 1);
    };
  }

  /** Scorciatoie candidate per un evento, nell'ordine in cui provarle. */
  lookup(event: KeyLike): Keybinding[] {
    const out: Keybinding[] = [];
    for (const name of eventKeyNames(event)) out.push(...(this.byKey.get(name) ?? []));
    return out;
  }

  /** Scorciatoia da mostrare nei suggerimenti di un comando, se esiste. */
  shortcutFor(commandId: string, args?: unknown): string | undefined {
    for (const [key, list] of this.byKey) {
      if (list.some((b) => b.commandId === commandId && sameArgs(b.args, args))) return formatShortcut(key, this.mac);
    }
    return undefined;
  }
}

export function sameArgs(a: unknown, b: unknown): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

/** Testo leggibile di una scorciatoia canonica, es. 'Ctrl-Shift-b' → 'Ctrl+Maiusc+B'. */
export function formatShortcut(normalized: string, mac = isMacPlatform()): string {
  const parts = normalized.endsWith('--') ? [...normalized.slice(0, -2).split('-'), '-'] : normalized.split('-');
  const names: Record<string, string> = mac
    ? { Alt: '⌥', Ctrl: '⌃', Meta: '⌘', Shift: '⇧' }
    : { Alt: 'Alt', Ctrl: 'Ctrl', Meta: 'Win', Shift: 'Maiusc' };
  const keys: Record<string, string> = { Space: 'Spazio', Enter: 'Invio', Backspace: 'Backspace', Delete: 'Canc', Escape: 'Esc' };
  const out = parts.map((p, i) =>
    i < parts.length - 1 ? names[p] ?? p : keys[p] ?? (p.length === 1 ? p.toUpperCase() : p),
  );
  return out.join(mac ? '' : '+');
}

/** Scorciatoie predefinite, sul modello di Word. */
export function defaultKeybindings(mac = isMacPlatform()): Keybinding[] {
  const word = mac ? 'Alt' : 'Mod';
  const line = (dir: 'lineStart' | 'lineEnd'): Keybinding[] =>
    mac ? [{ key: dir === 'lineStart' ? 'Mod-ArrowLeft' : 'Mod-ArrowRight', commandId: 'edit.moveCaret', args: { direction: dir } }] : [];
  const caret = (key: string, direction: string): Keybinding[] => [
    { key, commandId: 'edit.moveCaret', args: { direction } },
    { key: `Shift-${key}`, commandId: 'edit.moveCaret', args: { direction, extend: true } },
  ];
  return [
    // Cronologia
    { key: 'Mod-z', commandId: 'edit.undo' },
    { key: 'Mod-y', commandId: 'edit.redo' },
    { key: 'Mod-Shift-z', commandId: 'edit.redo' },

    // Testo
    { key: 'Enter', commandId: 'edit.splitParagraph' },
    { key: 'Shift-Enter', commandId: 'edit.insertLineBreak' },
    { key: 'Tab', commandId: 'list.increaseLevel' },
    { key: 'Tab', commandId: 'edit.insertTab' },
    { key: 'Shift-Tab', commandId: 'list.decreaseLevel' },
    { key: 'Shift-Tab', commandId: 'paragraph.outdent' },
    { key: 'Backspace', commandId: 'edit.deleteBackward' },
    { key: 'Shift-Backspace', commandId: 'edit.deleteBackward' },
    { key: `${word}-Backspace`, commandId: 'edit.deleteBackward', args: { unit: 'word' } },
    { key: 'Delete', commandId: 'edit.deleteForward' },
    { key: `${word}-Delete`, commandId: 'edit.deleteForward', args: { unit: 'word' } },
    { key: 'Mod-a', commandId: 'edit.selectAll' },

    // Cursore (su/giù dipendono dall'impaginazione: non ci sono ancora comandi)
    ...caret('ArrowLeft', 'left'),
    ...caret('ArrowRight', 'right'),
    ...caret(`${word}-ArrowLeft`, 'wordLeft'),
    ...caret(`${word}-ArrowRight`, 'wordRight'),
    ...caret('Home', 'lineStart'),
    ...caret('End', 'lineEnd'),
    ...line('lineStart'),
    ...line('lineEnd'),
    ...caret('Mod-Home', 'docStart'),
    ...caret('Mod-End', 'docEnd'),

    // Carattere
    { key: 'Mod-b', commandId: 'format.bold' },
    { key: 'Mod-i', commandId: 'format.italic' },
    { key: 'Mod-u', commandId: 'format.underline' },
    { key: 'Mod-=', commandId: 'format.subscript' },
    { key: 'Mod-Shift-=', commandId: 'format.superscript' },
    { key: 'Mod-]', commandId: 'format.growFont' },
    { key: 'Mod-[', commandId: 'format.shrinkFont' },
    { key: 'Mod-Shift-.', commandId: 'format.growFont' },
    { key: 'Mod-Shift-,', commandId: 'format.shrinkFont' },
    { key: 'Mod-Space', commandId: 'format.clear', args: { charOnly: true } },

    // Paragrafo
    { key: 'Mod-l', commandId: 'paragraph.align', args: { align: 'left' } },
    { key: 'Mod-e', commandId: 'paragraph.align', args: { align: 'center' } },
    { key: 'Mod-r', commandId: 'paragraph.align', args: { align: 'right' } },
    { key: 'Mod-j', commandId: 'paragraph.align', args: { align: 'justify' } },
    { key: 'Mod-m', commandId: 'paragraph.indent' },
    { key: 'Mod-Shift-m', commandId: 'paragraph.outdent' },
    { key: 'Mod-1', commandId: 'paragraph.lineSpacing', args: { lineSpacing: { rule: 'multiple', value: 1 } } },
    { key: 'Mod-5', commandId: 'paragraph.lineSpacing', args: { lineSpacing: { rule: 'multiple', value: 1.5 } } },
    { key: 'Mod-2', commandId: 'paragraph.lineSpacing', args: { lineSpacing: { rule: 'multiple', value: 2 } } },
    { key: 'Mod-Alt-1', commandId: 'style.apply', args: { styleId: 'Heading1' } },
    { key: 'Mod-Alt-2', commandId: 'style.apply', args: { styleId: 'Heading2' } },
    { key: 'Mod-Alt-3', commandId: 'style.apply', args: { styleId: 'Heading3' } },
    { key: 'Mod-Shift-n', commandId: 'style.apply', args: { styleId: 'Normal' } },
    { key: 'Mod-Shift-l', commandId: 'list.toggleBullet' },

    // Inserisci
    { key: 'Mod-Enter', commandId: 'insert.pageBreak' },

    // File e vista
    { key: 'Mod-n', commandId: 'file.new' },
    { key: 'Mod-p', commandId: 'view.print' },
    { key: 'Mod-Shift-8', commandId: 'view.toggleFormattingMarks' },
  ];
}
