import { describe, expect, it } from 'vitest';
import { Keymap, defaultKeybindings, eventKeyNames, formatShortcut, normalizeKey } from '../../src/ui/keymap';
import { COMMAND_IDS } from '../../src/core';

const ev = (key: string, mods: Partial<{ alt: boolean; ctrl: boolean; meta: boolean; shift: boolean }> = {}, code?: string) => ({
  key,
  code,
  altKey: !!mods.alt,
  ctrlKey: !!mods.ctrl,
  metaKey: !!mods.meta,
  shiftKey: !!mods.shift,
});

describe('normalizeKey', () => {
  it('risolve Mod per piattaforma e ordina i modificatori', () => {
    expect(normalizeKey('Mod-b', false)).toBe('Ctrl-b');
    expect(normalizeKey('Mod-b', true)).toBe('Meta-b');
    expect(normalizeKey('shift-Mod-B', false)).toBe('Ctrl-Shift-b');
    expect(normalizeKey('Alt-Mod-1', false)).toBe('Alt-Ctrl-1');
    expect(normalizeKey('Mod--', false)).toBe('Ctrl--');
    expect(normalizeKey('Mod- ', false)).toBe('Ctrl-Space');
    expect(() => normalizeKey('Hyper-x', false)).toThrow();
  });

  it('formatta le scorciatoie in italiano', () => {
    expect(formatShortcut('Ctrl-Shift-b', false)).toBe('Ctrl+Maiusc+B');
    expect(formatShortcut('Meta-z', true)).toBe('⌘Z');
    expect(formatShortcut('Enter', false)).toBe('Invio');
  });
});

describe('Keymap', () => {
  it('riconosce un evento dal carattere e dal tasto fisico', () => {
    expect(eventKeyNames(ev('B', { ctrl: true, shift: true }, 'KeyB'))).toEqual(['Ctrl-Shift-b']);
    // Ctrl+Maiusc+8 produce '*' ma il tasto è '8'.
    expect(eventKeyNames(ev('*', { ctrl: true, shift: true }, 'Digit8'))).toEqual(['Ctrl-Shift-*', 'Ctrl-Shift-8']);
  });

  it('trova le scorciatoie nell\'ordine di definizione', () => {
    const km = new Keymap(defaultKeybindings(false), false);
    expect(km.lookup(ev('b', { ctrl: true }, 'KeyB')).map((b) => b.commandId)).toEqual(['format.bold']);
    expect(km.lookup(ev('Tab')).map((b) => b.commandId)).toEqual(['list.increaseLevel', 'edit.insertTab']);
    expect(km.lookup(ev('*', { ctrl: true, shift: true }, 'Digit8')).map((b) => b.commandId)).toEqual(['view.toggleFormattingMarks']);
    expect(km.lookup(ev('b', { meta: true }, 'KeyB'))).toEqual([]);
  });

  it('su macOS usa Cmd e Opzione per le parole', () => {
    const km = new Keymap(defaultKeybindings(true), true);
    expect(km.lookup(ev('b', { meta: true })).map((b) => b.commandId)).toEqual(['format.bold']);
    expect(km.lookup(ev('ArrowLeft', { alt: true }))[0]?.args).toEqual({ direction: 'wordLeft' });
    expect(km.lookup(ev('ArrowLeft', { meta: true }))[0]?.args).toEqual({ direction: 'lineStart' });
  });

  it('mostra la scorciatoia di un comando con i suoi argomenti', () => {
    const km = new Keymap(defaultKeybindings(false), false);
    expect(km.shortcutFor('format.bold')).toBe('Ctrl+B');
    expect(km.shortcutFor('paragraph.align', { align: 'center' })).toBe('Ctrl+E');
    expect(km.shortcutFor('paragraph.align', { align: 'nessuno' })).toBeUndefined();
  });

  it('usa solo comandi del catalogo', () => {
    for (const b of defaultKeybindings(false)) expect(COMMAND_IDS).toContain(b.commandId);
  });
});
