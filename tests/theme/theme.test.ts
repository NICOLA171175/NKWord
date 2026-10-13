// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { Editor } from '../../src/core';
import {
  createThemeModule,
  loadTheme,
  resolveTheme,
  ribbonItems,
  saveTheme,
  type SystemTheme,
  type ThemeStorage,
} from '../../src/theme';
import * as themeExports from '../../src/theme';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { createApp } from '../../src/ui/shell';
import { createUiModule } from '../../src/ui/module';

afterEach(() => {
  document.body.replaceChildren();
  delete document.documentElement.dataset.theme;
});

class MapStorage implements ThemeStorage {
  readonly data = new Map<string, string>();
  getItem(k: string) {
    return this.data.get(k) ?? null;
  }
  setItem(k: string, v: string) {
    this.data.set(k, v);
  }
}

class FakeSystem implements SystemTheme {
  dark = false;
  private readonly listeners = new Set<() => void>();
  set(dark: boolean) {
    this.dark = dark;
    for (const l of [...this.listeners]) l();
  }
  onChange(listener: () => void) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
}

function setup(storage = new MapStorage(), system = new FakeSystem()) {
  const root = document.createElement('div');
  const editor = new Editor();
  const theme = createThemeModule({ storage, system, root });
  editor.use(theme);
  return { editor, theme, root, storage, system };
}

describe('scelta del tema', () => {
  it('automatico segue il sistema', () => {
    expect(resolveTheme('auto', true)).toBe('dark');
    expect(resolveTheme('auto', false)).toBe('light');
    expect(resolveTheme('dark', false)).toBe('dark');
    expect(resolveTheme('light', true)).toBe('light');
  });

  it('si ricorda fra un avvio e l’altro; valori strani tornano ad automatico', () => {
    const storage = new MapStorage();
    expect(loadTheme(storage)).toBe('auto');
    saveTheme(storage, 'dark');
    expect(loadTheme(storage)).toBe('dark');
    storage.setItem('nkoffice-tema', 'viola');
    expect(loadTheme(storage)).toBe('auto');
    expect(loadTheme(null)).toBe('auto');
  });
});

describe('comando view.theme', () => {
  it('applica il tema all’interfaccia e lo salva', () => {
    const { editor, theme, root, storage } = setup();
    expect(root.dataset.theme).toBe('light');
    expect(editor.canExecute('view.theme', { theme: 'dark' })).toBe(true);
    expect(editor.canExecute('view.theme', { theme: 'viola' })).toBe(false);
    void editor.executeAsync('view.theme', { theme: 'dark' });
    expect(theme.choice).toBe('dark');
    expect(root.dataset.theme).toBe('dark');
    expect(storage.getItem('nkoffice-tema')).toBe('dark');
  });

  it('riapre con il tema scelto', () => {
    const storage = new MapStorage();
    storage.setItem('nkoffice-tema', 'dark');
    const { root, theme } = setup(storage);
    expect(theme.choice).toBe('dark');
    expect(root.dataset.theme).toBe('dark');
  });

  it('in automatico cambia quando cambia il sistema, altrimenti no', () => {
    const { theme, root, system } = setup();
    system.set(true);
    expect(root.dataset.theme).toBe('dark');
    theme.setTheme('light');
    system.set(false);
    system.set(true);
    expect(root.dataset.theme).toBe('light');
  });

  it('non entra in Annulla', () => {
    const { editor } = setup();
    void editor.executeAsync('view.theme', { theme: 'dark' });
    expect(editor.canExecute('edit.undo')).toBe(false);
  });
});

describe('barra multifunzione', () => {
  it('il modulo si trova da solo e ha il menu Tema in Visualizza', () => {
    const modules = collectModules([themeExports as Record<string, unknown>]);
    expect(modules.map((m) => m.id)).toEqual(['theme']);
    const items = collectRibbonItems([themeExports as Record<string, unknown>]);
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({ tab: 'view', commandId: 'view.theme', kind: 'dropdown' });
    expect(ribbonItems[0]!.options.map((o) => o.label)).toEqual(['Chiaro', 'Scuro', 'Automatico (come il sistema)']);
  });

  it('il menu mostra la scelta attuale e la cambia', () => {
    const { editor, root } = setup();
    const ui = createUiModule();
    editor.use(ui);
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp(host, editor, { ui, ribbonItems: collectRibbonItems([themeExports as Record<string, unknown>]) });
    app.ribbon.selectTab('view');
    const select = app.ribbon.element.querySelector('select[aria-label="Tema"]') as HTMLSelectElement;
    expect(select).not.toBeNull();
    expect(select.options[select.selectedIndex]!.textContent).toBe('Automatico (come il sistema)');
    select.value = '1';
    select.dispatchEvent(new Event('change'));
    expect(root.dataset.theme).toBe('dark');
    expect(select.options[select.selectedIndex]!.textContent).toBe('Scuro');
  });
});
