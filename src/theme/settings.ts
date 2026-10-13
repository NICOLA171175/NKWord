/**
 * Scelta del tema dell'interfaccia (File > Account > Tema di Office in Word):
 * chiaro, scuro o automatico (segue l'impostazione del sistema). La scelta si
 * ricorda fra un avvio e l'altro nella memoria del browser (localStorage).
 */

export type ThemeChoice = 'light' | 'dark' | 'auto';
export type ResolvedTheme = 'light' | 'dark';

export const THEME_CHOICES: readonly ThemeChoice[] = ['light', 'dark', 'auto'];
export const THEME_LABELS: { readonly [K in ThemeChoice]: string } = {
  light: 'Chiaro',
  dark: 'Scuro',
  auto: 'Automatico (come il sistema)',
};
export const DEFAULT_THEME: ThemeChoice = 'auto';
const KEY = 'nkoffice-tema';

/** Dove si conserva la scelta: basta un getItem/setItem (es. localStorage). */
export interface ThemeStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

export function browserStorage(): ThemeStorage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

export function isThemeChoice(value: unknown): value is ThemeChoice {
  return typeof value === 'string' && (THEME_CHOICES as readonly string[]).includes(value);
}

export function loadTheme(storage: ThemeStorage | null): ThemeChoice {
  try {
    const raw = storage?.getItem(KEY);
    return isThemeChoice(raw) ? raw : DEFAULT_THEME;
  } catch {
    return DEFAULT_THEME;
  }
}

export function saveTheme(storage: ThemeStorage | null, choice: ThemeChoice): void {
  try {
    storage?.setItem(KEY, choice);
  } catch {
    // Memoria piena o non disponibile: il tema vale solo per questa sessione.
  }
}

/** Il tema effettivo: "automatico" diventa chiaro o scuro come il sistema. */
export function resolveTheme(choice: ThemeChoice, systemDark: boolean): ResolvedTheme {
  return choice === 'auto' ? (systemDark ? 'dark' : 'light') : choice;
}

/** Il sistema preferisce il tema scuro? Si ascolta per seguirne i cambi. */
export interface SystemTheme {
  readonly dark: boolean;
  onChange(listener: () => void): () => void;
}

export function browserSystemTheme(): SystemTheme {
  const query = typeof globalThis.matchMedia === 'function' ? globalThis.matchMedia('(prefers-color-scheme: dark)') : null;
  return {
    get dark() {
      return query?.matches ?? false;
    },
    onChange(listener) {
      if (!query?.addEventListener) return () => {};
      query.addEventListener('change', listener);
      return () => query.removeEventListener('change', listener);
    },
  };
}
