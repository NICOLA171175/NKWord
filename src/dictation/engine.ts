/**
 * Come si detta, a seconda di dove gira NKoffice:
 * - `web`: nel browser Chrome o Edge, con il riconoscimento vocale del
 *   browser (Web Speech API). Serve Internet: l'audio viene riconosciuto
 *   dal servizio di Google (Chrome) o Microsoft (Edge), senza chiavi né account.
 * - `windows`: nel programma installato su Windows si apre la dettatura di
 *   Windows (Win+H), che scrive nel documento come la tastiera. Il
 *   riconoscimento del browser in Electron non funziona: manca il servizio
 *   di Google, e risponde sempre con un errore di rete.
 * - `mac`: nel programma per Mac si usa la Dettatura del Mac (Fn Fn).
 * - `none`: browser senza riconoscimento vocale (es. Firefox).
 */

export type DictationMode = 'web' | 'windows' | 'mac' | 'none';

/** Il minimo della Web Speech API che serve qui. */
export interface SpeechRecognitionLike {
  lang: string;
  continuous: boolean;
  interimResults: boolean;
  onresult: ((event: SpeechResultEventLike) => void) | null;
  onerror: ((event: { readonly error: string }) => void) | null;
  onend: (() => void) | null;
  start(): void;
  stop(): void;
  abort(): void;
}

export interface SpeechResultEventLike {
  readonly resultIndex: number;
  readonly results: ArrayLike<{ readonly isFinal: boolean; readonly 0: { readonly transcript: string } }>;
}

export type SpeechRecognitionCtor = new () => SpeechRecognitionLike;

/** Il ponte del programma desktop (desktop/preload.cjs). */
export interface DesktopDictation {
  readonly platform?: string;
  startSystemDictation?(): Promise<unknown>;
}

export interface DictationEnvironment {
  readonly recognition: SpeechRecognitionCtor | null;
  readonly desktop: DesktopDictation | null;
  /** Il PC usa Windows (anche nel browser): per suggerire Win+H. */
  readonly windows: boolean;
}

/** Legge l'ambiente dalla finestra del browser. */
export function detectEnvironment(win: Window | undefined = typeof window === 'undefined' ? undefined : window): DictationEnvironment {
  const w = win as (Window & Record<string, unknown>) | undefined;
  const recognition = (w?.SpeechRecognition ?? w?.webkitSpeechRecognition ?? null) as SpeechRecognitionCtor | null;
  const desktop = (w?.nkofficeDesktop ?? null) as DesktopDictation | null;
  const ua = w?.navigator?.userAgent ?? '';
  return { recognition, desktop, windows: desktop?.platform === 'win32' || /Windows/.test(ua) };
}

export function dictationMode(env: DictationEnvironment): DictationMode {
  if (env.desktop) {
    if (env.desktop.platform === 'win32' && env.desktop.startSystemDictation) return 'windows';
    if (env.desktop.platform === 'darwin') return 'mac';
    return 'none';
  }
  return env.recognition ? 'web' : 'none';
}

/** Messaggio per un errore del riconoscimento del browser; null se si può continuare. */
export function errorMessage(code: string, windows: boolean): string | null {
  const hint = windows ? ' Puoi usare la dettatura di Windows: premi Win+H.' : '';
  switch (code) {
    case 'no-speech':
    case 'aborted':
      return null;
    case 'not-allowed':
      return "Il microfono non è consentito. Clicca sul lucchetto accanto all'indirizzo e consenti il microfono.";
    case 'audio-capture':
      return 'Nessun microfono trovato. Collega un microfono e riprova.';
    case 'network':
      return `La dettatura del browser ha bisogno di Internet: controlla la connessione.${hint}`;
    case 'language-not-supported':
      return `Il riconoscimento vocale in italiano non è disponibile in questo browser.${hint}`;
    default:
      return `Il riconoscimento vocale non è disponibile in questo browser.${hint}`;
  }
}
