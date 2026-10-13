/**
 * Servizio di traduzione. Si usa Google Traduttore attraverso il suo
 * indirizzo pubblico gratuito (lo stesso delle estensioni dei browser):
 * non servono chiavi né account, basta la connessione a Internet. Il testo
 * da tradurre viene inviato a Google, come fa Word con Microsoft Translator.
 *
 * Più pezzi di testo viaggiano insieme in una sola richiesta, uno per
 * riga; se il servizio restituisce un numero di righe diverso si ripiega
 * su una richiesta per pezzo.
 */

import { AUTO_DETECT } from './languages';

export interface TranslationResult {
  /** Un testo tradotto per ogni testo ricevuto, nello stesso ordine. */
  readonly texts: readonly string[];
  /** Lingua riconosciuta dal servizio (quando "Da" è Rileva automaticamente). */
  readonly detected?: string;
}

export interface Translator {
  translate(texts: readonly string[], to: string, from?: string, signal?: AbortSignal): Promise<TranslationResult>;
}

export type FetchLike = (url: string, init: { method: string; headers: Record<string, string>; body: string; signal?: AbortSignal }) => Promise<{
  readonly ok: boolean;
  readonly status: number;
  json(): Promise<unknown>;
}>;

export const GOOGLE_ENDPOINT = 'https://translate.googleapis.com/translate_a/single';
/** Caratteri per richiesta: oltre, il servizio rifiuta o taglia. */
export const MAX_REQUEST_CHARS = 4000;

/** Errore mostrato all'utente così com'è. */
export class TranslationError extends Error {}

/** Divide i testi in gruppi che stanno in una richiesta. */
export function batches(texts: readonly string[], max = MAX_REQUEST_CHARS): number[][] {
  const out: number[][] = [];
  let current: number[] = [];
  let size = 0;
  texts.forEach((t, i) => {
    if (current.length && size + t.length + 1 > max) {
      out.push(current);
      current = [];
      size = 0;
    }
    current.push(i);
    size += t.length + 1;
  });
  if (current.length) out.push(current);
  return out;
}

/** Legge la risposta di Google: `[[["tradotto","originale",...],...],null,"it",...]`. */
export function parseGoogleResponse(data: unknown): { text: string; detected?: string } {
  if (!Array.isArray(data) || !Array.isArray(data[0])) throw new TranslationError('Risposta del servizio di traduzione non valida.');
  let text = '';
  for (const chunk of data[0] as unknown[]) {
    if (Array.isArray(chunk) && typeof chunk[0] === 'string') text += chunk[0];
  }
  const detected = typeof data[2] === 'string' ? data[2] : undefined;
  return detected ? { text, detected } : { text };
}

export function createGoogleTranslator(fetchFn: FetchLike = (url, init) => globalThis.fetch(url, init)): Translator {
  const request = async (text: string, to: string, from: string, signal?: AbortSignal) => {
    const params = new URLSearchParams({ client: 'gtx', sl: from, tl: to, dt: 't' });
    let response;
    try {
      response = await fetchFn(`${GOOGLE_ENDPOINT}?${params}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
        body: new URLSearchParams({ q: text }).toString(),
        signal,
      });
    } catch (error) {
      if ((error as { name?: string })?.name === 'AbortError') throw error;
      throw new TranslationError('Impossibile raggiungere il servizio di traduzione. Controlla la connessione a Internet e riprova.');
    }
    if (response.status === 429) throw new TranslationError('Il servizio di traduzione ha ricevuto troppe richieste. Riprova fra qualche minuto.');
    if (!response.ok) throw new TranslationError(`Il servizio di traduzione non ha risposto (errore ${response.status}). Riprova più tardi.`);
    return parseGoogleResponse(await response.json());
  };

  return {
    async translate(texts, to, from = AUTO_DETECT, signal) {
      const out: string[] = new Array<string>(texts.length).fill('');
      let detected: string | undefined;
      for (const group of batches(texts)) {
        const joined = group.map((i) => texts[i]!).join('\n');
        const result = await request(joined, to, from, signal);
        detected ??= result.detected;
        const lines = result.text.split('\n');
        if (lines.length === group.length) {
          group.forEach((i, k) => (out[i] = lines[k]!));
          continue;
        }
        for (const i of group) out[i] = (await request(texts[i]!, to, from, signal)).text;
      }
      return detected ? { texts: out, detected } : { texts: out };
    },
  };
}
