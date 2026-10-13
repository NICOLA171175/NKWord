/** Lingue del Traduttore, con il nome in italiano, in ordine alfabetico. */

export interface Language {
  /** Codice della lingua per il servizio (ISO 639-1, salvo il cinese). */
  readonly code: string;
  readonly name: string;
}

export const LANGUAGES: readonly Language[] = [
  { code: 'sq', name: 'Albanese' },
  { code: 'ar', name: 'Arabo' },
  { code: 'bg', name: 'Bulgaro' },
  { code: 'ca', name: 'Catalano' },
  { code: 'cs', name: 'Ceco' },
  { code: 'zh-CN', name: 'Cinese (semplificato)' },
  { code: 'zh-TW', name: 'Cinese (tradizionale)' },
  { code: 'ko', name: 'Coreano' },
  { code: 'hr', name: 'Croato' },
  { code: 'da', name: 'Danese' },
  { code: 'he', name: 'Ebraico' },
  { code: 'et', name: 'Estone' },
  { code: 'fi', name: 'Finlandese' },
  { code: 'fr', name: 'Francese' },
  { code: 'el', name: 'Greco' },
  { code: 'ja', name: 'Giapponese' },
  { code: 'hi', name: 'Hindi' },
  { code: 'id', name: 'Indonesiano' },
  { code: 'en', name: 'Inglese' },
  { code: 'it', name: 'Italiano' },
  { code: 'lv', name: 'Lettone' },
  { code: 'lt', name: 'Lituano' },
  { code: 'no', name: 'Norvegese' },
  { code: 'nl', name: 'Olandese' },
  { code: 'fa', name: 'Persiano' },
  { code: 'pl', name: 'Polacco' },
  { code: 'pt', name: 'Portoghese' },
  { code: 'ro', name: 'Rumeno' },
  { code: 'ru', name: 'Russo' },
  { code: 'sr', name: 'Serbo' },
  { code: 'sk', name: 'Slovacco' },
  { code: 'sl', name: 'Sloveno' },
  { code: 'es', name: 'Spagnolo' },
  { code: 'sv', name: 'Svedese' },
  { code: 'de', name: 'Tedesco' },
  { code: 'th', name: 'Thailandese' },
  { code: 'tr', name: 'Turco' },
  { code: 'uk', name: 'Ucraino' },
  { code: 'hu', name: 'Ungherese' },
  { code: 'vi', name: 'Vietnamita' },
];

/** Valore di "Da" che lascia riconoscere la lingua al servizio. */
export const AUTO_DETECT = 'auto';

/** Nome italiano di una lingua; per i codici sconosciuti restituisce il codice. */
export function languageName(code: string): string {
  if (code === AUTO_DETECT) return 'Rileva automaticamente';
  const lower = code.toLowerCase();
  return (
    LANGUAGES.find((l) => l.code.toLowerCase() === lower)?.name ??
    LANGUAGES.find((l) => l.code.toLowerCase() === lower.split('-')[0])?.name ??
    code
  );
}
