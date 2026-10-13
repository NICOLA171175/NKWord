/**
 * Tabella dei simboli: i caratteri che non si trovano sulla tastiera,
 * divisi in gruppi come il "Sottoinsieme" della finestra Simbolo di Word.
 */

export interface SymbolCategory {
  readonly id: string;
  readonly label: string;
  readonly chars: readonly string[];
}

const split = (s: string) => [...s.replace(/\s+/g, '')];

export const SYMBOL_CATEGORIES: readonly SymbolCategory[] = [
  { id: 'comuni', label: 'Simboli comuni', chars: split('€ © ® ™ ° ± × ÷ ½ ¼ ¾ § ¶ • … – — « » ¿ ¡ † ‡ ‰ № ✓ ✗ ★ ☆ ♥') },
  { id: 'valute', label: 'Valute', chars: split('€ $ £ ¥ ¢ ₹ ₽ ₩ ₺ ₪ ₫ ₴ ₦ ₱ ₿ ¤ ฿ ₣') },
  { id: 'punteggiatura', label: 'Punteggiatura e tipografia', chars: split('– — ‐ … « » ‹ › “ ” ‘ ’ „ ‚ ¿ ¡ · • ′ ″ § ¶ † ‡ ¦ ¨ ´ ¸ ¯ ˜ ˆ ‰ ※') },
  { id: 'matematica', label: 'Matematica', chars: split('± × ÷ ≠ ≈ ≡ ≤ ≥ ∞ √ ∛ ∑ ∏ ∫ ∂ ∆ ∇ π µ ½ ⅓ ⅔ ¼ ¾ ⅛ ¹ ² ³ ⁴ ⁿ ₀ ₁ ₂ ₃ ∈ ∉ ∩ ∪ ⊂ ⊃ ⊆ ⊇ ∅ ∀ ∃ ¬ ∧ ∨ ∝ ∠ ⊥ ∥ ° ‰ ∴ ∵') },
  { id: 'frecce', label: 'Frecce', chars: split('← → ↑ ↓ ↔ ↕ ↖ ↗ ↘ ↙ ⇐ ⇒ ⇑ ⇓ ⇔ ⇕ ↩ ↪ ↺ ↻ ➔ ➜ ➤ ▶ ◀ ▲ ▼ ► ◄') },
  { id: 'maiuscole', label: 'Lettere maiuscole accentate', chars: split('À Á Â Ã Ä Å Æ Ç È É Ê Ë Ì Í Î Ï Ð Ñ Ò Ó Ô Õ Ö Ø Ù Ú Û Ü Ý Þ Œ Š Ž Ł Ő Ű Č Ř Ğ İ Ş ẞ') },
  { id: 'minuscole', label: 'Lettere minuscole accentate', chars: split('à á â ã ä å æ ç è é ê ë ì í î ï ð ñ ò ó ô õ ö ø ù ú û ü ý ÿ þ œ š ž ł ő ű č ř ğ ı ş ß') },
  { id: 'greco', label: 'Greco', chars: split('Α Β Γ Δ Ε Ζ Η Θ Ι Κ Λ Μ Ν Ξ Ο Π Ρ Σ Τ Υ Φ Χ Ψ Ω α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ ς τ υ φ χ ψ ω') },
  { id: 'forme', label: 'Forme e segni', chars: split('■ □ ▪ ▫ ● ○ ◆ ◇ ▲ △ ▼ ▽ ★ ☆ ✓ ✔ ✗ ✘ ☐ ☑ ☒ ♠ ♣ ♥ ♦ ♪ ♫ ☀ ☁ ☂ ☎ ✉ ✂ ✏ ☺ ☹ ⌘ ⌚ ⚠ ☯ ♻') },
];

/** Nomi in italiano dei simboli più usati (gli altri mostrano solo il codice). */
const NAMES: Readonly<Record<string, string>> = {
  '€': 'Euro', '$': 'Dollaro', '£': 'Sterlina', '¥': 'Yen', '¢': 'Centesimo', '₹': 'Rupia indiana', '₽': 'Rublo', '₩': 'Won',
  '₺': 'Lira turca', '₿': 'Bitcoin', '¤': 'Valuta generica', '₣': 'Franco',
  '©': 'Copyright', '®': 'Marchio registrato', '™': 'Marchio', '°': 'Grado', '±': 'Più o meno', '×': 'Per (moltiplicazione)',
  '÷': 'Diviso', '½': 'Un mezzo', '¼': 'Un quarto', '¾': 'Tre quarti', '⅓': 'Un terzo', '⅔': 'Due terzi', '⅛': 'Un ottavo',
  '§': 'Paragrafo (sezione)', '¶': 'Segno di paragrafo', '•': 'Punto elenco', '…': 'Puntini di sospensione',
  '–': 'Trattino medio (en dash)', '—': 'Trattino lungo (em dash)', '‐': 'Trattino', '«': 'Virgolette basse aperte', '»': 'Virgolette basse chiuse',
  '‹': 'Virgoletta singola aperta', '›': 'Virgoletta singola chiusa', '“': 'Virgolette alte aperte', '”': 'Virgolette alte chiuse',
  '‘': 'Apostrofo aperto', '’': 'Apostrofo', '„': 'Virgolette basse tedesche', '¿': 'Punto interrogativo rovesciato', '¡': 'Punto esclamativo rovesciato',
  '†': 'Croce', '‡': 'Croce doppia', '‰': 'Per mille', '№': 'Numero', '·': 'Punto centrato', '′': 'Primo', '″': 'Secondo',
  '≠': 'Diverso', '≈': 'Circa uguale', '≡': 'Identico', '≤': 'Minore o uguale', '≥': 'Maggiore o uguale', '∞': 'Infinito',
  '√': 'Radice quadrata', '∛': 'Radice cubica', '∑': 'Sommatoria', '∏': 'Produttoria', '∫': 'Integrale', '∂': 'Derivata parziale',
  '∆': 'Delta', '∇': 'Nabla', 'π': 'Pi greco', 'µ': 'Micro', '¹': 'Apice uno', '²': 'Al quadrato', '³': 'Al cubo',
  '∈': 'Appartiene', '∉': 'Non appartiene', '∩': 'Intersezione', '∪': 'Unione', '⊂': 'Sottoinsieme', '∅': 'Insieme vuoto',
  '∀': 'Per ogni', '∃': 'Esiste', '¬': 'Non', '∠': 'Angolo', '⊥': 'Perpendicolare', '∥': 'Parallelo', '∴': 'Quindi',
  '←': 'Freccia a sinistra', '→': 'Freccia a destra', '↑': 'Freccia in su', '↓': 'Freccia in giù', '↔': 'Freccia doppia',
  '↕': 'Freccia su e giù', '⇐': 'Freccia doppia a sinistra', '⇒': 'Implica', '⇔': 'Se e solo se', '↩': 'Freccia di ritorno',
  '↺': 'Senso antiorario', '↻': 'Senso orario', '➔': 'Freccia piena', '▶': 'Triangolo a destra', '◀': 'Triangolo a sinistra',
  '▲': 'Triangolo in su', '▼': 'Triangolo in giù',
  '✓': 'Segno di spunta', '✔': 'Spunta in grassetto', '✗': 'Croce (no)', '✘': 'Croce in grassetto', '★': 'Stella piena', '☆': 'Stella vuota',
  '♥': 'Cuori', '♦': 'Quadri', '♣': 'Fiori', '♠': 'Picche', '♪': 'Nota musicale', '♫': 'Note musicali', '☐': 'Casella vuota',
  '☑': 'Casella con spunta', '☒': 'Casella con croce', '■': 'Quadrato pieno', '□': 'Quadrato vuoto', '●': 'Cerchio pieno',
  '○': 'Cerchio vuoto', '◆': 'Rombo pieno', '◇': 'Rombo vuoto', '☀': 'Sole', '☁': 'Nuvola', '☂': 'Ombrello', '☎': 'Telefono',
  '✉': 'Busta', '✂': 'Forbici', '✏': 'Matita', '☺': 'Faccina sorridente', '☹': 'Faccina triste', '⚠': 'Attenzione', '♻': 'Riciclo',
  'È': 'E maiuscola con accento grave', 'É': 'E maiuscola con accento acuto', 'À': 'A maiuscola con accento grave',
  'Ì': 'I maiuscola con accento grave', 'Ò': 'O maiuscola con accento grave', 'Ù': 'U maiuscola con accento grave',
  'Ç': 'C con cediglia', 'ç': 'c con cediglia', 'Ñ': 'N con tilde', 'ñ': 'n con tilde', 'ß': 'Esse tedesca (eszett)', 'ẞ': 'Esse tedesca maiuscola',
  'Æ': 'Legatura AE', 'æ': 'Legatura ae', 'Œ': 'Legatura OE', 'œ': 'Legatura oe', 'Ø': 'O barrata', 'ø': 'o barrata', 'Å': 'A con anello', 'å': 'a con anello',
  'Ä': 'A con dieresi', 'Ö': 'O con dieresi', 'Ü': 'U con dieresi', 'ä': 'a con dieresi', 'ö': 'o con dieresi', 'ü': 'u con dieresi',
  'Ω': 'Omega maiuscola', 'α': 'Alfa', 'β': 'Beta', 'γ': 'Gamma', 'δ': 'Delta minuscola', 'Δ': 'Delta maiuscola', 'λ': 'Lambda',
  'σ': 'Sigma', 'Σ': 'Sigma maiuscola', 'ω': 'Omega', 'θ': 'Theta', 'φ': 'Phi', 'ε': 'Epsilon', 'μ': 'Mu',
};

/** Codice Unicode in esadecimale, come "Codice carattere" in Word: '€' → '20AC'. */
export function charCode(char: string): string {
  return (char.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0');
}

/** Nome del simbolo da mostrare, es. 'Euro'; se non lo conosciamo, il codice. */
export function symbolName(char: string): string {
  return NAMES[char] ?? `Carattere U+${charCode(char)}`;
}

/**
 * Converte un codice esadecimale scritto dall'utente ('20AC', 'U+20AC',
 * '0x20ac') nel carattere; null se non è un carattere stampabile valido.
 */
export function charFromCode(code: string): string | null {
  const hex = code.trim().replace(/^(u\+|0x)/i, '');
  if (!/^[0-9a-f]{1,6}$/i.test(hex)) return null;
  const cp = parseInt(hex, 16);
  if (cp < 0x20 || (cp >= 0x7f && cp < 0xa0) || (cp >= 0xd800 && cp <= 0xdfff) || cp > 0x10ffff) return null;
  return String.fromCodePoint(cp);
}
