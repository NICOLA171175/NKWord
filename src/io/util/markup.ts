/**
 * Parser di markup minimo, senza DOM: serve sia per l'XML di DOCX sia per
 * l'HTML (anche malformato, come quello degli appunti). In modalità HTML i
 * nomi sono in minuscolo, gli elementi vuoti (`br`, `img`...) non hanno
 * figli e alcuni tag chiudono implicitamente quelli aperti (`<p>`, `<li>`,
 * `<td>`...), come fanno i browser.
 */

export interface MarkupElement {
  readonly name: string;
  readonly attrs: Readonly<Record<string, string>>;
  readonly children: MarkupNode[];
}

export type MarkupNode = MarkupElement | string;

const VOID = new Set([
  'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr',
]);
const RAW_TEXT = new Set(['script', 'style', 'textarea', 'title', 'xmp']);

const BLOCK_CLOSES_P = [
  'address', 'article', 'aside', 'blockquote', 'div', 'dl', 'fieldset', 'figure', 'footer', 'form', 'h1', 'h2',
  'h3', 'h4', 'h5', 'h6', 'header', 'hr', 'main', 'nav', 'ol', 'p', 'pre', 'section', 'table', 'ul',
];

/** Tag aperti che un nuovo tag chiude implicitamente (solo in cima alla pila). */
const IMPLICIT_CLOSE: Record<string, readonly string[]> = {
  li: ['li', 'p'],
  dt: ['dt', 'dd', 'p'],
  dd: ['dt', 'dd', 'p'],
  tr: ['tr', 'td', 'th', 'p'],
  td: ['td', 'th', 'p'],
  th: ['td', 'th', 'p'],
  thead: ['thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'p'],
  tbody: ['thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'p'],
  tfoot: ['thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'p'],
  option: ['option'],
};
for (const t of BLOCK_CLOSES_P) IMPLICIT_CLOSE[t] ??= ['p'];

const NAMED_ENTITIES: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', copy: '©', reg: '®', trade: '™',
  hellip: '…', mdash: '—', ndash: '–', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', laquo: '«',
  raquo: '»', bull: '•', middot: '·', euro: '€', pound: '£', yen: '¥', cent: '¢', sect: '§', para: '¶',
  deg: '°', plusmn: '±', times: '×', divide: '÷', frac12: '½', frac14: '¼', frac34: '¾', shy: '­',
  ensp: ' ', emsp: ' ', thinsp: ' ', zwnj: '‌', zwj: '‍', iexcl: '¡', iquest: '¿',
  agrave: 'à', aacute: 'á', acirc: 'â', atilde: 'ã', auml: 'ä', aring: 'å', aelig: 'æ', ccedil: 'ç',
  egrave: 'è', eacute: 'é', ecirc: 'ê', euml: 'ë', igrave: 'ì', iacute: 'í', icirc: 'î', iuml: 'ï',
  ntilde: 'ñ', ograve: 'ò', oacute: 'ó', ocirc: 'ô', otilde: 'õ', ouml: 'ö', oslash: 'ø', ugrave: 'ù',
  uacute: 'ú', ucirc: 'û', uuml: 'ü', yacute: 'ý', yuml: 'ÿ', szlig: 'ß', Agrave: 'À', Aacute: 'Á',
  Acirc: 'Â', Atilde: 'Ã', Auml: 'Ä', Aring: 'Å', AElig: 'Æ', Ccedil: 'Ç', Egrave: 'È', Eacute: 'É',
  Ecirc: 'Ê', Euml: 'Ë', Igrave: 'Ì', Iacute: 'Í', Icirc: 'Î', Iuml: 'Ï', Ntilde: 'Ñ', Ograve: 'Ò',
  Oacute: 'Ó', Ocirc: 'Ô', Otilde: 'Õ', Ouml: 'Ö', Oslash: 'Ø', Ugrave: 'Ù', Uacute: 'Ú', Ucirc: 'Û',
  Uuml: 'Ü', Yacute: 'Ý',
};

export function decodeEntities(s: string): string {
  if (!s.includes('&')) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z][a-zA-Z0-9]*);?/g, (m, body: string) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : m;
    }
    return NAMED_ENTITIES[body] ?? m;
  });
}

/** Escape per testo XML/HTML; toglie i caratteri di controllo non ammessi in XML 1.0. */
export function escapeText(s: string): string {
  return s
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

export function escapeAttr(s: string): string {
  return escapeText(s).replace(/"/g, '&quot;');
}

function isNameStart(ch: string | undefined): boolean {
  return !!ch && /[A-Za-z_:]/.test(ch);
}

/** Analizza un documento; restituisce un elemento radice `#document`. */
export function parseMarkup(source: string, mode: 'xml' | 'html'): MarkupElement {
  const html = mode === 'html';
  const root: MarkupElement = { name: '#document', attrs: {}, children: [] };
  const stack: MarkupElement[] = [root];
  const top = () => stack[stack.length - 1]!;
  const len = source.length;
  let i = 0;

  const pushText = (text: string) => {
    if (!text) return;
    const children = top().children;
    const last = children[children.length - 1];
    if (typeof last === 'string') children[children.length - 1] = last + text;
    else children.push(text);
  };

  while (i < len) {
    const lt = source.indexOf('<', i);
    if (lt < 0) {
      pushText(decodeEntities(source.slice(i)));
      break;
    }
    if (lt > i) pushText(decodeEntities(source.slice(i, lt)));
    i = lt;

    if (source.startsWith('<!--', i)) {
      const end = source.indexOf('-->', i + 4);
      i = end < 0 ? len : end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', i)) {
      const end = source.indexOf(']]>', i + 9);
      pushText(source.slice(i + 9, end < 0 ? len : end));
      i = end < 0 ? len : end + 3;
      continue;
    }
    if (source[i + 1] === '!' || source[i + 1] === '?') {
      const end = source.indexOf('>', i);
      i = end < 0 ? len : end + 1;
      continue;
    }
    if (source[i + 1] === '/') {
      const end = source.indexOf('>', i);
      let name = source.slice(i + 2, end < 0 ? len : end).trim();
      i = end < 0 ? len : end + 1;
      if (html) name = name.toLowerCase();
      for (let k = stack.length - 1; k > 0; k--) {
        if (stack[k]!.name === name) {
          stack.length = k;
          break;
        }
      }
      continue;
    }
    if (!isNameStart(source[i + 1])) {
      pushText('<');
      i += 1;
      continue;
    }

    // Tag di apertura.
    let j = i + 1;
    while (j < len && !/[\s/>]/.test(source[j]!)) j++;
    let name = source.slice(i + 1, j);
    if (html) name = name.toLowerCase();
    const attrs: Record<string, string> = {};
    let selfClosing = false;
    while (j < len) {
      while (j < len && /\s/.test(source[j]!)) j++;
      const ch = source[j];
      if (ch === '>') {
        j++;
        break;
      }
      if (ch === '/') {
        j++;
        if (source[j] === '>') {
          selfClosing = true;
          j++;
          break;
        }
        continue;
      }
      let k = j;
      while (k < len && !/[\s/>=]/.test(source[k]!)) k++;
      let attrName = source.slice(j, k);
      if (html) attrName = attrName.toLowerCase();
      j = k;
      while (j < len && /\s/.test(source[j]!)) j++;
      let value = '';
      if (source[j] === '=') {
        j++;
        while (j < len && /\s/.test(source[j]!)) j++;
        const q = source[j];
        if (q === '"' || q === "'") {
          const end = source.indexOf(q, j + 1);
          value = source.slice(j + 1, end < 0 ? len : end);
          j = end < 0 ? len : end + 1;
        } else {
          let e = j;
          while (e < len && !/[\s>]/.test(source[e]!)) e++;
          value = source.slice(j, e);
          j = e;
        }
      }
      if (attrName && !(attrName in attrs)) attrs[attrName] = decodeEntities(value);
    }
    i = j;

    if (html) {
      const closes = IMPLICIT_CLOSE[name];
      if (closes) while (stack.length > 1 && closes.includes(top().name)) stack.pop();
    }
    const el: MarkupElement = { name, attrs, children: [] };
    top().children.push(el);
    if (selfClosing || (html && VOID.has(name))) continue;
    if (html && RAW_TEXT.has(name)) {
      const re = new RegExp(`</${name}\\s*>`, 'i');
      const m = re.exec(source.slice(i));
      const end = m ? i + m.index : len;
      const raw = source.slice(i, end);
      if (raw) el.children.push(name === 'title' || name === 'textarea' ? decodeEntities(raw) : raw);
      i = m ? end + m[0].length : len;
      continue;
    }
    stack.push(el);
  }
  return root;
}

export const parseXml = (source: string): MarkupElement => parseMarkup(source, 'xml');
export const parseHtml = (source: string): MarkupElement => parseMarkup(source, 'html');

// ---------------------------------------------------------------------------
// Navigazione

export function isElement(node: MarkupNode | undefined): node is MarkupElement {
  return typeof node === 'object' && node !== null;
}

export function elements(el: MarkupElement, name?: string): MarkupElement[] {
  const out: MarkupElement[] = [];
  for (const c of el.children) if (isElement(c) && (!name || c.name === name)) out.push(c);
  return out;
}

export function child(el: MarkupElement | undefined, name: string): MarkupElement | undefined {
  if (!el) return undefined;
  for (const c of el.children) if (isElement(c) && c.name === name) return c;
  return undefined;
}

/** Primo discendente (in profondità) con il nome dato. */
export function descendant(el: MarkupElement | undefined, name: string): MarkupElement | undefined {
  if (!el) return undefined;
  for (const c of el.children) {
    if (!isElement(c)) continue;
    if (c.name === name) return c;
    const d = descendant(c, name);
    if (d) return d;
  }
  return undefined;
}

export function textContent(node: MarkupNode): string {
  if (typeof node === 'string') return node;
  return node.children.map(textContent).join('');
}
