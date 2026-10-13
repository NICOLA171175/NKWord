/** Costanti e corrispondenze condivise fra lettura e scrittura DOCX. */

import type { Alignment, NumberFormat, TabAlignment, UnderlineStyle } from '../../../core';

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  coreProps: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
} as const;

export const CONTENT_TYPES = {
  document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
} as const;

export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export const JC_TO_DOCX: Record<Alignment, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' };
export const JC_FROM_DOCX: Record<string, Alignment> = {
  left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify',
};

export const UNDERLINE_TO_DOCX: Record<UnderlineStyle, string> = {
  single: 'single', double: 'double', dotted: 'dotted', dashed: 'dash', wave: 'wave',
};
export function underlineFromDocx(v: string): UnderlineStyle | null {
  if (v === 'none') return null;
  if (v === 'double') return 'double';
  if (/^dotted/.test(v)) return 'dotted';
  if (/^dash|^dotDash|^dashLong/.test(v)) return 'dashed';
  if (/^wav/.test(v)) return 'wave';
  return 'single';
}

export const NUMFMT_TO_DOCX: Record<NumberFormat, string> = {
  bullet: 'bullet', decimal: 'decimal', lowerLetter: 'lowerLetter', upperLetter: 'upperLetter',
  lowerRoman: 'lowerRoman', upperRoman: 'upperRoman', none: 'none',
};
export function numFmtFromDocx(v: string | undefined): NumberFormat {
  switch (v) {
    case 'bullet':
    case 'lowerLetter':
    case 'upperLetter':
    case 'lowerRoman':
    case 'upperRoman':
    case 'none':
      return v;
    default:
      return 'decimal';
  }
}

export const TAB_FROM_DOCX: Record<string, TabAlignment> = {
  left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', decimal: 'decimal',
};

/** Nomi interni di Word per gli stili predefiniti (i loro id variano con la lingua). */
export const BUILTIN_STYLE_NAMES: Record<string, string> = {
  Normal: 'Normal',
  Heading1: 'heading 1',
  Heading2: 'heading 2',
  Heading3: 'heading 3',
  Heading4: 'heading 4',
  Heading5: 'heading 5',
  Heading6: 'heading 6',
  Title: 'Title',
  Subtitle: 'Subtitle',
  Hyperlink: 'Hyperlink',
  FootnoteText: 'footnote text',
};
