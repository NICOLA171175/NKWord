/**
 * Appunti: un intervallo del documento diventa HTML (per Word, browser e
 * altre app) più testo semplice; in incolla l'HTML ha la precedenza sul
 * testo. Il frammento restituito da `parse` va inserito con
 * `edit.insertFragment` (modulo editing).
 */

import type { ClipboardService, DocumentModel, Position } from '../core';
import { sliceDocument } from './fragment';
import { documentToHtml, htmlToDocument } from './formats/html';
import { documentToText, textToDocument } from './formats/text';

export class DefaultClipboardService implements ClipboardService {
  serialize(doc: DocumentModel, from: Position, to: Position): { html: string; text: string } {
    const fragment = sliceDocument(doc, from, to);
    return { html: documentToHtml(fragment, { fragment: true }), text: documentToText(fragment) };
  }

  parse(data: { html?: string; text?: string }): DocumentModel {
    // Il documento intero, non solo fra StartFragment/EndFragment: i browser
    // mettono fuori dai marcatori gli antenati (<ul>, <table>) del frammento.
    if (data.html && /<[a-z!]/i.test(data.html)) return htmlToDocument(data.html);
    return textToDocument(data.text ?? data.html ?? '');
  }
}
