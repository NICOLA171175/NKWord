/**
 * Modulo di editing: registra i comandi edit.* (tranne annulla/ripeti),
 * format.*, paragraph.* e style.apply. Tutti i comandi sono puri:
 * stato → transazione.
 */

import type { Command, CommandArgs, CommandId, Editor, EditorModule, EditorState } from '../core';
import { find, replace } from './find';
import { applyCharChange, applyToggle, clearFormatting, growSize, isToggleActive, setLink, shrinkSize, TOGGLES, MAX_FONT_SIZE, MIN_FONT_SIZE } from './format';
import { moveCaret, selectAll, setSelection, type LineBounds } from './navigation';
import { applyStyle, setParaFormat, shiftIndent } from './paragraph';
import { deleteBackward, deleteForward, deleteSelection, insertFragment, insertInline, insertText, splitParagraph } from './text';
import { isCollapsed } from './selection';

type EditingCommand<K extends CommandId> = Command<CommandArgs[K]> & { readonly id: K };

function cmd<K extends CommandId>(id: K, def: Omit<EditingCommand<K>, 'id'>): EditingCommand<K> {
  return { id, ...def };
}

const hasSelection = ({ state }: { state: EditorState }) => !isCollapsed(state.selection);

/** Comandi del modulo. `lineBounds` permette Inizio/Fine di riga sulle righe visive. */
export function editingCommands(lineBounds?: LineBounds): Command<any>[] {
  const toggle = (id: CommandId, key: keyof typeof TOGGLES, label: string) =>
    cmd(id as 'format.bold', {
      label,
      isActive: ({ state }) => isToggleActive(state, TOGGLES[key]),
      run: ({ state }) => applyToggle(state, TOGGLES[key], label),
    });

  return [
    // Testo e selezione -------------------------------------------------
    cmd('edit.insertText', { label: 'Inserisci testo', run: ({ state }, a) => insertText(state, a.text) }),
    cmd('edit.deleteBackward', { label: 'Elimina indietro', run: ({ state }, a) => deleteBackward(state, a?.unit) }),
    cmd('edit.deleteForward', { label: 'Elimina avanti', run: ({ state }, a) => deleteForward(state, a?.unit) }),
    cmd('edit.deleteSelection', {
      label: 'Elimina',
      isEnabled: hasSelection,
      run: ({ state }) => deleteSelection(state),
    }),
    cmd('edit.splitParagraph', { label: 'Nuovo paragrafo', run: ({ state }) => splitParagraph(state) }),
    cmd('edit.insertLineBreak', {
      label: 'Interruzione di riga',
      run: ({ state }) => insertInline(state, { type: 'break', kind: 'line' }, 'Interruzione di riga'),
    }),
    cmd('edit.insertTab', {
      label: 'Tabulazione',
      run: ({ state }) => insertInline(state, { type: 'tab' }, 'Tabulazione'),
    }),
    cmd('edit.selectAll', { label: 'Seleziona tutto', run: ({ state }) => selectAll(state) }),
    cmd('edit.setSelection', { label: 'Selezione', run: ({ state }, a) => setSelection(state, a.selection) }),
    cmd('edit.moveCaret', {
      label: 'Sposta cursore',
      run: ({ state }, a) => moveCaret(state, a.direction, a.extend, lineBounds),
    }),
    cmd('edit.insertFragment', { label: 'Incolla', run: ({ state }, a) => insertFragment(state, a.fragment) }),
    cmd('edit.find', {
      label: 'Trova',
      isEnabled: (_ctx, a) => !!a?.query,
      run: ({ state }, a) => find(state, a.query, a, a.from),
    }),
    cmd('edit.replace', {
      label: 'Sostituisci',
      isEnabled: (_ctx, a) => !!a?.query,
      run: ({ state }, a) => replace(state, a.query, a.replacement, a),
    }),

    // Carattere ----------------------------------------------------------
    toggle('format.bold', 'bold', 'Grassetto'),
    toggle('format.italic', 'italic', 'Corsivo'),
    toggle('format.underline', 'underline', 'Sottolineato'),
    toggle('format.strike', 'strike', 'Barrato'),
    toggle('format.superscript', 'superscript', 'Apice'),
    toggle('format.subscript', 'subscript', 'Pedice'),
    cmd('format.fontFamily', {
      label: 'Tipo di carattere',
      isEnabled: (_ctx, a) => !!a?.fontFamily,
      run: ({ state }, a) => applyCharChange(state, () => ({ fontFamily: a.fontFamily }), 'Tipo di carattere'),
    }),
    cmd('format.fontSize', {
      label: 'Dimensione carattere',
      isEnabled: (_ctx, a) => !!a && a.fontSize >= MIN_FONT_SIZE && a.fontSize <= MAX_FONT_SIZE,
      run: ({ state }, a) => applyCharChange(state, () => ({ fontSize: a.fontSize }), 'Dimensione carattere'),
    }),
    cmd('format.growFont', {
      label: 'Aumenta dimensione carattere',
      run: ({ state }) =>
        applyCharChange(state, (f) => ({ fontSize: growSize(f.fontSize ?? 11) }), 'Aumenta dimensione carattere'),
    }),
    cmd('format.shrinkFont', {
      label: 'Riduci dimensione carattere',
      run: ({ state }) =>
        applyCharChange(state, (f) => ({ fontSize: shrinkSize(f.fontSize ?? 11) }), 'Riduci dimensione carattere'),
    }),
    cmd('format.color', {
      label: 'Colore carattere',
      run: ({ state }, a) => applyCharChange(state, () => ({ color: a.color }), 'Colore carattere'),
    }),
    cmd('format.highlight', {
      label: 'Colore evidenziazione testo',
      run: ({ state }, a) => applyCharChange(state, () => ({ highlight: a.color }), 'Evidenziazione'),
    }),
    cmd('format.setChar', {
      label: 'Carattere',
      run: ({ state }, a) => applyCharChange(state, () => a.format, 'Carattere'),
    }),
    cmd('format.clear', { label: 'Cancella formattazione', run: ({ state }) => clearFormatting(state) }),
    cmd('format.link', { label: 'Collegamento', run: ({ state }, a) => setLink(state, a.url) }),

    // Paragrafo e stili --------------------------------------------------
    cmd('paragraph.align', {
      label: 'Allineamento',
      run: ({ state }, a) => setParaFormat(state, { align: a.align }, 'Allineamento'),
    }),
    cmd('paragraph.indent', { label: 'Aumenta rientro', run: ({ state }) => shiftIndent(state, 1) }),
    cmd('paragraph.outdent', { label: 'Riduci rientro', run: ({ state }) => shiftIndent(state, -1) }),
    cmd('paragraph.lineSpacing', {
      label: 'Interlinea',
      run: ({ state }, a) => setParaFormat(state, { lineSpacing: a.lineSpacing }, 'Interlinea'),
    }),
    cmd('paragraph.spacing', {
      label: 'Spaziatura paragrafo',
      run: ({ state }, a) =>
        setParaFormat(
          state,
          {
            ...(a.before !== undefined && { spaceBefore: Math.max(0, a.before) }),
            ...(a.after !== undefined && { spaceAfter: Math.max(0, a.after) }),
          },
          'Spaziatura paragrafo',
        ),
    }),
    cmd('paragraph.setFormat', {
      label: 'Paragrafo',
      run: ({ state }, a) => setParaFormat(state, a.format, 'Paragrafo'),
    }),
    cmd('style.apply', {
      label: 'Stili',
      isEnabled: ({ state }, a) => !!a && state.doc.styles.styles[a.styleId]?.type !== 'table' && !!state.doc.styles.styles[a.styleId],
      run: ({ state }, a) => applyStyle(state, a.styleId),
    }),
  ];
}

/** Inizio/fine della riga visiva ricavati dal servizio di impaginazione, se presente. */
function layoutLineBounds(editor: Editor): LineBounds {
  return (doc, pos) => {
    const engine = editor.services.layout;
    if (!engine) return null;
    // Fra due righe (offset = fine dell'una = inizio dell'altra) vince la seconda.
    let atEnd: [number, number] | null = null;
    for (const page of engine.layout(doc).pages) {
      for (const line of page.lines) {
        if (line.paragraphId !== pos.paragraphId) continue;
        if (line.start <= pos.offset && pos.offset < line.end) return [line.start, line.end];
        if (pos.offset === line.end) atEnd = [line.start, line.end];
      }
    }
    return atEnd;
  };
}

export const editingModule: EditorModule = {
  id: 'editing',
  activate(editor) {
    const disposers = editingCommands(layoutLineBounds(editor)).map((c) => editor.commands.register(c));
    return () => disposers.forEach((d) => d());
  },
};
