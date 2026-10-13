/**
 * Finestra "Modifica dati": una piccola tabella come il foglio di Excel che
 * Word apre per i grafici. Righe = categorie, colonne = serie.
 */

import type { ChartData } from '../core';
import { CHART_LABELS } from './model';

/** Numero scritto all'italiana ("3,5") o all'inglese ("3.5"); vuoto o non valido = 0. */
export function parseNumber(text: string): number {
  const clean = text.trim().replace(/\s/g, '');
  if (!clean) return 0;
  const normalized = /,\d*$/.test(clean) ? clean.replace(/\./g, '').replace(',', '.') : clean.replace(/,/g, '');
  const n = Number(normalized);
  return Number.isFinite(n) ? n : 0;
}

/** Contenuto della tabella: prima riga = nomi delle serie, prima colonna = categorie. */
export interface ChartGrid {
  readonly type: ChartData['type'];
  readonly title: string;
  readonly cells: readonly (readonly string[])[];
}

export function chartToGrid(chart: ChartData): ChartGrid {
  const header = ['', ...chart.series.map((s) => s.name)];
  const rows = chart.categories.map((c, i) => [c, ...chart.series.map((s) => String(s.values[i] ?? 0).replace('.', ','))]);
  return { type: chart.type, title: chart.title ?? '', cells: [header, ...rows] };
}

export function gridToChart(grid: ChartGrid, previous?: ChartData): ChartData {
  const [header = [''], ...rows] = grid.cells;
  const seriesCount = Math.max(1, header.length - 1);
  const series = Array.from({ length: seriesCount }, (_, si) => {
    const old = previous?.series[si];
    return {
      name: header[si + 1]?.trim() || `Serie ${si + 1}`,
      values: rows.map((r) => parseNumber(r[si + 1] ?? '')),
      ...(old?.color ? { color: old.color } : {}),
    };
  });
  const title = grid.title.trim();
  return {
    type: grid.type,
    ...(title ? { title } : {}),
    categories: rows.map((r, i) => r[0]?.trim() || `Categoria ${i + 1}`),
    series,
  };
}

/** Apre la finestra; `onSave` riceve i dati nuovi. Restituisce la funzione per chiuderla. */
export function openChartEditor(chart: ChartData, onSave: (chart: ChartData) => void, onClose: () => void = () => {}): () => void {
  let grid = chartToGrid(chart);
  const backdrop = document.createElement('div');
  backdrop.className = 'drawing-dialog-backdrop';
  const dialog = document.createElement('div');
  dialog.className = 'drawing-dialog chart-editor';
  dialog.setAttribute('role', 'dialog');
  dialog.setAttribute('aria-label', 'Modifica dati del grafico');
  backdrop.append(dialog);

  const title = document.createElement('h2');
  title.textContent = 'Modifica dati del grafico';

  const typeSelect = document.createElement('select');
  typeSelect.className = 'chart-editor-type';
  for (const [value, label] of Object.entries(CHART_LABELS)) {
    const o = document.createElement('option');
    o.value = value;
    o.textContent = label;
    typeSelect.append(o);
  }
  typeSelect.value = grid.type;
  const titleInput = document.createElement('input');
  titleInput.className = 'chart-editor-title';
  titleInput.value = grid.title;
  const field = (label: string, control: HTMLElement) => {
    const l = document.createElement('label');
    l.className = 'drawing-dialog-field';
    l.append(label, control);
    return l;
  };

  const table = document.createElement('table');
  table.className = 'chart-editor-grid';
  /** Legge la tabella mostrata (l'utente può aver modificato le celle). */
  const read = (): ChartGrid => ({
    type: typeSelect.value as ChartData['type'],
    title: titleInput.value,
    cells: Array.from(table.rows).map((tr) => Array.from(tr.querySelectorAll('input')).map((i) => i.value)),
  });
  const renderGrid = () => {
    table.replaceChildren();
    grid.cells.forEach((row, r) => {
      const tr = table.insertRow();
      row.forEach((value, c) => {
        const td = tr.insertCell();
        const input = document.createElement('input');
        input.value = value;
        if (r === 0 && c === 0) {
          input.disabled = true;
          input.className = 'chart-editor-corner';
        } else if (r === 0) input.className = 'chart-editor-series';
        else if (c === 0) input.className = 'chart-editor-category';
        else {
          input.className = 'chart-editor-value';
          input.inputMode = 'decimal';
        }
        td.append(input);
      });
    });
  };
  renderGrid();

  const button = (label: string, cls: string, run: () => void) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = cls;
    b.textContent = label;
    b.addEventListener('click', run);
    return b;
  };
  const edit = (change: (cells: string[][]) => void) => {
    const current = read();
    const cells = current.cells.map((r) => [...r]);
    change(cells);
    grid = { ...current, cells };
    renderGrid();
  };
  const tools = document.createElement('div');
  tools.className = 'drawing-dialog-tools';
  tools.append(
    button('+ Categoria', 'chart-add-row', () => edit((c) => c.push([`Categoria ${c.length}`, ...c[0]!.slice(1).map(() => '0')]))),
    button('− Categoria', 'chart-remove-row', () => edit((c) => c.length > 2 && c.pop())),
    button('+ Serie', 'chart-add-series', () => edit((c) => c.forEach((r, i) => r.push(i ? '0' : `Serie ${c[0]!.length}`)))),
    button('− Serie', 'chart-remove-series', () => edit((c) => c[0]!.length > 2 && c.forEach((r) => r.pop()))),
  );

  const close = () => {
    document.removeEventListener('keydown', onKey, true);
    backdrop.remove();
    onClose();
  };
  const save = () => {
    const result = gridToChart(read(), chart);
    close();
    onSave(result);
  };
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    } else if (e.key === 'Enter' && !(e.target instanceof HTMLButtonElement)) {
      e.preventDefault();
      save();
    }
  };
  document.addEventListener('keydown', onKey, true);

  const footer = document.createElement('div');
  footer.className = 'drawing-dialog-footer';
  footer.append(button('Annulla', 'drawing-dialog-cancel', close), button('OK', 'drawing-dialog-ok', save));

  dialog.append(title, field('Tipo di grafico', typeSelect), field('Titolo', titleInput), table, tools, footer);
  document.body.append(backdrop);
  titleInput.focus();
  return close;
}
