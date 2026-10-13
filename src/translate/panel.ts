/**
 * Riquadro "Traduttore" a destra del documento, come in Word: le schede
 * Selezione e Documento, le lingue "Da" e "A", il testo tradotto con il
 * pulsante Inserisci, oppure il pulsante Traduci per l'intero documento.
 */

import { AUTO_DETECT, LANGUAGES, languageName } from './languages';
import type { TranslateModule, TranslateMode } from './module';
import './styles.css';

export interface TranslatePanelHost {
  readonly view: { readonly element: HTMLElement };
  /** Riporta il focus al documento. */
  readonly input?: { focus(): void };
}

function languageSelect(d: Document, label: string, withAuto: boolean): HTMLSelectElement {
  const select = d.createElement('select');
  select.className = 'translate-select';
  select.setAttribute('aria-label', label);
  const options = withAuto ? [{ code: AUTO_DETECT, name: languageName(AUTO_DETECT) }, ...LANGUAGES] : LANGUAGES;
  for (const lang of options) {
    const option = d.createElement('option');
    option.value = lang.code;
    option.textContent = lang.name;
    select.append(option);
  }
  return select;
}

function field(d: Document, label: string, control: HTMLElement): HTMLLabelElement {
  const wrap = d.createElement('label');
  wrap.className = 'translate-field';
  const caption = d.createElement('span');
  caption.textContent = label;
  wrap.append(caption, control);
  return wrap;
}

/**
 * Affianca il riquadro alla vista del documento, in una riga
 * `.translate-workspace` messa al posto della vista.
 */
export function mountTranslatePanel(host: TranslatePanelHost, module: TranslateModule, d: Document = document): () => void {
  const view = host.view.element;
  const row = d.createElement('div');
  row.className = 'translate-workspace';
  view.replaceWith(row);

  const pane = d.createElement('aside');
  pane.className = 'translate-pane';
  pane.setAttribute('aria-label', 'Traduttore');
  row.append(view, pane);

  const header = d.createElement('header');
  header.className = 'translate-header';
  const title = d.createElement('span');
  title.textContent = 'Traduttore';
  const close = d.createElement('button');
  close.type = 'button';
  close.className = 'translate-close';
  close.textContent = '✕';
  close.title = 'Chiudi';
  close.setAttribute('aria-label', 'Chiudi il Traduttore');
  header.append(title, close);

  const tabs = d.createElement('div');
  tabs.className = 'translate-tabs';
  tabs.setAttribute('role', 'tablist');
  const tab = (mode: TranslateMode, label: string) => {
    const b = d.createElement('button');
    b.type = 'button';
    b.className = 'translate-tab';
    b.textContent = label;
    b.setAttribute('role', 'tab');
    b.dataset.mode = mode;
    b.addEventListener('click', () => module.show(mode));
    return b;
  };
  const selectionTab = tab('selection', 'Selezione');
  const documentTab = tab('document', 'Documento');
  tabs.append(selectionTab, documentTab);

  // Scheda Selezione.
  const selectionBody = d.createElement('div');
  selectionBody.className = 'translate-body';
  const fromSelect = languageSelect(d, 'Da', true);
  const sourceBox = d.createElement('div');
  sourceBox.className = 'translate-text translate-source';
  const toSelect = languageSelect(d, 'A', false);
  const resultBox = d.createElement('div');
  resultBox.className = 'translate-text translate-result';
  resultBox.setAttribute('aria-live', 'polite');
  const insert = d.createElement('button');
  insert.type = 'button';
  insert.className = 'translate-primary';
  insert.textContent = 'Inserisci';
  insert.title = 'Sostituisce il testo selezionato con la traduzione';
  selectionBody.append(field(d, 'Da', fromSelect), sourceBox, field(d, 'A', toSelect), resultBox, insert);

  // Scheda Documento.
  const documentBody = d.createElement('div');
  documentBody.className = 'translate-body';
  const docFromSelect = languageSelect(d, 'Da', true);
  const docToSelect = languageSelect(d, 'Traduci in', false);
  const explain = d.createElement('p');
  explain.className = 'translate-note';
  explain.textContent =
    'Tutto il testo del documento (anche tabelle, intestazioni, piè di pagina, note e caselle di testo) viene sostituito dalla traduzione. Con Annulla (Ctrl+Z) torni all\'originale.';
  const translateDoc = d.createElement('button');
  translateDoc.type = 'button';
  translateDoc.className = 'translate-primary';
  translateDoc.textContent = 'Traduci';
  documentBody.append(field(d, 'Da', docFromSelect), field(d, 'Traduci in', docToSelect), explain, translateDoc);

  const status = d.createElement('p');
  status.className = 'translate-status';
  status.setAttribute('role', 'status');
  const footer = d.createElement('footer');
  footer.className = 'translate-footer';
  footer.textContent = 'Traduzione di Google Traduttore. Serve la connessione a Internet.';

  pane.append(header, tabs, selectionBody, documentBody, status, footer);

  const keepFocus = (b: HTMLElement) => b.addEventListener('mousedown', (e) => e.preventDefault());
  keepFocus(insert);
  close.addEventListener('click', () => {
    module.close();
    host.input?.focus();
  });
  insert.addEventListener('click', () => {
    module.insert();
    host.input?.focus();
  });
  translateDoc.addEventListener('click', () => void module.translateDocument());
  fromSelect.addEventListener('change', () => module.setLanguages({ from: fromSelect.value }));
  toSelect.addEventListener('change', () => module.setLanguages({ to: toSelect.value }));
  docFromSelect.addEventListener('change', () => module.setLanguages({ from: docFromSelect.value }));
  docToSelect.addEventListener('change', () => module.setLanguages({ to: docToSelect.value }));

  const render = () => {
    pane.hidden = !module.open;
    const selection = module.mode === 'selection';
    selectionTab.classList.toggle('active', selection);
    selectionTab.setAttribute('aria-selected', String(selection));
    documentTab.classList.toggle('active', !selection);
    documentTab.setAttribute('aria-selected', String(!selection));
    selectionBody.hidden = !selection;
    documentBody.hidden = selection;
    fromSelect.value = docFromSelect.value = module.from;
    toSelect.value = docToSelect.value = module.to;

    const busy = module.status === 'busy';
    sourceBox.textContent = module.source;
    sourceBox.hidden = !module.source;
    resultBox.textContent = busy && selection ? 'Traduzione in corso…' : module.result;
    resultBox.hidden = !module.result && !(busy && selection);
    resultBox.classList.toggle('translate-pending', busy);
    insert.disabled = module.status !== 'done' || !module.result;
    translateDoc.disabled = busy;
    translateDoc.textContent = busy && !selection ? 'Traduzione in corso…' : 'Traduci';

    let text = module.message ?? '';
    if (!text && module.status === 'done' && module.from === AUTO_DETECT && module.detected) {
      text = `Lingua rilevata: ${languageName(module.detected)}.`;
    }
    status.textContent = text;
    status.hidden = !text;
    status.classList.toggle('translate-error', module.status === 'error');
  };

  const off = module.onChange(render);
  render();
  return () => {
    off();
    row.replaceWith(view);
  };
}
