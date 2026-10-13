// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { Editor, firstParagraph, paragraphText } from '../../src/core';
import { createIoModule } from '../../src/io';
import { MAX_RECENT, MemoryRecentStore, createRecentModule, whenUsed, type DesktopBridge, type RecentHost } from '../../src/recent';
import type { FileHandle, OpenedFile } from '../../src/ui/shell';
import { createApp } from '../../src/ui/shell';

afterEach(() => document.body.replaceChildren());

const tick = () => new Promise((r) => setTimeout(r, 0));
const txt = (name: string, text = 'ciao') => new File([text], name);

function setup(options: { desktop?: DesktopBridge | null; store?: MemoryRecentStore } = {}) {
  const editor = new Editor();
  editor.use(createIoModule({ save: () => {} }));
  let t = Date.parse('2026-10-13T08:00:00Z');
  const recent = createRecentModule({ store: options.store ?? new MemoryRecentStore(), desktop: options.desktop ?? null, now: () => new Date((t += 1000)) });
  editor.use(recent);
  return { editor, recent };
}

/** Shell finta: apre i file con `file.open` come quella vera. */
function host(editor: Editor, ribbon?: HTMLElement) {
  const listeners = new Set<(o: OpenedFile) => void>();
  const h: RecentHost & { picked: number } = {
    editor,
    picked: 0,
    ...(ribbon ? { ribbon: { element: ribbon } } : {}),
    async openFile(file, from = {}) {
      const ok = await editor.executeAsync('file.open', { data: await file.text(), formatId: 'txt' });
      if (ok) for (const l of listeners) l({ file, ...from });
      return ok;
    },
    pickFile() {
      h.picked++;
    },
    onFileOpened(l) {
      listeners.add(l);
      return () => listeners.delete(l);
    },
  };
  return h;
}

function handle(name: string, text: string, permission: PermissionState = 'granted'): FileHandle & { asked: number } {
  const h = {
    kind: 'file' as const,
    name,
    asked: 0,
    getFile: async () => txt(name, text),
    isSameEntry: async (o: FileHandle) => o === h,
    queryPermission: async () => permission,
    requestPermission: async () => {
      h.asked++;
      return permission === 'denied' ? ('denied' as const) : ('granted' as const);
    },
  };
  return h;
}

describe('documenti recenti', () => {
  it('ricorda i file aperti, il più recente in cima, senza doppioni', async () => {
    const store = new MemoryRecentStore();
    const { recent } = setup({ store });
    await recent.ready;
    await recent.record({ file: txt('a.txt') });
    await recent.record({ file: txt('b.txt') });
    await recent.record({ file: txt('a.txt') });
    expect(recent.entries.map((e) => e.name)).toEqual(['a.txt', 'b.txt']);
    expect((await store.list()).length).toBe(2);

    // Riavviando l'app l'elenco si rilegge dall'archivio, nello stesso ordine.
    const again = setup({ store }).recent;
    await again.ready;
    expect(again.entries.map((e) => e.name)).toEqual(['a.txt', 'b.txt']);
  });

  it(`tiene al massimo ${MAX_RECENT} documenti`, async () => {
    const store = new MemoryRecentStore();
    const { recent } = setup({ store });
    for (let i = 0; i < MAX_RECENT + 3; i++) await recent.record({ file: txt(`f${i}.txt`) });
    expect(recent.entries).toHaveLength(MAX_RECENT);
    expect(recent.entries[0]!.name).toBe(`f${MAX_RECENT + 2}.txt`);
    expect(await store.list()).toHaveLength(MAX_RECENT);
  });

  it('nel browser riapre con la maniglia del file, chiedendo il permesso', async () => {
    const { editor, recent } = setup();
    recent.attach(host(editor));
    const h = handle('lettera.txt', 'Cara Anna', 'prompt');
    await recent.record({ file: txt('lettera.txt'), handle: h });
    await recent.record({ file: txt('lettera.txt'), handle: h });
    expect(recent.entries).toHaveLength(1);
    expect(await recent.open(recent.entries[0]!.id)).toBe('opened');
    expect(h.asked).toBe(1);
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('Cara Anna');

    const no = handle('segreto.txt', 'x', 'denied');
    await recent.record({ file: txt('segreto.txt'), handle: no });
    expect(await recent.open(recent.entries[0]!.id)).toBe('denied');
  });

  it('senza maniglia mostra la finestra per scegliere il file', async () => {
    const { editor, recent } = setup();
    const h = host(editor);
    recent.attach(h);
    await recent.record({ file: txt('vecchio.txt') });
    expect(await recent.open(recent.entries[0]!.id)).toBe('picked');
    expect(h.picked).toBe(1);
  });

  it('salvando nel browser aggiunge il nome del file', async () => {
    const { editor, recent } = setup();
    await editor.executeAsync('file.export', { formatId: 'docx' });
    await tick();
    expect(recent.entries.map((e) => e.name)).toEqual(['Documento.docx']);
  });

  it('nel programma desktop usa i percorsi sul disco', async () => {
    let saved: ((p: string) => void) | undefined;
    const files: Record<string, string> = { 'C:\\Documenti\\Relazione.txt': 'Relazione annuale' };
    const desktop: DesktopBridge = {
      pathForFile: (f) => (f.name === 'scelto.txt' ? 'C:\\Documenti\\scelto.txt' : ''),
      readFile: async (p) => {
        if (!(p in files)) throw new Error('ENOENT');
        return new TextEncoder().encode(files[p]);
      },
      onFileSaved: (l) => ((saved = l), () => (saved = undefined)),
    };
    const { editor, recent } = setup({ desktop });
    recent.attach(host(editor));
    await recent.record({ file: txt('scelto.txt') });
    expect(recent.entries[0]).toMatchObject({ name: 'scelto.txt', path: 'C:\\Documenti\\scelto.txt' });

    saved!('C:\\Documenti\\Relazione.txt');
    saved!('C:\\Documenti\\foto.png'); // non è un documento: ignorato
    await tick();
    expect(recent.entries.map((e) => e.name)).toEqual(['Relazione.txt', 'scelto.txt']);

    // Riaprendolo torna in cima senza doppioni (anche con maiuscole diverse).
    expect(await recent.open(recent.entries[0]!.id)).toBe('opened');
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('Relazione annuale');
    await recent.record({ file: txt('x'), path: 'c:\\documenti\\SCELTO.txt' });
    expect(recent.entries).toHaveLength(2);

    expect(await recent.open(recent.entries.find((e) => e.name === 'Relazione.txt')!.id)).toBe('opened');
    delete files['C:\\Documenti\\Relazione.txt'];
    expect(await recent.open(recent.entries[0]!.id)).toBe('missing');
  });

  it('comandi per togliere un documento o svuotare l\'elenco', async () => {
    const { editor, recent } = setup();
    await recent.record({ file: txt('a.txt') });
    await recent.record({ file: txt('b.txt') });
    expect(editor.canExecute('recent.clear')).toBe(true);
    await editor.executeAsync('recent.remove', { id: recent.entries[0]!.id });
    expect(recent.entries.map((e) => e.name)).toEqual(['a.txt']);
    await editor.executeAsync('recent.clear');
    expect(recent.entries).toEqual([]);
    expect(editor.canExecute('recent.clear')).toBe(false);
  });

  it('scrive le date in parole', () => {
    const today = new Date(2026, 9, 13, 18, 0);
    expect(whenUsed(new Date(2026, 9, 13, 9, 5).toISOString(), today)).toBe('Oggi alle 09:05');
    expect(whenUsed(new Date(2026, 9, 12, 9, 5).toISOString(), today)).toMatch(/^Ieri/);
    expect(whenUsed(new Date(2026, 8, 1).toISOString(), today)).toBe('1 settembre');
  });
});

describe('elenco nella scheda File', () => {
  it('si apre dal pulsante, riapre un documento e toglie quelli spariti', async () => {
    const { editor, recent } = setup();
    const ribbon = document.createElement('div');
    const button = document.createElement('button');
    button.dataset.command = 'recent.show';
    ribbon.append(button);
    document.body.append(ribbon);
    const h = host(editor, ribbon);
    recent.mount(h);
    await recent.ready;

    editor.execute('recent.show');
    const panel = document.querySelector<HTMLElement>('.recent-panel')!;
    expect(panel.hidden).toBe(false);
    expect(panel.textContent).toContain('Nessun documento recente');

    // Un file aperto dalla shell entra nell'elenco.
    await h.openFile(txt('nota.txt', 'Promemoria'), { handle: handle('nota.txt', 'Promemoria') });
    await tick();
    expect(Array.from(panel.querySelectorAll('.recent-name')).map((e) => e.textContent)).toEqual(['nota.txt']);
    expect(panel.querySelector('.recent-where')!.textContent).toContain('Sul computer');

    await editor.executeAsync('file.new');
    panel.querySelector<HTMLButtonElement>('.recent-open')!.click();
    await tick();
    expect(panel.hidden).toBe(true);
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('Promemoria');

    // File sparito: si chiede se toglierlo dall'elenco.
    const gone = handle('perso.txt', '');
    gone.getFile = async () => Promise.reject(Object.assign(new Error('no'), { name: 'NotFoundError' }));
    await recent.record({ file: txt('perso.txt'), handle: gone });
    const confirm = vi.fn(() => true);
    window.confirm = confirm;
    editor.execute('recent.show');
    panel.querySelector<HTMLButtonElement>('.recent-open')!.click();
    await tick();
    await tick();
    expect(confirm).toHaveBeenCalledOnce();
    expect(recent.entries.map((e) => e.name)).toEqual(['nota.txt']);

    editor.execute('recent.show');
    panel.querySelector<HTMLButtonElement>('.recent-remove')!.click();
    await tick();
    expect(panel.textContent).toContain('Nessun documento recente');

    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(panel.hidden).toBe(true);
  });

  it('la shell avvisa dei file aperti e usa la finestra con maniglia se c\'è', async () => {
    const root = document.createElement('div');
    document.body.append(root);
    const editor = new Editor();
    editor.use(createIoModule({ save: () => {} }));
    const h = handle('scelto.txt', 'Dal disco');
    (window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async () => [h];
    try {
      const app = createApp(root, editor);
      const opened: OpenedFile[] = [];
      app.onFileOpened((o) => opened.push(o));
      app.pickFile();
      await tick();
      await tick();
      expect(paragraphText(firstParagraph(editor.state.doc))).toBe('Dal disco');
      expect(opened).toHaveLength(1);
      expect(opened[0]!.handle).toBe(h);
      app.destroy();
    } finally {
      delete (window as unknown as { showOpenFilePicker?: unknown }).showOpenFilePicker;
    }
  });
});
