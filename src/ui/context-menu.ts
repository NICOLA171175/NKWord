/**
 * Menu contestuale (clic destro) del documento. Le voci arrivano dalle
 * estensioni della vista (`ViewExtension.contextMenu`): se nessuna ne
 * fornisce, resta il menu del browser.
 */

import type { Editor, Position } from '../core';
import { viewExtensions, type ContextMenuItem } from './extensions';

export function contextMenuItems(editor: Editor, pos: Position): ContextMenuItem[] {
  const out: ContextMenuItem[] = [];
  for (const ext of viewExtensions(editor)) {
    const items = ext.contextMenu?.(pos, editor.state) ?? [];
    items.forEach((item, i) => out.push(i === 0 && out.length > 0 ? { ...item, separator: true } : item));
  }
  return out;
}

export class ContextMenu {
  readonly element: HTMLElement;
  private readonly cleanup: (() => void)[] = [];

  constructor(
    items: readonly ContextMenuItem[],
    x: number,
    y: number,
    private readonly onClose: () => void = () => {},
  ) {
    this.element = document.createElement('div');
    this.element.className = 'context-menu';
    this.element.setAttribute('role', 'menu');
    for (const item of items) {
      if (item.separator) {
        const sep = document.createElement('div');
        sep.className = 'context-menu-separator';
        sep.setAttribute('role', 'separator');
        this.element.append(sep);
      }
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'context-menu-item';
      b.setAttribute('role', 'menuitem');
      b.textContent = item.label;
      b.disabled = !!item.disabled || !item.run;
      if (item.strong) b.classList.add('context-menu-strong');
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => {
        this.close();
        item.run?.();
      });
      this.element.append(b);
    }
    this.element.style.left = `${x}px`;
    this.element.style.top = `${y}px`;
    document.body.append(this.element);

    // Resta dentro la finestra.
    const r = this.element.getBoundingClientRect();
    const vw = globalThis.innerWidth ?? 0;
    const vh = globalThis.innerHeight ?? 0;
    if (vw && r.right > vw) this.element.style.left = `${Math.max(0, x - r.width)}px`;
    if (vh && r.bottom > vh) this.element.style.top = `${Math.max(0, y - r.height)}px`;

    const outside = (e: Event) => {
      if (!this.element.contains(e.target as Node)) this.close();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        this.close();
      }
    };
    document.addEventListener('mousedown', outside, true);
    document.addEventListener('keydown', onKey, true);
    globalThis.addEventListener?.('blur', this.closeNow);
    this.cleanup.push(
      () => document.removeEventListener('mousedown', outside, true),
      () => document.removeEventListener('keydown', onKey, true),
      () => globalThis.removeEventListener?.('blur', this.closeNow),
    );
    this.element.querySelector<HTMLButtonElement>('.context-menu-item:not(:disabled)')?.focus();
  }

  get isOpen(): boolean {
    return this.element.isConnected;
  }

  private readonly closeNow = () => this.close();

  close(): void {
    if (!this.isOpen) return;
    for (const c of this.cleanup.splice(0)) c();
    this.element.remove();
    this.onClose();
  }
}
