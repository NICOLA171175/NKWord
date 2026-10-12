import { describe, expect, it } from 'vitest';
import { collectModules, isEditorModule } from '../../src/ui/modules';

describe('scoperta dei moduli', () => {
  it('riconosce oggetti EditorModule e fabbriche create…Module', () => {
    const a = { id: 'a', activate: () => {} };
    const b = { id: 'b', activate: () => {} };
    const found = collectModules([
      { a, helper: () => 1, costante: 3 },
      { createBModule: () => b, a, createConOpzioni: (_x: number) => ({ id: 'x', activate() {} }) },
      { ui: { id: 'ui', activate() {} } },
    ], ['ui']);
    expect(found.map((m) => m.id)).toEqual(['a', 'b']);
    expect(isEditorModule(null)).toBe(false);
  });
});
