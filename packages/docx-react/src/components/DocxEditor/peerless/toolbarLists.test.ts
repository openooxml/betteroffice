import { expect, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

for (const [id, intent] of [
  ['bulletList', { kind: 'toggleList', payload: { kind: 'bullet' } }],
  ['numberedList', { kind: 'toggleList', payload: { kind: 'numbered' } }],
  ['indent', { kind: 'changeIndent', payload: { direction: 'increase' } }],
  ['outdent', { kind: 'changeIndent', payload: { direction: 'decrease' } }],
] as const) {
  test(`worker ${id} uses ordered list formatting and matches evaluated state`, async () => {
    const { controller, format, flush, context, parity } = toolbarFixture();
    parity(id);
    expect(controller.commands.store.getState(id).enabled).toBe(true);
    flush.mockImplementation(() => new Promise<void>(() => {}));
    expect(await controller.commands.store.execute(id, null)).toMatchObject({ ok: true });
    expect(format).toHaveBeenCalledWith(intent, { fence: true });
    expect(flush).not.toHaveBeenCalled();
    context({ paragraphProperties: { numPr: { numId: 1, ilvl: 0 } } });
    parity(id);
    context({ paragraphProperties: {} });
    parity(id);
  });
}
