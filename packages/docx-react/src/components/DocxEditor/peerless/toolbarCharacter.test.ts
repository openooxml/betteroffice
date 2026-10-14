import { expect, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

for (const [id, intent] of [
  ['strikethrough', { kind: 'toggleStrike', payload: {} }],
  ['superscript', { kind: 'toggleMark', payload: { mark: { type: 'superscript' } } }],
  ['subscript', { kind: 'toggleMark', payload: { mark: { type: 'subscript' } } }],
  ['clearFormatting', { kind: 'clearFormatting', payload: {} }],
] as const) {
  test(`worker ${id} uses causal formatting and matches evaluated state`, async () => {
    const { controller, format, flush, context, parity } = toolbarFixture();
    parity(id);
    expect(controller.commands.store.getState(id).enabled).toBe(true);
    flush.mockImplementation(() => new Promise<void>(() => {}));
    expect(await controller.commands.store.execute(id, null)).toEqual({ ok: true, status: 'executed' });
    expect(format).toHaveBeenCalledWith(intent, { fence: true });
    expect(flush).not.toHaveBeenCalled();
    context({ strike: 'mixed', superscript: true, subscript: false });
    parity(id);
    controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, readOnly: true });
    parity(id);
    expect(await controller.commands.store.execute(id, null)).toMatchObject({ ok: false, failure: { code: 'read-only' } });
    expect(format).toHaveBeenCalledTimes(1);
  });
}
