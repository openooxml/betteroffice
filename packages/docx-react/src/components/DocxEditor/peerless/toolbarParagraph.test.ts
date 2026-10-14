import { expect, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

for (const [id, args, attrs] of [
  ['alignment', { value: 'center' }, { alignment: 'center' }],
  ['lineSpacing', { value: 360 }, { lineSpacing: 360, lineSpacingRule: 'auto' }],
  ['setLtr', null, { bidi: false }],
  ['setRtl', null, { bidi: true }],
] as const) {
  test(`worker ${id} submits paragraph attributes and matches evaluated state`, async () => {
    const { controller, format, flush, context, parity } = toolbarFixture();
    parity(id, args as never);
    expect(controller.commands.store.getState(id, args as never).enabled).toBe(true);
    flush.mockImplementation(() => new Promise<void>(() => {}));
    expect(await controller.commands.store.execute(id, args as never)).toMatchObject({ ok: true });
    expect(format).toHaveBeenCalledWith({ kind: 'setParagraphAttrs', payload: { attrs } }, { fence: true });
    expect(flush).not.toHaveBeenCalled();
    context({ alignment: 'center', paragraphProperties: { lineSpacing: 360, bidi: true } });
    parity(id, args as never);
  });
}
