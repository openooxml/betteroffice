import { expect, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

for (const [id, args, attrs] of [
  ['alignment', { value: 'center' }, { alignment: 'center' }],
  ['lineSpacing', { value: 360 }, { lineSpacing: 360, lineSpacingRule: 'auto' }],
  ['paragraphSpacing', { before: 240 }, { spaceBefore: 240 }],
  ['paragraphSpacing', { after: 0 }, { spaceAfter: 0 }],
  ['paragraphSpacing', { before: 240, after: 240 }, { spaceBefore: 240, spaceAfter: 240 }],
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

for (const mode of ['editing', 'suggesting', 'viewing'] as const) {
  test(`worker paragraph spacing follows line spacing admission in ${mode}`, async () => {
    const { controller, format } = toolbarFixture();
    controller.update({ mode, zoom: 1, focus() {}, read: (async () => ({})) as never });
    expect(controller.commands.store.getState('paragraphSpacing').enabled).toBe(mode !== 'viewing');
    expect(controller.commands.store.getState('paragraphSpacingDialog').enabled).toBe(mode !== 'viewing');
    await controller.commands.store.execute('paragraphSpacing', { before: 240 });
    expect(format).toHaveBeenCalledTimes(mode === 'viewing' ? 0 : 1);
  });
}

test('worker paragraph spacing refuses read-only edits', async () => {
  const { controller, format } = toolbarFixture();
  controller.update({ readOnly: true, zoom: 1, focus() {}, read: (async () => ({})) as never });
  expect(controller.commands.store.getState('paragraphSpacing').enabled).toBe(false);
  await controller.commands.store.execute('paragraphSpacing', { after: 240 });
  expect(format).not.toHaveBeenCalled();
});
