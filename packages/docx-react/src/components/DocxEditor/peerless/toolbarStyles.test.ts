import { expect, mock, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';

test('worker styles use versioned metadata with hybrid labels and preview units', async () => {
  const { controller, read, format, parity } = toolbarFixture();
  const readAt = mock(async () => ({ status: 'ok', version: 'native', value: { styles: [
    { styleId: 'Heading1', name: 'heading 1', fontSize: 40, bold: true },
    { styleId: 'Custom', name: 'Custom style', italic: true, color: '123456' },
  ], fontFamilies: ['Ignored face'] } }));
  controller.update({ read: read as never, readAt: readAt as never, focus() {}, zoom: 1 });
  const result = await controller.query({ kind: 'toolbarMetadata' });
  controller.toolbarMetadata(result);
  expect(readAt).toHaveBeenCalledWith({ kind: 'toolbarMetadata' }, 'native');
  expect(controller.environment().styles).toEqual([
    { args: { styleId: 'Heading1' }, label: 'Heading 1', preview: { fontSize: 20, bold: true } },
    { args: { styleId: 'Custom' }, label: 'Custom style', preview: { italic: true, color: '123456' } },
  ]);
  expect(controller.environment().fonts.some(option => option.label === 'Ignored face')).toBe(false);
  parity('paragraphStyle', { styleId: 'Custom' });
  expect(await controller.commands.store.execute('paragraphStyle', { styleId: 'Custom' })).toMatchObject({ ok: true });
  expect(format).toHaveBeenCalledWith({ kind: 'applyParagraphStyle', payload: { styleId: 'Custom' } }, { fence: true });
  parity('paragraphStyle', { styleId: 'Missing' });
  controller.empty();
  controller.toolbarMetadata(result);
  expect(controller.environment().styles.some(option => option.args.styleId === 'Custom')).toBe(false);
});

test('empty worker style metadata offers the same fallback presets', async () => {
  const { controller } = toolbarFixture();
  expect(controller.commands.store.getState('paragraphStyle', { styleId: 'Heading1' }).enabled).toBe(true);
});
