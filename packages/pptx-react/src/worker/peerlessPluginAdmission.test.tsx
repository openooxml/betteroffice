import { open } from './peerlessTestHarness';
import { expect, test } from 'bun:test';
import { definePptxPlugin } from '../plugins/definePptxPlugin';
import type { PptxPluginContext } from '../plugins/types';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

test('plugin admission refuses during resize and preserves the pointer commit', async () => {
  let context: PptxPluginContext<null> | undefined;
  const plugin = definePptxPlugin({ id: 'resize', createState: () => null, onEvent(next) { context = next; } });
  const value = await open({ plugins: [plugin], pluginGrants: { resize: { document: 'write', editBatches: true } } });
  await waitFor(() => expect(context).toBeDefined());
  await value.select();
  const read = await value.api.readContent();
  if (!read.ok) throw new Error(read.failure.message);
  const request = { expectVersion: read.version, steps: [{ op: 'setSlideNotes' as const,
    target: { slideId: read.slides[0]!.id }, text: 'Blocked resize' }] };
  const start = value.point(value.box.x + value.box.w, value.box.y + value.box.h);
  const end = { ...start, clientX: start.clientX + 40, clientY: start.clientY + 40, buttons: 1 };
  act(() => { fireEvent.pointerDown(value.view.getByTestId('pptx-resize-se'), start); fireEvent.pointerMove(value.stage, end); });
  await expect(value.api.applyEdits(request)).rejects.toThrow('gesture-active');
  let result: unknown;
  await act(async () => { result = await context!.edits!.applyEdits(request); });
  expect(result).toMatchObject({ ok: false, failure: { code: 'input-failed' } });
  expect(await context!.read.readContent()).toMatchObject({ ok: false, failure: { code: 'input-failed' } });
  expect(await context!.navigation.goToSlide({ slideId: read.slides[1]!.id }, { expectVersion: read.version }))
    .toMatchObject({ ok: false, failure: { code: 'input-failed' } });
  expect(value.commands.filter((command) => command.kind === 'applyEdits')).toHaveLength(0);
  await act(async () => { fireEvent.pointerUp(value.stage, end); await value.api.flushPendingInput(); });
  expect(value.commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(1);
  const after = await value.api.readContent();
  expect(after.ok && after.slides[0]!.notes).not.toBe('Blocked resize');
});
