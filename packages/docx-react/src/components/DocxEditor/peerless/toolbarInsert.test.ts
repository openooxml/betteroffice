import { expect, mock, test } from 'bun:test';
import { toolbarFixture } from './toolbarCommands.testing';
import type { DocxCommandResult } from '../../../commands/types';

test('worker table insertion submits dimensions through the shared tail', async () => {
  const { controller, submit, flush, parity } = toolbarFixture();
  parity('insertTable', { rows: 2, columns: 3 });
  expect(controller.commands.store.getState('insertTable').enabled).toBe(true);
  flush.mockImplementation(() => new Promise<void>(() => {}));
  expect(await controller.commands.store.execute('insertTable', { rows: 2, columns: 3 })).toMatchObject({ ok: true });
  expect(submit).toHaveBeenCalledWith({ kind: 'insertTable', payload: { rows: 2, columns: 3 } }, { fence: true });
  expect(flush).not.toHaveBeenCalled();
});

test('worker insertion gates include nested body stories and reject note stories', () => {
  const { controller, viewStore, selection } = toolbarFixture();
  for (const story of ['body:t0:r0c0', 'fn:1']) {
    const view = viewStore.getSnapshot();
    viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: {
      anchor: { ...selection.anchor, story }, focus: { ...selection.focus, story },
    } });
    for (const id of ['insertTable', 'insertImage'] as const) {
      expect(controller.commands.store.getState(id).enabled).toBe(story.startsWith('body:'));
    }
  }
});

test('worker picked images preserve the hybrid payload and refuse a moved selection', async () => {
  const { controller, submit, read, viewStore, selection, parity } = toolbarFixture();
  let insert!: (image: Readonly<Record<string, unknown>>) => Promise<DocxCommandResult>;
  const pickImage = mock((callback: typeof insert) => { insert = callback; });
  controller.update({ read: read as never, focus() {}, zoom: 1, pickImage });
  parity('insertImage');
  expect(await controller.commands.store.execute('insertImage', null)).toEqual({ ok: true, status: 'opened' });
  const image = { src: 'data:image/png;base64,AA==', alt: 'pixel.png', width: 300, height: 200,
    wrapType: 'inline', displayMode: 'inline', distTop: 0, distBottom: 0, distLeft: 0, distRight: 0 };
  expect(await insert(image)).toMatchObject({ ok: true });
  expect(submit).toHaveBeenCalledWith({ kind: 'insertImage', payload: { image } }, { fence: true });
  await controller.commands.store.execute('insertImage', null);
  const view = viewStore.getSnapshot();
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: { ...selection, focus: { ...selection.focus, offset: 2 } } });
  expect(await insert(image)).toMatchObject({ ok: false, failure: { code: 'target-changed' } });
  expect(submit).toHaveBeenCalledTimes(1);
  await controller.commands.store.execute('insertImage', null);
  controller.empty();
  expect(await insert(image)).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
});
