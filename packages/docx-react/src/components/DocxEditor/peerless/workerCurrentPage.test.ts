import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, mock, test } from 'bun:test';
import { readWorkerSection, type WorkerShellController } from './workerProperties';
import { useWorkerRulerControls } from './workerPropertyControls';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { renderHook, waitFor } = await import('@testing-library/react');
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const SECTIONS = [{ section: 's0', properties: { marginLeft: 1 } }, { section: 's1', properties: { marginLeft: 2 } }, { section: 'final', properties: { marginLeft: 3 } }];

/** Forty pages, the last twenty in section 1, with the reader on page 30. */
function controllerOf() {
  const getCurrentPage = mock(() => 30);
  const pages = Array.from({ length: 40 }, (_, pageIndex) => ({ page: { pageIndex }, displayPage: { sectionIndex: pageIndex < 20 ? 0 : 1 } }));
  const snapshot = { status: 'ready', view: { geometry: { pages } } };
  const value = {
    query: async (request: { kind: string }) => request.kind === 'partSnapshot' ? { value: { metadata: { body: { sections: [{}, {}, {}] } } } }
      : request.kind === 'sectionProperties' ? { version: 'v1', selection: null, value: { sections: SECTIONS, paragraphs: [] } } : { value: request },
    ref: { getCurrentPage }, getSnapshot: () => snapshot, subscribe: () => () => {},
    environment: () => ({ readOnly: true, mode: 'viewing' }),
  } as unknown as WorkerShellController;
  return { getCurrentPage, value };
}

test('reading the current section measures the current page once', async () => {
  const { value, getCurrentPage } = controllerOf();
  expect(await readWorkerSection(value)).toMatchObject({ value: { target: { kind: 'section', section: '1' } } });
  expect(getCurrentPage).toHaveBeenCalledTimes(1);
});

test('a ruler render with its section properties measures the current page once and shows its section', async () => {
  const { value, getCurrentPage } = controllerOf();
  const { result, unmount } = renderHook(() => useWorkerRulerControls(value));
  await waitFor(() => expect(result.current.sectionProps).toBeDefined());
  expect(result.current.sectionProps).toBe(SECTIONS[1]!.properties);
  expect(getCurrentPage).toHaveBeenCalledTimes(1);
  unmount();
});
