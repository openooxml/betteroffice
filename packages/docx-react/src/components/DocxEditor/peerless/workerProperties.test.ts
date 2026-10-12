import { expect, mock, test } from 'bun:test';
import { readWorkerSection, type WorkerShellController } from './workerProperties';

for (const [index, id, section] of [[0, undefined, '0'], [0, 'first', 'first'], [1, undefined, 'final']] as const) {
  test(`page setup selects section ${section} on page ${index + 1}`, async () => {
    const query = mock(async (request: { kind: string }) => request.kind === 'partSnapshot'
      ? { value: { metadata: { body: { sections: [{ id }, {}] } } } } : { value: request });
    const controller = { query, ref: { getCurrentPage: () => index + 1 }, getSnapshot: () => ({
      view: { geometry: { pages: [{ page: { pageIndex: index }, displayPage: { sectionIndex: index } }] } },
    }) } as unknown as WorkerShellController;
    await readWorkerSection(controller);
    expect(query).toHaveBeenLastCalledWith({ kind: 'properties', target: { kind: 'section', section } });
  });
}
