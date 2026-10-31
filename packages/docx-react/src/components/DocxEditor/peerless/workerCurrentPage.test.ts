import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, mock, test } from 'bun:test';
import type { SectionProperties } from '@betteroffice/docx/types/document';
import type { PeerlessInputIntent } from './inputEvents';
import { pageOrientationPatch } from '../../pageLayout';
import { readWorkerSection, type WorkerPropertyQuery, type WorkerShellController } from './workerProperties';
import { useWorkerPageSetupControls, useWorkerRulerControls } from './workerPropertyControls';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, renderHook, waitFor } = await import('@testing-library/react');
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

test('page setup reads its current section only when Layout requests it', async () => {
  const { value } = controllerOf();
  const query = mock(async (request: { kind: string; target?: { section: string } }) =>
    request.kind === 'partSnapshot'
      ? { value: { metadata: { body: { sections: [{}, {}, {}] } } } }
      : { value: { kind: 'section', section: request.target?.section, properties: { marginLeft: 720 } } });
  const controller = { ...value, query } as unknown as WorkerShellController;
  const { result, rerender, unmount } = renderHook(() => useWorkerPageSetupControls(controller));
  expect(query).not.toHaveBeenCalled();
  rerender();
  expect(query).not.toHaveBeenCalled();
  await act(async () => { await result.current.readCurrent(); });
  expect(query).toHaveBeenCalledTimes(2);
  expect(query).toHaveBeenLastCalledWith({ kind: 'properties', target: { kind: 'section', section: '1' } });
  expect(result.current.currentProps).toEqual({ marginLeft: 720 });
  expect(result.current.showPageSetup).toBe(false);
  rerender();
  expect(query).toHaveBeenCalledTimes(2);
  await act(async () => { await result.current.readCurrent(); });
  expect(query).toHaveBeenCalledTimes(4);
  unmount();
});

function pageSetupFixture() {
  const { value, getCurrentPage } = controllerOf();
  const current = {
    version: 'v1',
    sections: {
      s0: { pageWidth: 12240, pageHeight: 20160 },
      s1: { pageWidth: 12240, pageHeight: 15840 },
      final: {},
    } as Record<string, SectionProperties>,
  };
  const reads: WorkerPropertyQuery[] = [];
  const query = mock(async (request: { kind: string; target?: { section: string } }) => {
    if (request.kind === 'partSnapshot') {
      return { value: { metadata: { body: { sections: [{ id: 's0' }, { id: 's1' }, {}] } } } };
    }
    const section = request.target?.section ?? 'final';
    const result: WorkerPropertyQuery = { version: current.version, selection: null,
      value: { kind: 'section', section, properties: { ...current.sections[section] } } };
    reads.push(result);
    return result;
  });
  const applyQuery = mock(async (source: WorkerPropertyQuery, intent: PeerlessInputIntent) => {
    if (intent.kind !== 'setSectionProperties') throw new Error('Unexpected property command');
    if (source.version !== current.version || intent.payload.expectVersion !== current.version) {
      return { ok: false, failure: { code: 'command-failed', message: 'The document version changed.' } };
    }
    const { section, patch } = intent.payload;
    current.sections[section] = { ...current.sections[section], ...patch };
    current.version = 'v2';
    return { ok: true, status: 'executed' };
  });
  const focus = mock(() => {});
  const controller = { ...value, query, applyQuery,
    environment: () => ({ readOnly: false, mode: 'editing' }),
    ref: { ...value.ref, flushPendingInput: async () => {}, focus },
  } as unknown as WorkerShellController;
  const onError = mock((_error: Error) => {});
  const hook = renderHook(() => useWorkerPageSetupControls(controller, onError));
  return { ...hook, current, reads, query, applyQuery, getCurrentPage, onError, focus };
}

test('Layout applies to the section captured on open after the viewport scrolls', async () => {
  const view = pageSetupFixture();
  await act(async () => { await view.result.current.readCurrent(); });
  const opened = view.reads[0]!;
  expect(opened.value).toMatchObject({ kind: 'section', section: 's1' });
  view.getCurrentPage.mockReturnValue(1);
  const patch = { marginTop: 720, marginBottom: 720, marginLeft: 720, marginRight: 720 };
  await act(async () => { await view.result.current.handlePageSetupApply(patch); });
  expect(view.applyQuery).toHaveBeenCalledTimes(1);
  expect(view.applyQuery.mock.calls[0]![0]).toBe(opened);
  expect(view.applyQuery.mock.calls[0]![1]).toEqual({ kind: 'setSectionProperties',
    payload: { section: 's1', patch, expectVersion: 'v1' } });
  expect(view.current.sections.s1).toMatchObject(patch);
  expect(view.current.sections.s0).toEqual({ pageWidth: 12240, pageHeight: 20160 });
  expect(view.onError).not.toHaveBeenCalled();
  view.unmount();
});

test('Layout refuses a stale orientation change instead of overwriting a collaborator paper size', async () => {
  const view = pageSetupFixture();
  await act(async () => { await view.result.current.readCurrent(); });
  const opened = view.reads[0]!;
  const patch = pageOrientationPatch(view.result.current.currentProps, 'landscape');
  view.current.version = 'v2';
  view.current.sections.s1 = { pageWidth: 11906, pageHeight: 16838 };
  await act(async () => { await view.result.current.handlePageSetupApply(patch); });
  expect(view.applyQuery).toHaveBeenCalledTimes(1);
  expect(view.applyQuery.mock.calls[0]![0]).toBe(opened);
  expect(view.applyQuery.mock.calls[0]![1]).toEqual({ kind: 'setSectionProperties', payload: {
    section: 's1', expectVersion: 'v1', patch: { orientation: 'landscape', pageWidth: 15840, pageHeight: 12240 },
  } });
  expect(view.current.sections.s1).toEqual({ pageWidth: 11906, pageHeight: 16838 });
  expect(view.onError).toHaveBeenCalledTimes(1);
  expect(view.onError.mock.calls[0]![0].message).toBe('The document version changed.');
  expect(view.query).toHaveBeenCalledTimes(2);
  expect(view.focus).not.toHaveBeenCalled();
  view.unmount();
});
