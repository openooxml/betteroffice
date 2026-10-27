import { expect, test } from 'bun:test';
import * as pptx from './index';
import * as worker from './worker';
import * as collaboration from './collaboration';

const SHIPPED_ROOT_VALUES = [
  'CollaborationError', 'CollaborationProvider', 'PRESENCE_LABEL_DURATION_MS', 'PptxExportError', 'StaleProposalError',
  'decodePresentationImage', 'exportPptxMarkdown', 'exportPptxStructured', 'initWasm', 'inspectPresentation',
  'isProposalsAvailable', 'isWasmAvailable', 'needsElementDecode', 'openPresentation', 'paintSlide',
  'presenceColorForClientId', 'presentationImageBlob', 'renderPptxMarkdown', 'sizeCanvasForSlide', 'slideToPng', 'wasmVersion',
];

const STATIC_EDITOR_INTERNALS = ['PptxPeerNotReadyError', 'PptxWorkerEditorDisposedError', 'openPresentationSession',
  'createPptxWorkerEditorSession'];

test('root keeps every shipped handle export', () => {
  for (const name of SHIPPED_ROOT_VALUES) expect(pptx).toHaveProperty(name);
});

test('worker exports live only in the worker entry', () => {
  for (const name of ['createPptxWorkerEditorSession', 'openPresentationSession', 'PptxPeerHydrationError',
    'PptxPeerNotReadyError', 'PptxWorkerEditorCollaborationError', 'PptxWorkerEditorDisposedError',
    'PptxWorkerEditorFailedError', 'createPptxWorkerSession', 'PptxWorkerUnsupportedError']) {
    expect(worker).toHaveProperty(name);
  }
  for (const name of Object.keys(worker)) {
    if (STATIC_EDITOR_INTERNALS.includes(name)) expect((pptx as Record<string, unknown>)[name]).toBe((worker as Record<string, unknown>)[name]);
    else expect(pptx).not.toHaveProperty(name);
  }
});

test('collaboration entry matches the root collaboration exports', () => {
  for (const [name, value] of Object.entries(collaboration)) {
    if (name === 'attachCollaborationPipe' || name === 'collaborationPipeFor') {
      expect(pptx).not.toHaveProperty(name);
      continue;
    }
    expect((pptx as Record<string, unknown>)[name]).toBe(value);
  }
});

test('package entries are the root, worker and collaboration', async () => {
  const manifest = await Bun.file(new URL('../package.json', import.meta.url)).json() as { exports: Record<string, unknown> };
  expect(Object.keys(manifest.exports)).toEqual(['.', './worker', './collaboration']);
});

test('sync access refusals share the cross-format code', () => {
  expect(new worker.PptxWorkerUnsupportedError('handle')).toMatchObject({ code: 'sync-api-unavailable', operation: 'handle' });
});
