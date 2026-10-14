import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { Document, Paragraph } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, wasmModuleMemories, WASM32_MEMORY_LIMIT_BYTES } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
  )
);
const edit = () => wasmModuleMemories().find((module) => module.label === 'docx-edit')!;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

it('counts the editing core heap as sessions grow and go', async () => {
  const before = edit();
  expect(before.bufferBytes).toBeGreaterThan(0);
  expect(before.bufferBytes).toBeLessThanOrEqual(WASM32_MEMORY_LIMIT_BYTES);
  expect(before.liveBytes).toBeLessThanOrEqual(before.bufferBytes);
  expect(before.failedAllocationBytes).toBe(0);

  const session = await createYrsSession({ clientId: 76101 });
  session.createStory('body', 'x'.repeat(4 * 1024 * 1024));
  const grown = edit();
  expect(grown.liveBytes!).toBeGreaterThan(before.liveBytes! + 4 * 1024 * 1024);
  expect(grown.peakBytes!).toBeGreaterThanOrEqual(grown.liveBytes!);
  expect(grown.bufferBytes).toBeGreaterThanOrEqual(grown.liveBytes!);

  session.destroy();
  const after = edit();
  expect(after.liveBytes!).toBeLessThan(grown.liveBytes! - 4 * 1024 * 1024);
  expect(after.peakBytes!).toBeGreaterThanOrEqual(grown.liveBytes!);
  expect(after.bufferBytes).toBeGreaterThanOrEqual(grown.bufferBytes);
});

it('releases cached paragraph measurements when the last session is destroyed', async () => {
  const document: Document = {
    package: {
      document: {
        content: Array.from({ length: 400 }, (_, index): Paragraph => ({
          type: 'paragraph',
          paraId: (index + 1).toString(16).padStart(8, '0'),
          content: [{
            type: 'run',
            content: [{
              type: 'text',
              text: `Paragraph ${index}: ${'Paragraph measurement cache words. '.repeat(32)}`,
            }],
          }],
        })),
      },
    },
  };
  const before = edit().liveBytes!;
  const session = await createYrsSession({ clientId: 76102 });
  try {
    documentToYrs(session, document);
    expect(session.paragraphs('body')).toHaveLength(400);
    const font = session.registerFont(FONT);
    const request = buildResidentRegionLayoutRequest(document, 24, {});
    const requirements = JSON.parse(
      session.layoutFontRequirementsJson(JSON.stringify(request))
    ) as Array<{ key: string }>;
    expect(requirements.length).toBeGreaterThan(0);
    request.measurement = {
      fontChains: Object.fromEntries(requirements.map((requirement) => [requirement.key, [font]])),
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      compat: { noLeading: false, doNotExpandShiftReturn: false },
      authoritativeShaping: true,
    };
    const result = JSON.parse(
      session.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request))
    ) as { layout: { pages: unknown[] } };
    expect(result.layout.pages.length).toBeGreaterThan(1);
    expect(edit().liveBytes!).toBeGreaterThan(before + 2 * 1024 * 1024);
  } finally {
    session.destroy();
  }
  expect(edit().liveBytes!).toBeLessThanOrEqual(before + 2 * 1024 * 1024);
});
