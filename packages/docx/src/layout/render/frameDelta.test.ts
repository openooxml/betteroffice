import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { createEditSession, preloadEditWasm } from '../../wasm/edit';
import { applyFrameDelta, decodeFrameDelta } from './frameDelta';
import type { DisplayList } from './displayList';

const WASM = resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

describe('FrameDelta wire round-trip', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('decodes wasm-encoded full and delta frames to the equivalent JSON list', () => {
    const session = createEditSession(11);
    const { paraId } = JSON.parse(session.create_story('body', 'Hello frame', 'Normal', 'left'));
    const fontId = session.register_measure_font(new Uint8Array(readFileSync(FONT)));
    const request = JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: {
        fontChains: { 'calibri|0|0': [fontId] },
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: true,
      },
      renderEnv: {},
    });

    const envelopeFor = (): string => {
      const output = JSON.parse(session.layout_document_with_regions_json(request)) as {
        measured: unknown;
        options: unknown;
        layout: unknown;
      };
      return JSON.stringify({
        measured: output.measured,
        options: output.options,
        layout: output.layout,
        fontChains: { 'calibri|0|0': [fontId] },
      });
    };

    const first = envelopeFor();
    const jsonList = JSON.parse(session.build_display_list_json(first)) as DisplayList;
    const fullFrame = session.build_display_list_frame(first, 0);
    const retained = applyFrameDelta(null, decodeFrameDelta(fullFrame));
    expect(retained.displayList).toEqual(jsonList);

    session.insert_text('body', paraId, 5, ' typed', undefined, undefined);
    const second = envelopeFor();
    const nextJsonList = JSON.parse(session.build_display_list_json(second)) as DisplayList;
    const deltaFrame = session.build_display_list_frame(second, retained.frameEpoch);
    const next = applyFrameDelta(retained, decodeFrameDelta(deltaFrame));
    expect(next.displayList).toEqual(nextJsonList);
    const pageText = next.displayList.pages[0].primitives
      .map((primitive) => ('text' in primitive ? (primitive.text ?? '') : ''))
      .join('');
    expect(pageText).toContain('typed');
  });
});
