import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { applyFrameDelta, decodeFrameDelta } from '../../packages/docx/src/layout/render/frameDelta';
import initWasm, * as wasm from '../../packages/docx/src/wasm/generated/edit/docx_edit.js';
import initPreflight, { font_preflight as wasmFontPreflight } from '../../packages/docx/src/wasm/generated/preflight/docx_preflight.js';

import { openFailureKind } from '../../packages/docx-react/src/components/openFailure';

type Engine = typeof wasm;
type Session = wasm.EditSession;
type ParagraphRead = {
  ok: boolean;
  version: string;
  paragraphs: { story: string; paraId: string; text: string }[];
};

const native = require('./index.js') as Engine;
const request = JSON.stringify({ story: 'body', view: 'accepted' });
let fixture: Uint8Array;
let fontBytes: Uint8Array;

beforeAll(async () => {
  await initWasm({
    module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)),
  });
  await initPreflight({
    module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/preflight/docx_preflight_bg.wasm', import.meta.url)),
  });
  fontBytes = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
  const encoder = new TextEncoder();
  fixture = wasm.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Engine text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'),
  });
});

function read(session: Session): ParagraphRead {
  const result = JSON.parse(session.read_paragraphs_json(request)) as ParagraphRead;
  expect(result.ok).toBe(true);
  expect(result.version).toMatch(/^[0-9a-f]{16}-\d+$/);
  return { ...result, version: '<version>' };
}

function paragraphId(session: Session): string {
  const paragraph = read(session).paragraphs[0];
  expect(paragraph).toBeDefined();
  return paragraph.paraId;
}

function expectPlainBytes(value: Uint8Array) {
  expect(value).toBeInstanceOf(Uint8Array);
  expect(Buffer.isBuffer(value)).toBe(false);
  expect(Object.getPrototypeOf(value)).toBe(Uint8Array.prototype);
  expect(value.byteOffset).toBe(0);
  expect(value.buffer.byteLength).toBe(value.byteLength);
}

function caughtValue(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected the engine call to throw');
}

function expectThrownParity(wasmCall: () => unknown, nativeCall: () => unknown): unknown {
  const expected = caughtValue(wasmCall);
  const actual = caughtValue(nativeCall);
  expect(typeof actual).toBe(typeof expected);
  expect(actual instanceof Error).toBe(expected instanceof Error);
  if (actual instanceof Error && expected instanceof Error) {
    expect(actual.name).toBe(expected.name);
    expect(actual.message).toBe(expected.message);
    expect((actual as Error & { code?: string }).code)
      .toBe((expected as Error & { code?: string }).code);
  } else {
    expect(actual).toBe(expected);
  }
  return actual;
}

describe('docx native engine parity', () => {
  test('an empty docDefaults element leaves unstyled paragraphs without spacing in both engines', () => {
    const encoder = new TextEncoder();
    const bytes = wasm.rezip_docx({
      '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'),
      '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
      'word/_rels/document.xml.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
      'word/styles.xml': encoder.encode('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style></w:styles>'),
      'word/document.xml': encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Engine text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'),
    });
    const expected = new wasm.EditSession(23);
    const actual = new native.EditSession(23);
    try {
      const lowered = (session: Session) => {
        session.seed_from_docx(bytes, 'opening');
        return JSON.parse(session.yrs_blocks_for_story('body', '{}')) as Array<{ kind: string; attrs?: { spacing?: { after?: number } } }>;
      };
      const wasmBlocks = lowered(expected);
      expect(lowered(actual)).toEqual(wasmBlocks);
      expect(wasmBlocks.find((block) => block.kind === 'paragraph')?.attrs?.spacing?.after ?? 0).toBe(0);
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('an inserted table takes the document table style margins in both engines', () => {
    const encoder = new TextEncoder();
    const styled = wasm.rezip_docx({
      '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'),
      '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
      'word/_rels/document.xml.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
      'word/styles.xml': encoder.encode('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>'),
      'word/document.xml': encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Engine text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'),
    });
    const hostEnv = JSON.stringify({
      tableStyleCellMargins: { TableNormal: { left: { value: 108, type: 'dxa' }, right: { value: 108, type: 'dxa' } } },
      defaultTableStyleId: 'TableNormal',
    });
    for (const [bytes, env] of [[fixture, hostEnv], [styled, '{}']] as const) {
      const expected = new wasm.EditSession(17);
      const actual = new native.EditSession(17);
      try {
        const lowered = (session: Session) => {
          session.seed_from_docx(bytes, 'opening');
          session.insert_table('body', paragraphId(session), 0, 1, 2, 'Owner', '2026-11-07T12:00:00Z');
          return JSON.parse(session.yrs_blocks_for_story('body', env)) as Array<{
            kind: string;
            rows?: Array<{ cells: Array<{ padding?: { left: number; right: number; top: number; bottom: number } }> }>;
          }>;
        };
        const wasmBlocks = lowered(expected);
        const nativeBlocks = lowered(actual);
        expect(nativeBlocks).toEqual(wasmBlocks);
        const padding = wasmBlocks.find((block) => block.kind === 'table')?.rows?.[0]?.cells[0]?.padding;
        expect(padding?.left).toBeCloseTo(7.2, 6);
        expect(padding?.right).toBeCloseTo(7.2, 6);
      } finally {
        actual.free();
        expected.free();
      }
    }
  });

  test('endnotes taller than the last page flow onto new pages, in both engines', () => {
    const encoder = new TextEncoder();
    const main = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const officeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
    const relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const words = (note: number) => Array.from({ length: 40 }, (_, index) => `n${note}t${index}`).join(' ');
    const notes = [1, 2, 3].map((id) => `<w:endnote w:id="${id}"><w:p><w:r><w:t xml:space="preserve">${words(id)}</w:t></w:r></w:p></w:endnote>`).join('');
    const body = [1, 2, 3].map((id) => `<w:p><w:r><w:t>Body ${id}</w:t></w:r><w:r><w:endnoteReference w:id="${id}"/></w:r></w:p>`).join('');
    const bytes = wasm.rezip_docx({
      '[Content_Types].xml': encoder.encode(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="${officeType}.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="${officeType}.endnotes+xml"/></Types>`),
      '_rels/.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relType}/officeDocument" Target="word/document.xml"/></Relationships>`),
      'word/_rels/document.xml.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="${relType}/endnotes" Target="endnotes.xml"/></Relationships>`),
      'word/endnotes.xml': encoder.encode(`<w:endnotes ${main}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>${notes}</w:endnotes>`),
      'word/document.xml': encoder.encode(`<w:document ${main}><w:body>${body}<w:sectPr><w:pgSz w:w="4320" w:h="2880"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`),
    });
    const expected = new wasm.EditSession(29);
    const actual = new native.EditSession(29);
    try {
      expected.open_docx(bytes, true, 'opening');
      actual.open_docx(bytes, true, 'opening');
      const font = expected.register_measure_font(fontBytes);
      expect(actual.register_measure_font(fontBytes)).toBe(font);
      const request = {
        bodyStory: 'body',
        regions: { sections: [{
          sectionId: 'main',
          properties: {
            pageWidth: 4320, pageHeight: 2880,
            marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
          },
        }] },
        notes: { contents: [1, 2, 3].map((id) => ({ id, noteKind: 'endnote', height: 0 })) },
        measurement: {
          fontChains: {} as Record<string, number[]>,
          defaults: { fontSize: 11, fontFamily: 'Calibri' },
          authoritativeShaping: true,
        },
        renderEnv: {},
      };
      const requirements = JSON.parse(expected.layout_font_requirements_json(JSON.stringify(request))) as { key: string }[];
      for (const requirement of requirements) request.measurement.fontChains[requirement.key] = [font];
      const input = JSON.stringify(request);
      const layout = actual.layout_document_with_regions_json(input);
      expect(layout).toBe(expected.layout_document_with_regions_json(input));
      type Area = { kind?: string; y: number; height: number; continuation?: boolean; notes: { id: number }[] };
      type Page = { fragments: { y: number; height: number }[]; size: { h: number }; margins: { bottom: number }; noteAreas?: Area[] };
      const pages = (JSON.parse(layout) as { layout: { pages: Page[] } }).layout.pages;
      const areas = pages.flatMap((page) => (page.noteAreas ?? []).filter((area) => area.kind === 'endnote').map((area) => ({ page, area })));
      expect(areas.length).toBeGreaterThan(1);
      expect(areas[0]!.area.continuation).toBeUndefined();
      for (const { area } of areas.slice(1)) expect(area.continuation).toBe(true);
      for (const { page, area } of areas) {
        const body = Math.max(0, ...page.fragments.map((fragment) => fragment.y + fragment.height));
        expect(area.y + 0.5).toBeGreaterThanOrEqual(body);
        expect(area.y + area.height).toBeLessThanOrEqual(page.size.h - page.margins.bottom + 0.5);
      }
      expect(areas.flatMap(({ area }) => area.notes.map((note) => note.id)).filter((id, index, ids) => ids[index - 1] !== id)).toEqual([1, 2, 3]);
      const display = actual.build_display_list_json(layout);
      expect(display).toBe(expected.build_display_list_json(layout));
      type DisplayArea = { kind?: string; primitives?: { groupId?: string }[]; continuedNoteIds?: number[] };
      const displayPages = (JSON.parse(display) as { pages: { noteAreas?: DisplayArea[] }[] }).pages;
      const shownOn = (id: number) => displayPages.filter((page) => (page.noteAreas ?? []).some((area) =>
        area.kind === 'endnote' && (area.primitives ?? []).some((primitive) => primitive.groupId === `endnote-${id}`))).length;
      const split = [1, 2, 3].filter((id) => shownOn(id) === 2);
      expect(split.length).toBeGreaterThan(0);
      for (const id of split) {
        expect(displayPages.filter((page) => (page.noteAreas ?? []).some((area) => area.continuedNoteIds?.includes(id))).length).toBe(1);
      }
    } finally {
      actual.free();
      expected.free();
    }
  });

  test.each([
    ['Arabic', 'مرحبا', 'بالعالم', true, () => new Uint8Array(readFileSync(new URL('../../packages/fonts/assets/NotoNaskhArabic-Regular.ttf', import.meta.url)))],
    ['Devanagari', 'नमस्ते', 'क्षत्रिय', false, () => wasm.decodeFontFace(new Uint8Array(readFileSync(new URL('../../packages/fonts-intl/assets/NotoSansDevanagari-Regular.ttf.zst', import.meta.url))), 'zstd').font],
  ] as const)('measured glyph runs of a %s paragraph match in both engines', (_script, first, second, rtl, scriptFont) => {
    const encoder = new TextEncoder();
    const main = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const words = Array.from({ length: 12 }, () => `${first} ${second}`).join(' ');
    const bytes = wasm.rezip_docx({
      '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
      '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
      'word/document.xml': encoder.encode(`<w:document ${main}><w:body><w:p>${rtl ? '<w:pPr><w:bidi/></w:pPr>' : ''}<w:r>${rtl ? '<w:rPr><w:rtl/></w:rPr>' : ''}<w:t xml:space="preserve">${words}</w:t></w:r><w:r><w:t xml:space="preserve"> Latin tail</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
    });
    const scriptBytes = scriptFont();
    const expected = new wasm.EditSession(31);
    const actual = new native.EditSession(31);
    try {
      expected.open_docx(bytes, true, 'opening');
      actual.open_docx(bytes, true, 'opening');
      const latin = expected.register_measure_font(fontBytes);
      expect(actual.register_measure_font(fontBytes)).toBe(latin);
      const script = expected.register_measure_font(scriptBytes);
      expect(actual.register_measure_font(scriptBytes)).toBe(script);
      const request = {
        bodyStory: 'body',
        regions: { sections: [{
          sectionId: 'main',
          properties: {
            pageWidth: 4320, pageHeight: 15840,
            marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
          },
        }] },
        measurement: {
          fontChains: {} as Record<string, number[]>,
          defaults: { fontSize: 11, fontFamily: 'Calibri' },
          authoritativeShaping: true,
        },
        renderEnv: {},
      };
      const requirements = JSON.parse(expected.layout_font_requirements_json(JSON.stringify(request))) as { key: string }[];
      for (const requirement of requirements) request.measurement.fontChains[requirement.key] = [latin, script];
      const input = JSON.stringify(request);
      const layout = actual.layout_document_with_regions_json(input);
      const expectedLayout = expected.layout_document_with_regions_json(input);
      expect(layout).toBe(expectedLayout);
      const glyphRuns = (value: unknown): string[] => {
        if (Array.isArray(value)) return value.flatMap(glyphRuns);
        if (value === null || typeof value !== 'object') return [];
        const record = value as Record<string, unknown>;
        const own = typeof record.glyphRuns === 'string' ? [record.glyphRuns] : [];
        return [...own, ...Object.values(record).flatMap(glyphRuns)];
      };
      const lines = glyphRuns(JSON.parse(layout));
      expect(lines.length).toBeGreaterThan(1);
      expect(lines).toEqual(glyphRuns(JSON.parse(expectedLayout)));
      const display = actual.build_display_list_json(layout);
      expect(display).toBe(expected.build_display_list_json(layout));
      // Only the font-aware frame build shapes; the JSON build above has no font store.
      const extras = JSON.stringify({ fontChains: request.measurement.fontChains });
      const frame = actual.build_display_list_frame(extras, 0);
      expect(frame).toEqual(expected.build_display_list_frame(extras, 0));
      const painted = applyFrameDelta(null, decodeFrameDelta(frame)).pages
        .flatMap((page) => page.page.primitives)
        .flatMap((primitive) => (primitive.kind === 'glyphRun' ? [primitive.text] : []));
      expect(painted).toContain(first);
      expect(painted).toContain(second);
    } finally {
      actual.free();
      expected.free();
    }
  });

  test.each(['open_docx', 'seed_from_docx'] as const)('%s reads the same synthetic package', (method) => {
    const expected = new wasm.EditSession(7);
    const actual = new native.EditSession(7);
    try {
      if (method === 'open_docx') {
        expected.open_docx(fixture, true, 'opening');
        actual.open_docx(fixture, true, 'opening');
      } else {
        expected.seed_from_docx(fixture, 'opening');
        actual.seed_from_docx(fixture, 'opening');
      }
      expect(actual.client_id()).toBe(7);
      expect(read(actual)).toEqual(read(expected));
      expect(read(actual).paragraphs.map((paragraph) => paragraph.text)).toEqual(['Engine text']);
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('set_row_height matches wasm and refuses the same input', () => {
    const expected = new wasm.EditSession(13);
    const actual = new native.EditSession(13);
    try {
      for (const session of [expected, actual]) {
        session.seed_from_docx(fixture, 'opening');
        session.insert_table('body', paragraphId(session), 0, 2, 2);
      }
      const at = JSON.stringify({ story: 'body', tableIndex: 0, row: 1, column: 0 });
      expect(JSON.parse(actual.set_row_height(at, 720, 'atLeast'))).toEqual(JSON.parse(expected.set_row_height(at, 720, 'atLeast')));
      const payload = JSON.parse(actual.table_payload('body', 0)!);
      expect(payload).toEqual(JSON.parse(expected.table_payload('body', 0)!));
      expect(payload.rows[1].trPr).toMatchObject({ height: 720, heightRule: 'atLeast' });
      expectThrownParity(() => expected.set_row_height(at, 0, 'atLeast'), () => actual.set_row_height(at, 0, 'atLeast'));
      expectThrownParity(() => expected.set_row_height(at, 720, 'auto'), () => actual.set_row_height(at, 720, 'auto'));
    } finally {
      expected.free();
      actual.free();
    }
  });

  test('native insert_text supports undo and redo', () => {
    const expected = new wasm.EditSession(11);
    const session = new native.EditSession(11);
    try {
      expected.seed_from_docx(fixture, 'opening');
      session.seed_from_docx(fixture, 'opening');
      const before = read(session);
      session.track_undo();
      expect(session.can_undo()).toBe(false);
      expect(session.insert_text('body', paragraphId(session), 0, 'Native '))
        .toBe(expected.insert_text('body', paragraphId(expected), 0, 'Native '));
      const edited = read(session);
      expect(edited).toEqual(read(expected));
      expect(edited.paragraphs[0].text).toBe('Native Engine text');
      expect(session.can_undo()).toBe(true);
      expect(session.undo()).toBe(true);
      expect(read(session)).toEqual(before);
      expect(session.can_redo()).toBe(true);
      expect(session.redo()).toBe(true);
      expect(read(session)).toEqual(edited);
    } finally {
      expected.free();
      session.free();
    }
  });

  test('undo grouping reads the JavaScript clock, as wasm does', () => {
    let time = 1_000;
    const clock = spyOn(Date, 'now').mockImplementation(() => time);
    const texts = (engine: Engine) => {
      const session = new engine.EditSession(13);
      try {
        session.seed_from_docx(fixture, 'opening');
        session.track_undo();
        const steps: string[] = [];
        for (const [at, text] of [[1_000, 'A'], [1_100, 'B'], [2_000, 'C']] as const) {
          time = at;
          session.insert_text('body', paragraphId(session), 0, text);
        }
        steps.push(read(session).paragraphs[0].text);
        while (session.undo()) steps.push(read(session).paragraphs[0].text);
        return steps;
      } finally {
        session.free();
      }
    };
    try {
      const expected = texts(wasm);
      expect(expected).toEqual(['CBAEngine text', 'BAEngine text', 'Engine text']);
      expect(texts(native)).toEqual(expected);
    } finally {
      clock.mockRestore();
    }
  });

  test('native edits survive a throwing clock and a clock that reads the session', () => {
    const session = new native.EditSession(17);
    let reentrant: (() => unknown) | undefined;
    let depth = 0;
    const clock = spyOn(Date, 'now').mockImplementation(() => {
      if (!reentrant) throw new Error('clock unavailable');
      if (depth === 0) {
        depth += 1;
        try {
          reentrant();
        } finally {
          depth -= 1;
        }
      }
      return 5_000;
    });
    try {
      session.seed_from_docx(fixture, 'opening');
      session.track_undo();
      session.insert_text('body', paragraphId(session), 0, 'A');
      expect(read(session).paragraphs[0].text).toBe('AEngine text');
      reentrant = () => read(session);
      session.insert_text('body', paragraphId(session), 0, 'B');
      expect(read(session).paragraphs[0].text).toBe('BAEngine text');
      expect(session.undo()).toBe(true);
    } finally {
      clock.mockRestore();
      session.free();
    }
  });

  test('worker updates applied as local edits group by the JavaScript clock, as wasm does', () => {
    let time = 1_000;
    const clock = spyOn(Date, 'now').mockImplementation(() => time);
    const texts = (engine: Engine) => {
      const source = new wasm.EditSession(21);
      const target = new engine.EditSession(23);
      try {
        source.seed_from_docx(fixture, 'opening');
        target.load(source.encode_state());
        target.track_undo();
        for (const [at, text] of [[1_000, 'A'], [1_100, 'B'], [2_000, 'C']] as const) {
          source.insert_text('body', paragraphId(source), 0, text);
          const update = source.encode_diff(target.encode_state_vector());
          time = at;
          target.apply_local_update(update);
        }
        const steps = [read(target).paragraphs[0].text];
        while (target.undo()) steps.push(read(target).paragraphs[0].text);
        return steps;
      } finally {
        target.free();
        source.free();
      }
    };
    try {
      const expected = texts(wasm);
      expect(expected).toEqual(['CBAEngine text', 'BAEngine text', 'Engine text']);
      expect(texts(native)).toEqual(expected);
    } finally {
      clock.mockRestore();
    }
  });

  test('native calls that borrow typed arrays never run the JavaScript clock', () => {
    const session = new native.EditSession(19);
    const bytes = new Uint8Array(fontBytes);
    let reads = 0;
    const clock = spyOn(Date, 'now').mockImplementation(() => {
      reads += 1;
      bytes.buffer.transfer(0);
      return 1_000;
    });
    try {
      expect(session.register_measure_font(bytes)).toBe(0);
      expect(reads).toBe(0);
      expect(bytes.byteLength).toBe(fontBytes.byteLength);
      session.seed_from_docx(fixture, 'opening');
      const vector = new Uint8Array(new Uint8Array([0]));
      clock.mockImplementation(() => {
        reads += 1;
        vector.buffer.transfer(0);
        return 1_000;
      });
      expect(session.encode_diff_chunks(vector, 1 << 20).length).toBeGreaterThan(0);
      expect(reads).toBe(0);
      expect(vector.byteLength).toBe(1);
    } finally {
      clock.mockRestore();
      session.free();
    }
  });

  test('state crosses both boundaries and native diffs update a wasm peer', () => {
    const nativeSource = new native.EditSession(21);
    const wasmPeer = new wasm.EditSession(22);
    const wasmSource = new wasm.EditSession(23);
    const nativePeer = new native.EditSession(24);
    try {
      nativeSource.seed_from_docx(fixture, 'opening');
      const nativeState = nativeSource.encode_state();
      expectPlainBytes(nativeState);
      wasmPeer.load(nativeState);
      expect(read(wasmPeer)).toEqual(read(nativeSource));

      wasmSource.seed_from_docx(fixture, 'opening');
      nativePeer.load(wasmSource.encode_state());
      expect(read(nativePeer)).toEqual(read(wasmSource));

      nativeSource.insert_text('body', paragraphId(nativeSource), 0, 'Synced ');
      const diff = nativeSource.encode_diff(wasmPeer.encode_state_vector());
      expectPlainBytes(diff);
      expect(wasmPeer.apply_update(diff)).toBeUndefined();
      expect(read(wasmPeer)).toEqual(read(nativeSource));

      wasmSource.insert_text('body', paragraphId(wasmSource), 0, 'Remote ');
      nativePeer.apply_update(wasmSource.encode_diff(nativePeer.encode_state_vector()));
      expect(read(nativePeer)).toEqual(read(wasmSource));
    } finally {
      nativePeer.free();
      wasmSource.free();
      wasmPeer.free();
      nativeSource.free();
    }
  });

  test('chunk arrays contain plain bytes and reject invalid entries before applying', () => {
    const source = new native.EditSession(31);
    const peer = new wasm.EditSession(32);
    const target = new native.EditSession(33);
    try {
      source.seed_from_docx(fixture, 'opening');
      const chunks = source.encode_diff_chunks(peer.encode_state_vector(), 64);
      expect(Array.isArray(chunks)).toBe(true);
      expect(chunks.length).toBeGreaterThan(0);
      for (const chunk of chunks) expectPlainBytes(chunk);
      peer.apply_update_chunks(chunks);
      expect(read(peer)).toEqual(read(source));

      const fromWasm = peer.encode_diff_chunks(target.encode_state_vector(), 64);
      const before = target.encode_state();
      expect(expectThrownParity(
        () => peer.apply_update_chunks([chunks[0], new Uint32Array([1])]),
        () => target.apply_update_chunks([chunks[0], new Uint32Array([1])]),
      )).toBe('sync updates must be Uint8Arrays');
      expect(target.encode_state()).toEqual(before);
      target.apply_update_chunks(fromWasm);
      expect(read(target)).toEqual(read(peer));
    } finally {
      target.free();
      peer.free();
      source.free();
    }
  });

  test('queued update observation is synchronous and can be cleared', () => {
    for (const engine of [wasm, native]) {
      const session = new engine.EditSession(41);
      try {
        session.seed_from_docx(fixture, 'opening');
        expect(session.start_update_event_observation()).toBeUndefined();
        session.start_update_event_observation();
        session.insert_text('body', paragraphId(session), 0, 'Observed ');
        const update = session.drain_update_event();
        expectPlainBytes(update);
        expect(update.byteLength).toBeGreaterThan(1);
        expect(update[0]).toBe(0);
        session.clear_update_event_observation();
        expect(session.drain_update_event().byteLength).toBe(0);
        session.insert_text('body', paragraphId(session), 0, 'Unobserved ');
        expect(session.drain_update_event().byteLength).toBe(0);
      } finally {
        session.free();
      }
    }
  });

  test('resident fonts decode, shape and outline identically', () => {
    const expected = new wasm.EditSession(81);
    const actual = new native.EditSession(81);
    try {
      const font = expected.register_measure_font(fontBytes);
      expect(actual.register_measure_font(fontBytes)).toBe(font);
      expect(actual.register_substitute_measure_font(font, 'Calibri'))
        .toBe(expected.register_substitute_measure_font(font, 'Calibri'));
      const decoded = actual.decode_font(fontBytes);
      expectPlainBytes(decoded);
      expect(decoded).toEqual(expected.decode_font(fontBytes));
      expect(decoded).toEqual(fontBytes);
      expect(actual.take_base_latin_miss()).toBe(expected.take_base_latin_miss());
      const echo = actual.echo_glyph_set_json(font, 11, true, 'Ab!');
      expect(echo).toBe(expected.echo_glyph_set_json(font, 11, true, 'Ab!'));
      const glyph = JSON.parse(echo).glyphIds[0] as number;
      expect(actual.outline_glyph_json(font, glyph)).toBe(expected.outline_glyph_json(font, glyph));
      const invalid = new Uint8Array([1, 2, 3]);
      expectThrownParity(() => expected.decode_font(invalid), () => actual.decode_font(invalid));
      expectThrownParity(
        () => expected.register_measure_font(invalid),
        () => actual.register_measure_font(invalid),
      );
      expectThrownParity(
        () => expected.echo_glyph_set_json(font, 0, true, 'A'),
        () => actual.echo_glyph_set_json(font, 0, true, 'A'),
      );
      expected.clear_measure_fonts();
      actual.clear_measure_fonts();
      expect(actual.register_measure_font(fontBytes)).toBe(expected.register_measure_font(fontBytes));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('region layouts, retained metadata, frames and resident input match', () => {
    const expected = new wasm.EditSession(91);
    const actual = new native.EditSession(91);
    try {
      expected.open_docx(fixture, true, 'opening');
      actual.open_docx(fixture, true, 'opening');
      const font = expected.register_measure_font(fontBytes);
      expect(actual.register_measure_font(fontBytes)).toBe(font);
      const fontChains = { 'calibri|0|0': [font] };
      const input = JSON.stringify({
        bodyStory: 'body',
        regions: { sections: [{
          sectionId: 'main',
          properties: {
            pageWidth: 4320, pageHeight: 2880,
            marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
          },
        }] },
        measurement: {
          fontChains,
          defaults: { fontSize: 11, fontFamily: 'Calibri' },
          authoritativeShaping: true,
        },
        renderEnv: {},
      });
      expect(actual.layout_font_requirements_json(input)).toBe(expected.layout_font_requirements_json(input));
      const layout = actual.layout_document_with_regions_json(input);
      expect(layout).toBe(expected.layout_document_with_regions_json(input));
      expect(JSON.parse(layout).layout.pages.length).toBeGreaterThan(0);
      const expectedMeta = expected.layout_document_with_regions_retained_meta(input);
      const actualMeta = actual.layout_document_with_regions_retained_meta(input);
      try {
        expect(actualMeta).toBeInstanceOf(native.RetainedLayoutMeta);
        expect(actualMeta.page_count).toBe(expectedMeta.page_count);
        expect(actualMeta.partial).toBe(expectedMeta.partial);
        expect(actualMeta.provisional).toBe(expectedMeta.provisional);
        expect(actualMeta.notes_converged).toBe(expectedMeta.notes_converged);
        expect(actualMeta.page_sizes()).toBeInstanceOf(Float64Array);
        expect(actualMeta.page_sizes()).toEqual(expectedMeta.page_sizes());
        expect(actualMeta.layout_shell_json()).toBe(expectedMeta.layout_shell_json());
      } finally {
        actualMeta.free();
        expectedMeta.free();
      }
      expect(actual.retained_layout_json()).toBe(expected.retained_layout_json());
      expect(actual.retained_headers_footers_json()).toBe(expected.retained_headers_footers_json());
      expect(actual.retained_kernel_inputs_json()).toBe(expected.retained_kernel_inputs_json());
      const display = JSON.stringify({ fontChains });
      const frame = actual.build_display_list_frame(display, 0);
      expectPlainBytes(frame);
      expect(frame).toEqual(expected.build_display_list_frame(display, 0));
      expected.set_selection('body', paragraphId(expected), 2, paragraphId(expected), 2);
      actual.set_selection('body', paragraphId(actual), 2, paragraphId(actual), 2);
      const caret = actual.resident_caret_snapshot_json();
      expect(caret).toBe(expected.resident_caret_snapshot_json());
      const { frameEpoch, caretRect: rect } = JSON.parse(caret);
      expect(rect).not.toBeNull();
      const y = rect.y + rect.height / 2;
      expect(actual.display_hit_test_regions_json(rect.pageIndex, rect.x, y))
        .toBe(expected.display_hit_test_regions_json(rect.pageIndex, rect.x, y));
      expect(actual.display_range_rects_json(1, 4)).toBe(expected.display_range_rects_json(1, 4));
      expect(actual.display_range_rects_on_pages_json(1, 4, 0, 0))
        .toBe(expected.display_range_rects_on_pages_json(1, 4, 0, 0));
      expect(actual.display_range_rects_region_json('body', '', 1, 4))
        .toBe(expected.display_range_rects_region_json('body', '', 1, 4));
      const ranges = new Float64Array([1, 4, 2, 2]);
      const rects = actual.display_range_rects_batch('body', '', ranges);
      expect(rects).toBeInstanceOf(Float64Array);
      expect(rects).toEqual(expected.display_range_rects_batch('body', '', ranges));
      expect(actual.display_vertical_move_json(2, 'down', rect.x))
        .toBe(expected.display_vertical_move_json(2, 'down', rect.x));
      const edited = actual.apply_input('!', frameEpoch);
      expectPlainBytes(edited);
      expect(edited).toEqual(expected.apply_input('!', frameEpoch));
      expect(actual.resident_caret_snapshot_json()).toBe(expected.resident_caret_snapshot_json());
      expect(read(actual)).toEqual(read(expected));
      const nextEpoch = JSON.parse(actual.resident_caret_snapshot_json()).frameEpoch;
      const deleted = actual.apply_delete('backward', nextEpoch, 1);
      expectPlainBytes(deleted);
      expect(deleted).toEqual(expected.apply_delete('backward', nextEpoch, 1));
      expect(actual.resident_deleted_units()).toBe(expected.resident_deleted_units());
      expect(actual.resident_caret_snapshot_json()).toBe(expected.resident_caret_snapshot_json());
      const epoch = JSON.parse(actual.resident_caret_snapshot_json()).frameEpoch;
      expect(actual.build_display_pages_frame(new Uint32Array([0]), epoch))
        .toEqual(expected.build_display_pages_frame(new Uint32Array([0]), epoch));
      expectThrownParity(
        () => expected.build_display_list_frame(display, -1),
        () => actual.build_display_list_frame(display, -1),
      );
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('update callbacks deliver matching bytes synchronously and can be replaced and cleared', () => {
    const seed = new wasm.EditSession(101);
    const expected = new wasm.EditSession(102);
    const actual = new native.EditSession(102);
    const peer = new wasm.EditSession(103);
    const remote = new wasm.EditSession(104);
    const expectedUpdates: [Uint8Array, number][] = [];
    const actualUpdates: [Uint8Array, number][] = [];
    try {
      seed.seed_from_docx(fixture, 'opening');
      const state = seed.encode_state();
      for (const session of [expected, actual, peer, remote]) session.load(state);
      const observe = (updates: [Uint8Array, number][]) => (bytes: Uint8Array, isRemote: number) => {
        expectPlainBytes(bytes);
        expect([0, 1]).toContain(isRemote);
        updates.push([bytes, isRemote]);
      };
      expected.set_update_observer(observe(expectedUpdates));
      actual.set_update_observer(observe(actualUpdates));
      expected.insert_text('body', paragraphId(expected), 0, 'Local ');
      actual.insert_text('body', paragraphId(actual), 0, 'Local ');
      expect(actualUpdates.length).toBeGreaterThan(0);
      expect(actualUpdates).toEqual(expectedUpdates);
      expect(actualUpdates.every(([, origin]) => origin === 0)).toBe(true);
      const replacedCalls = actualUpdates.length;
      const replacedExpectedCalls = expectedUpdates.length;
      const replacementExpected: [Uint8Array, number][] = [];
      const replacementActual: [Uint8Array, number][] = [];
      expected.set_update_observer(observe(replacementExpected));
      actual.set_update_observer(observe(replacementActual));
      remote.insert_text('body', paragraphId(remote), 0, 'Remote ');
      const update = remote.encode_diff(seed.encode_state_vector());
      expected.apply_update(update);
      actual.apply_update(update);
      expect(actualUpdates.length).toBe(replacedCalls);
      expect(expectedUpdates.length).toBe(replacedExpectedCalls);
      expect(replacementActual).toEqual(replacementExpected);
      expect(replacementActual.some(([, origin]) => origin === 1)).toBe(true);
      for (const [bytes] of [...actualUpdates, ...replacementActual]) peer.apply_update(bytes);
      expect(read(peer)).toEqual(read(actual));
      expect(read(actual)).toEqual(read(expected));
      expected.clear_update_observer();
      actual.clear_update_observer();
      const clearedCalls = replacementActual.length;
      expected.insert_text('body', paragraphId(expected), 0, 'Cleared ');
      actual.insert_text('body', paragraphId(actual), 0, 'Cleared ');
      expect(replacementActual.length).toBe(clearedCalls);
      expect(replacementExpected.length).toBe(clearedCalls);
      let throwingCalls = 0;
      actual.set_update_observer(() => { throwingCalls += 1; throw new Error('observer'); });
      expected.insert_text('body', paragraphId(expected), 0, 'Throwing ');
      expect(() => actual.insert_text('body', paragraphId(actual), 0, 'Throwing ')).not.toThrow();
      expect(throwingCalls).toBeGreaterThan(0);
      expect(read(actual)).toEqual(read(expected));
      actual.clear_update_observer();
    } finally {
      remote.free();
      peer.free();
      actual.free();
      expected.free();
      seed.free();
    }
  });

  test('update callbacks can read versions and schedule state-vector reads', async () => {
    const seed = new wasm.EditSession(121);
    const expected = new wasm.EditSession(122);
    const actual = new native.EditSession(122);
    const remote = new wasm.EditSession(123);
    type Event = { update: Uint8Array; origin: number; version: string };
    const expectedEvents: Event[] = [];
    const actualEvents: Event[] = [];
    const expectedVectors: Uint8Array[] = [];
    const actualVectors: Uint8Array[] = [];
    try {
      seed.seed_from_docx(fixture, 'opening');
      const state = seed.encode_state();
      for (const session of [expected, actual, remote]) session.load(state);
      const expectedBefore = expected.version();
      const actualBefore = actual.version();
      const observe = (session: Session, events: Event[], vectors: Uint8Array[]) =>
        (update: Uint8Array, origin: number) => {
          events.push({ update, origin, version: session.version() });
          queueMicrotask(() => vectors.push(session.encode_state_vector()));
        };
      expected.set_update_observer(observe(expected, expectedEvents, expectedVectors));
      actual.set_update_observer(observe(actual, actualEvents, actualVectors));
      expected.insert_text('body', paragraphId(expected), 0, 'Local ');
      actual.insert_text('body', paragraphId(actual), 0, 'Local ');
      expect(expectedEvents).toHaveLength(1);
      expect(actualEvents).toHaveLength(1);
      expect(expectedEvents[0].version).toBe(expected.version());
      expect(actualEvents[0].version).toBe(actual.version());
      expect(expectedEvents[0].version).not.toBe(expectedBefore);
      expect(actualEvents[0].version).not.toBe(actualBefore);
      expect(actualEvents[0].update).toEqual(expectedEvents[0].update);
      expect(actualEvents[0].origin).toBe(0);
      expect(expectedEvents[0].origin).toBe(0);
      expect(actualEvents[0].version.slice(17)).toBe(expectedEvents[0].version.slice(17));
      await Promise.resolve();
      expect(actualVectors).toHaveLength(1);
      expect(actualVectors).toEqual(expectedVectors);
      expect(actualVectors[0]).toEqual(actual.encode_state_vector());
      expect(expectedVectors[0]).toEqual(expected.encode_state_vector());
      const expectedLocalVersion = expected.version();
      const actualLocalVersion = actual.version();
      remote.insert_text('body', paragraphId(remote), 0, 'Remote ');
      const update = remote.encode_diff(seed.encode_state_vector());
      expected.apply_update(update);
      actual.apply_update(update);
      expect(actualEvents.length).toBeGreaterThan(1);
      expect(actualEvents.length).toBe(expectedEvents.length);
      expect(expectedEvents.at(-1)?.version).toBe(expected.version());
      expect(actualEvents.at(-1)?.version).toBe(actual.version());
      expect(expected.version()).not.toBe(expectedLocalVersion);
      expect(actual.version()).not.toBe(actualLocalVersion);
      const normalize = (events: Event[], initial: string) => events.map((event) => {
        expect(event.version.slice(0, 16)).toBe(initial.slice(0, 16));
        return { ...event, version: event.version.slice(17) };
      });
      expect(actualEvents[0].update).toEqual(expectedEvents[0].update);
      expect(actualEvents.map(({ origin }) => origin)).toEqual(expectedEvents.map(({ origin }) => origin));
      const replay = (events: Event[]) => {
        const peer = new wasm.EditSession(124);
        try {
          peer.load(state);
          for (const { update } of events) peer.apply_update(update);
          return read(peer);
        } finally {
          peer.free();
        }
      };
      expect(replay(actualEvents)).toEqual(replay(expectedEvents));
      for (const event of actualEvents.slice(1)) expect(event.version).toBe(actual.version());
      expect(normalize(actualEvents.slice(-1), actualBefore))
        .toEqual(normalize(expectedEvents.slice(-1), expectedBefore));
      expect(actualEvents.slice(1).some((event) => event.origin === 1)).toBe(true);
      await Promise.resolve();
      expect(actualVectors.length).toBe(actualEvents.length);
      expect(actualVectors).toEqual(expectedVectors);
      expect(actualVectors.at(-1)).toEqual(actual.encode_state_vector());
      expect(expectedVectors.at(-1)).toEqual(expected.encode_state_vector());
      expect(read(actual)).toEqual(read(expected));
    } finally {
      remote.free();
      actual.free();
      expected.free();
      seed.free();
    }
  });

  test('native callbacks read committed session state synchronously', () => {
    const session = new native.EditSession(112);
    const remote = new wasm.EditSession(113);
    const snapshot = () => ({
      stateVector: session.encode_state_vector(),
      paragraphs: session.read_paragraphs_json(request),
      changedStories: session.stories_changed_since(0),
      version: session.version(),
      clientId: session.client_id(),
      directBatches: session.direct_batches_applied(),
    });
    const snapshots: ReturnType<typeof snapshot>[] = [];
    const drained: Uint8Array[] = [];
    try {
      session.seed_from_docx(fixture, 'opening');
      remote.load(session.encode_state());
      session.start_update_event_observation();
      session.set_update_observer(() => {
        snapshots.push(snapshot());
        drained.push(session.drain_update_event());
      });
      session.insert_text('body', paragraphId(session), 0, 'Local ');
      expect(snapshots).toEqual([snapshot()]);
      const localCalls = snapshots.length;
      remote.insert_text('body', paragraphId(remote), 0, 'Remote ');
      session.apply_update(remote.encode_diff(session.encode_state_vector()));
      expect(snapshots.length).toBeGreaterThan(localCalls);
      for (const observed of snapshots.slice(localCalls)) expect(observed).toEqual(snapshot());
      expect(drained).toHaveLength(snapshots.length);
      for (const update of drained) {
        expectPlainBytes(update);
        expect(update.byteLength).toBeGreaterThan(1);
      }
      expect(drained[0][0]).toBe(0);
      expect(drained.slice(localCalls).some((update) => update[0] === 1)).toBe(true);
      expect(session.drain_update_event()).toEqual(new Uint8Array());
    } finally {
      remote.free();
      session.free();
    }
  });

  test('native callbacks queue edits until the current callback returns', () => {
    const session = new native.EditSession(114);
    const peer = new wasm.EditSession(115);
    const updates: [Uint8Array, number][] = [];
    const calls: string[] = [];
    try {
      session.seed_from_docx(fixture, 'opening');
      peer.load(session.encode_state());
      const paraId = paragraphId(session);
      session.set_update_observer((update: Uint8Array, origin: number) => {
        updates.push([update, origin]);
        if (updates.length === 1) {
          calls.push('before edit');
          session.insert_text('body', paraId, 0, 'Nested ');
          calls.push('after edit');
        }
        calls.push(`callback ${updates.length}`);
      });
      session.insert_text('body', paraId, 0, 'Outer ');
      expect(calls).toEqual(['before edit', 'after edit', 'callback 1', 'callback 2']);
      expect(updates).toHaveLength(2);
      for (const [update, origin] of updates) {
        expectPlainBytes(update);
        expect(origin).toBe(0);
        peer.apply_update(update);
      }
      expect(read(session).paragraphs[0].text).toBe('Nested Outer Engine text');
      expect(read(peer)).toEqual(read(session));
    } finally {
      peer.free();
      session.free();
    }
  });

  test('clearing a native observer inside its callback stops queued delivery', () => {
    const session = new native.EditSession(116);
    let calls = 0;
    try {
      session.set_update_observer(() => {
        calls += 1;
        session.clear_update_observer();
      });
      session.load_json(JSON.stringify([
        { storyId: 'body', paragraphs: [{ text: 'First' }] },
        { storyId: 'other', paragraphs: [{ text: 'Second' }] },
      ]));
      expect(calls).toBe(1);
      expect(session.story_ids()).toEqual(['body', 'other']);
      session.insert_text('body', paragraphId(session), 0, 'After ');
      expect(calls).toBe(1);
    } finally {
      session.free();
    }
  });

  test('freeing a native session inside its callback stops queued delivery', () => {
    const session = new native.EditSession(117);
    let calls = 0;
    try {
      session.set_update_observer(() => {
        calls += 1;
        session.free();
      });
      expect(() => session.load_json(JSON.stringify([
        { storyId: 'body', paragraphs: [{ text: 'First' }] },
        { storyId: 'other', paragraphs: [{ text: 'Second' }] },
      ]))).not.toThrow();
      expect(calls).toBe(1);
      expect(() => session.version()).toThrow('EditSession was freed');
      expect(() => session.encode_state_vector()).toThrow('EditSession was freed');
      expect(() => session.read_paragraphs_json(request)).toThrow('EditSession was freed');
    } finally {
      session.free();
    }
  });

  test('a failing call preserves its own exception while delivering native callbacks', () => {
    const expected = new wasm.EditSession(118);
    const actual = new native.EditSession(118);
    const expectedUpdates: [Uint8Array, number][] = [];
    const actualUpdates: [Uint8Array, number][] = [];
    const vectors: Uint8Array[] = [];
    const replayed = (updates: [Uint8Array, number][]) => {
      const peer = new wasm.EditSession(119);
      try {
        for (const [update] of updates) peer.apply_update(update);
        return read(peer);
      } finally {
        peer.free();
      }
    };
    const origins = (updates: [Uint8Array, number][]) => updates.map(([, origin]) => origin);
    try {
      expected.set_update_observer((update: Uint8Array, origin: number) => {
        expectedUpdates.push([update, origin]);
        throw new Error('observer failure');
      });
      actual.set_update_observer((update: Uint8Array, origin: number) => {
        actualUpdates.push([update, origin]);
        vectors.push(actual.encode_state_vector());
        actual.set_undo_capture_mode('invalid');
      });
      const input = JSON.stringify([
        { storyId: 'body', paragraphs: [{ text: 'Committed' }] },
        {},
      ]);
      expect(expectThrownParity(() => expected.load_json(input), () => actual.load_json(input)))
        .toBe('a story entry requires a string "storyId"');
      expect(actualUpdates.length).toBeGreaterThan(0);
      expect(origins(actualUpdates)).toEqual(origins(expectedUpdates));
      expect(replayed(actualUpdates)).toEqual(replayed(expectedUpdates));
      expect(vectors).toHaveLength(actualUpdates.length);
      for (const vector of vectors) expect(vector).toEqual(actual.encode_state_vector());
      expect(read(actual)).toEqual(read(expected));
      const before = actualUpdates.length;
      expected.insert_text('body', paragraphId(expected), 0, 'After ');
      expect(() => actual.insert_text('body', paragraphId(actual), 0, 'After ')).not.toThrow();
      expect(actualUpdates.length).toBeGreaterThan(before);
      expect(origins(actualUpdates)).toEqual(origins(expectedUpdates));
      expect(replayed(actualUpdates)).toEqual(replayed(expectedUpdates));
      expect(read(actual)).toEqual(read(expected));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('an unreadable relationships part warns the same way in both engines', () => {
    const parts = wasm.unzip_docx(fixture) as Record<string, Uint8Array>;
    const damaged = wasm.rezip_docx({ ...parts, 'word/_rels/header1.xml.rels': new TextEncoder().encode('<Relationships><not closed') });
    const output = native.parse_docx_s4(damaged);
    expect(output).toBe(wasm.parse_docx_s4(damaged));
    expect((JSON.parse(output) as { projection: { warnings?: string[] } }).projection.warnings?.[0]).toContain('word/_rels/header1.xml.rels');
    expect((JSON.parse(native.parse_docx_s4(fixture)) as { projection: Record<string, unknown> }).projection).not.toHaveProperty('warnings');
  });

  test('free functions preserve ZIP shapes, parse JSON, exports and sanitization', () => {
    const expected = wasm.unzip_docx(fixture) as Record<string, Uint8Array>;
    const actual = native.unzip_docx(fixture) as Record<string, Uint8Array>;
    expect(Object.getPrototypeOf(actual)).toBe(Object.prototype);
    expect(Object.keys(actual)).toEqual(Object.keys(expected));
    for (const [name, bytes] of Object.entries(actual)) {
      expectPlainBytes(bytes);
      expect(bytes).toEqual(expected[name]);
    }
    const zipped = native.rezip_docx(actual);
    expectPlainBytes(zipped);
    expect(zipped).toEqual(wasm.rezip_docx(expected));
    expect(native.unzip_docx(zipped)).toEqual(expected);
    expect(native.parse_docx_s2(fixture)).toBe(wasm.parse_docx_s2(fixture));
    expect(native.export_docx_structured_json(fixture, '{"revisionView":"accepted"}'))
      .toBe(wasm.export_docx_structured_json(fixture, '{"revisionView":"accepted"}'));
    expect(native.paragraphPatchWriterVersion()).toBe(wasm.paragraphPatchWriterVersion());
    const sanitized = native.sanitizeOoxml(fixture, 'docx');
    expectPlainBytes(sanitized);
    expect(sanitized).toEqual(wasm.sanitizeOoxml(fixture, 'docx'));
    expectThrownParity(
      () => wasm.parse_docx_s2(new Uint8Array([1, 2, 3])),
      () => native.parse_docx_s2(new Uint8Array([1, 2, 3])),
    );
    for (const value of [null, undefined, 1, 'entries']) {
      expect(expectThrownParity(() => wasm.rezip_docx(value), () => native.rezip_docx(value)))
        .toBe('rezip_docx: expected an object');
    }
    const entries = Object.assign(Object.create({ ignored: new Uint8Array([9]) }), {
      data: [1, 2, 255],
      [Symbol('ignored')]: new Uint8Array([8]),
    });
    expect(native.unzip_docx(native.rezip_docx(entries)))
      .toEqual(wasm.unzip_docx(wasm.rezip_docx(entries)));
    expect(native.wasm_heap_counted()).toBe(false);
    native.set_wasm_heap_limit(1);
    expect(native.wasm_live_bytes()).toBe(0);
    expect(native.wasm_peak_bytes()).toBe(0);
    expect(native.wasm_failed_allocation_bytes()).toBe(0);
    expect(native.reset_wasm_peak_bytes()).toBeUndefined();
    expect(native.install_panic_hook()).toBeUndefined();
  });

  test('module measurement fonts share the free glyph API', () => {
    wasm.clear_measure_fonts();
    native.clear_measure_fonts();
    try {
      const font = wasm.register_measure_font(fontBytes);
      expect(native.register_measure_font(fontBytes)).toBe(font);
      expect(native.register_substitute_measure_font(font, 'Calibri'))
        .toBe(wasm.register_substitute_measure_font(font, 'Calibri'));
      expect(native.outline_glyph_json(font, 36)).toBe(wasm.outline_glyph_json(font, 36));
      expectThrownParity(
        () => wasm.register_measure_font(new Uint8Array([1])),
        () => native.register_measure_font(new Uint8Array([1])),
      );
    } finally {
      native.clear_measure_fonts();
      wasm.clear_measure_fonts();
    }
  });

  test.each([-1, 0.5, Number.NaN, Number.POSITIVE_INFINITY, 9_007_199_254_740_992])(
    'constructor rejects invalid client id %s with a string', (clientId) => {
      expect(expectThrownParity(
        () => new wasm.EditSession(clientId),
        () => new native.EditSession(clientId),
      )).toBe('client_id must be a non-negative safe integer');
    },
  );

  test('message strings and PeerMetadataError fields match wasm', () => {
    const expected = new wasm.EditSession(51);
    const actual = new native.EditSession(51);
    try {
      expect(expectThrownParity(
        () => expected.set_undo_capture_mode('invalid'),
        () => actual.set_undo_capture_mode('invalid'),
      )).toBe('undo capture mode must be auto or manual');
      const peerError = expectThrownParity(
        () => expected.encode_peer_metadata(),
        () => actual.encode_peer_metadata(),
      ) as Error & { code: string };
      expect(peerError.name).toBe('PeerMetadataError');
      expect(peerError.code).toBe('missing-source');
      expectThrownParity(() => expected.run_inheritance('{'), () => actual.run_inheritance('{'));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('optional returns use undefined', () => {
    const session = new native.EditSession(61);
    try {
      expect(session.materialize_docx()).toBeUndefined();
      expect(session.media_bytes('missing')).toBeUndefined();
      expect(session.media_type('missing')).toBeUndefined();
      expect(session.media_data_url('missing')).toBeUndefined();
      expect(session.clear_cell_selection()).toBeUndefined();
    } finally {
      session.free();
    }
  });

  test('prime_save_seed_states caches within the same budgets as wasm', () => {
    const expected = new wasm.EditSession(62);
    const actual = new native.EditSession(62);
    try {
      expected.seed_from_docx(fixture, 'opening');
      actual.seed_from_docx(fixture, 'opening');
      for (const [paragraphs, inflated] of [[0, 0], [1, 0], [1_000, 64 * 1024 * 1024]]) {
        expect(actual.prime_save_seed_states(paragraphs, inflated))
          .toBe(expected.prime_save_seed_states(paragraphs, inflated));
      }
      expect(read(actual)).toEqual(read(expected));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('prime_save_seed_states_slice answers each slice as wasm does', () => {
    const expected = new wasm.EditSession(63);
    const actual = new native.EditSession(63);
    try {
      expected.seed_from_docx(fixture, 'opening');
      actual.seed_from_docx(fixture, 'opening');
      for (const [paragraphs, inflated, bytes] of [[0, 0, 1], [1_000, 64 * 1024 * 1024, 1], [1_000, 64 * 1024 * 1024, 1]]) {
        expect(actual.prime_save_seed_states_slice(paragraphs, inflated, bytes))
          .toBe(expected.prime_save_seed_states_slice(paragraphs, inflated, bytes));
      }
      let answer = 1;
      for (let slice = 0; slice < 10_000 && answer === 1; slice += 1) {
        answer = expected.prime_save_seed_states_slice(1_000, 64 * 1024 * 1024, 64);
        expect(actual.prime_save_seed_states_slice(1_000, 64 * 1024 * 1024, 64)).toBe(answer);
      }
      expect(read(actual)).toEqual(read(expected));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('RunInheritance resolves the same attributes and formatting', () => {
    const expectedSession = new wasm.EditSession(71);
    const actualSession = new native.EditSession(71);
    const styles = JSON.stringify({
      docDefaults: { rPr: { fontSize: 22 } },
      styles: [
        { type: 'paragraph', styleId: 'Normal', default: true },
        { type: 'paragraph', styleId: 'Heading', rPr: { bold: true }, pPr: { spaceAfter: 80 } },
        { type: 'character', styleId: 'Emphasis', rPr: { italic: true } },
      ],
    });
    const expected = expectedSession.run_inheritance(styles);
    const actual = actualSession.run_inheritance(styles);
    try {
      expect(JSON.parse(actual.attributes())).toEqual(JSON.parse(expected.attributes()));
      expect(JSON.parse(actual.attributes('Heading', 'Emphasis')))
        .toEqual(JSON.parse(expected.attributes('Heading', 'Emphasis')));
      expect(JSON.parse(actual.attributes(null, null, '{"bold":true}')))
        .toEqual(JSON.parse(expected.attributes(null, null, '{"bold":true}')));
      expect(JSON.parse(actual.paragraph_formatting(undefined, 'null')))
        .toEqual(JSON.parse(expected.paragraph_formatting(undefined, 'null')));
      expect(JSON.parse(actual.paragraph_formatting('Heading', '{}')))
        .toEqual(JSON.parse(expected.paragraph_formatting('Heading', '{}')));
      const table = JSON.stringify({ type: 'table', rows: [{ cells: [{ type: 'tableCell', content: [] }] }] });
      expect(JSON.parse(actual.table_attributes(table))).toEqual(JSON.parse(expected.table_attributes(table)));
      expectThrownParity(() => expected.attributes(null, null, '{'), () => actual.attributes(null, null, '{'));
      expectThrownParity(() => expected.paragraph_formatting(null, '{'), () => actual.paragraph_formatting(null, '{'));
      expectThrownParity(() => expected.table_attributes('{'), () => actual.table_attributes('{'));
    } finally {
      actual.free();
      expected.free();
      actualSession.free();
      expectedSession.free();
    }
  });

  test('story_checksum of a million-character formatted run matches and stays within wasm memory', () => {
    const encoder = new TextEncoder();
    const text = 'lorem ipsum '.repeat(83_334).slice(0, 1_000_000);
    const rPr = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial"/><w:b/><w:i/><w:color w:val="1F3864"/><w:sz w:val="22"/><w:u w:val="single"/><w:lang w:val="en-GB"/><w:highlight w:val="yellow"/></w:rPr>';
    const bytes = wasm.rezip_docx({
      '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
      '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
      'word/document.xml': encoder.encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
    });
    const expected = new wasm.EditSession(7);
    const actual = new native.EditSession(7);
    try {
      expected.open_docx(bytes, true);
      actual.open_docx(bytes, true);
      const checksum = expected.story_checksum('body');
      expect(checksum).toMatch(/^\d+$/);
      expect(actual.story_checksum('body')).toBe(checksum);
    } finally {
      actual.free();
      expected.free();
    }
  }, 120_000);

  test('decodeFontFace and font_preflight failures have the wasm error shape', () => {
    const nativePreflight = (native as unknown as { font_preflight: typeof wasmFontPreflight }).font_preflight;
    const invalid = new Uint8Array([1, 2, 3, 4]);
    for (const [codec, dict, maxOutput] of [['woff2', undefined, undefined], ['unknown', undefined, undefined], ['', undefined, undefined]] as const) {
      const actual = expectThrownParity(
        () => wasm.decodeFontFace(invalid, codec, dict, maxOutput),
        () => native.decodeFontFace(invalid, codec, dict, maxOutput),
      );
      expect(actual instanceof Error && 'code' in actual).toBe(false);
    }
    expectThrownParity(() => wasmFontPreflight(invalid), () => nativePreflight(invalid));
  });
});

test('native and wasm track and resolve character formatting with the same optional author arguments', () => {
  for (const engine of [wasm, native]) {
    for (const accept of [true, false]) {
      const session = new engine.EditSession(9303);
      try {
        session.open_docx(fixture, true, 'format-tracking');
        const id = paragraphId(session);
        session.toggle_mark('body', id, 0, id, 6, '{"type":"bold"}', 'Writer', '2026-10-01T12:00:00Z');
        session.format_range('body', id, 1, id, 5, '{"italic":true}', 'Writer', '2026-10-01T12:00:00Z');
        session.clear_formatting('body', id, 2, id, 4, 'Writer', '2026-10-01T12:00:00Z');
        const revisions = JSON.parse(session.list_revisions());
        expect(revisions).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'rPrChange', author: 'Writer' })]));
        for (const revisionId of new Set(revisions.map((revision: { revisionId: string }) => revision.revisionId))) {
          const target = JSON.stringify({ revisionId });
          if (accept) session.accept_change(target);
          else session.reject_change(target);
        }
        expect(JSON.parse(session.list_revisions())).toEqual([]);
        const context = JSON.parse(session.selection_context('body', id, 0, id, 1));
        expect(context.bold).toBe(accept);
        session.set_hyperlink('body', id, 0, id, 6, '{"href":"#target"}', 'Writer', '2026-10-01T12:00:00Z');
        expect(JSON.parse(session.list_revisions()).length).toBeGreaterThan(0);
      } finally {
        session.free();
      }
    }
  }
});

test('native and wasm mark the same table row headers in the display list', () => {
  const encoder = new TextEncoder();
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const row = (texts: string[], header = false) => `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${texts.map(cell).join('')}</w:tr>`;
  const bytes = wasm.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': encoder.encode(`<w:document xmlns:w="${w}"><w:body><w:tbl><w:tblPr><w:tblLook w:val="04A0" w:firstRow="1" w:firstColumn="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>${row(['Region', 'Q1'], true)}${row(['North', '12'])}</w:tbl><w:p/><w:sectPr/></w:body></w:document>`),
  });
  type Cell = { row: number; col: number; isHeader?: boolean; isRowHeader?: boolean; headerIds?: string[] };
  const run = (engine: Engine) => {
    const session = new engine.EditSession(77);
    try {
      session.open_docx(bytes, true);
      const font = session.register_measure_font(fontBytes);
      const fontChains = { 'calibri|0|0': [font], 'calibri|1|0': [font] };
      const layout = session.layout_document_with_regions_json(JSON.stringify({
        bodyStory: 'body',
        regions: { sections: [{ sectionId: 'main', properties: {
          pageWidth: 12240, pageHeight: 15840, marginTop: 1440, marginRight: 1440, marginBottom: 1440, marginLeft: 1440,
        } }] },
        measurement: { fontChains, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
        renderEnv: { themeColors: (JSON.parse(session.materialize_docx() ?? '{}') as {
          document?: { package?: { theme?: { colorScheme?: Record<string, string> } } };
        }).document?.package?.theme?.colorScheme ?? {} },
      }));
      const list = JSON.parse(session.build_display_list_json(JSON.stringify({ ...JSON.parse(layout), fontChains }))) as {
        pages: { primitives: { kind: string; cell?: Cell }[] }[];
      };
      const cells = new Map<string, Omit<Cell, 'headerIds'> & { headerIds: number }>();
      for (const primitive of list.pages.flatMap((page) => page.primitives)) {
        const ref = primitive.cell;
        if (!ref || cells.has(`${ref.row}:${ref.col}`)) continue;
        cells.set(`${ref.row}:${ref.col}`, {
          row: ref.row, col: ref.col, isHeader: ref.isHeader, isRowHeader: ref.isRowHeader, headerIds: ref.headerIds?.length ?? 0,
        });
      }
      return [...cells.values()].sort((a, b) => a.row - b.row || a.col - b.col);
    } finally {
      session.free();
    }
  };
  const nativeCells = run(native);
  expect(nativeCells).toEqual([
    { row: 0, col: 0, isHeader: true, isRowHeader: undefined, headerIds: 0 },
    { row: 0, col: 1, isHeader: true, isRowHeader: undefined, headerIds: 0 },
    { row: 1, col: 0, isHeader: undefined, isRowHeader: true, headerIds: 0 },
    { row: 1, col: 1, isHeader: undefined, isRowHeader: undefined, headerIds: 2 },
  ]);
  expect(nativeCells).toEqual(run(wasm));
});

test('native and wasm read the same page text, cutting paragraphs and tables where the page does', () => {
  const encoder = new TextEncoder();
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const long = Array.from({ length: 3000 }, (_, index) => `w${index}`).join(' ');
  const tall = Array.from({ length: 1500 }, (_, index) => `t${index}`).join(' ');
  const row = (index: number) => `<w:tr>${index === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}<w:tc><w:p><w:r><w:t>R${index}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>V${index}</w:t></w:r></w:p></w:tc></w:tr>`;
  const field = (code: string) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const body = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Chapter one</w:t></w:r></w:p>'
    + `<w:p><w:r><w:t>${long}</w:t></w:r></w:p>`
    + `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${Array.from({ length: 60 }, (_, index) => row(index)).join('')}</w:tbl>`
    + `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>${tall}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`
    + `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${[
      '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged across</w:t></w:r></w:p></w:tc></w:tr>',
      '<w:tr><w:trPr><w:hidden/></w:trPr><w:tc><w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden row</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden cell</w:t></w:r></w:p></w:tc></w:tr>',
      '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Merged down</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Side one</w:t></w:r></w:p></w:tc></w:tr>',
      '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>Side two</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr>',
    ].join('')}</w:tbl>`
    + '<w:p><w:r><w:t>Visible</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hiddenword</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>Kept</w:t></w:r><w:del w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>Struck</w:delText></w:r></w:del></w:p>'
    + '<w:p><w:r><w:t>Note here</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>'
    + `<w:p><w:r><w:t xml:space="preserve">Of </w:t></w:r>${field('NUMPAGES')}</w:p>`
    + '<w:p><w:r><w:t>After the table</w:t></w:r></w:p>'
    + '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    + '<w:p w14:paraId="0A0B0C0D"><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Section two</w:t></w:r></w:p>';
  const bytes = wasm.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'),
    'word/_rels/document.xml.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/footnotes" Target="footnotes.xml"/></Relationships>`),
    'word/footnotes.xml': encoder.encode(`<w:footnotes xmlns:w="${w}"><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>Note text</w:t></w:r></w:p></w:footnote></w:footnotes>`),
    '_rels/.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/document.xml': encoder.encode(`<w:document xmlns:w="${w}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr/></w:body></w:document>`),
  });
  type PageText = { pageIndex: number; paragraphs: { text: string; level?: number; inTable?: boolean }[] };
  const run = (engine: Engine, columns: object = {}, paint = false) => {
    const session = new engine.EditSession(79);
    try {
      session.open_docx(bytes, true);
      const font = session.register_measure_font(fontBytes);
      const fontChains = { 'calibri|0|0': [font], 'calibri|1|0': [font] };
      const layout = session.layout_document_with_regions_json(JSON.stringify({
        bodyStory: 'body',
        regions: { sections: [{ sectionId: 'main', properties: {
          pageWidth: 12240, pageHeight: 15840, marginTop: 1440, marginRight: 1440, marginBottom: 1440, marginLeft: 1440, ...columns,
        } }] },
        measurement: { fontChains, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
        renderEnv: { themeColors: {} },
      }));
      const list = paint ? JSON.parse(session.build_display_list_json(JSON.stringify({ ...JSON.parse(layout), fontChains }))) as {
        pages: { primitives: { text?: string }[] }[];
      } : { pages: [] };
      return {
        pages: JSON.parse(session.page_texts_json(Uint32Array.from({ length: 200 }, (_, index) => index))) as PageText[],
        painted: list.pages.map((page) => page.primitives.map(({ text }) => text ?? '').join('').replace(/\s+/g, '')),
      };
    } finally {
      session.free();
    }
  };
  const { pages, painted } = run(native, {}, true);
  expect(pages.length).toBeGreaterThanOrEqual(3);
  expect(pages.map(({ pageIndex }) => pageIndex)).toEqual(pages.map((_, index) => index));
  expect(pages[0]!.paragraphs[0]).toEqual({ text: 'Chapter one', level: 0 });
  const plain = pages.flatMap(({ paragraphs }) => paragraphs.filter(({ inTable, level }) => !inTable && level === undefined));
  const pieces = plain.filter(({ text }) => text.startsWith('w') || /\bw\d+/.test(text)).map(({ text }) => text);
  expect(pieces.length).toBeGreaterThanOrEqual(2);
  expect(pieces.join('').replace(/\s+/g, ' ').trim()).toBe(long);
  const cells = pages.flatMap(({ pageIndex, paragraphs }) => paragraphs.filter(({ inTable }) => inTable).map(({ text }) => [pageIndex, text] as const));
  const expectedCells = [...Array.from({ length: 60 }, (_, index) => [`R${index}`, `V${index}`]).flat(), tall, 'Merged across', 'Merged down', 'Side one', 'Side two', 'Nested'];
  expect(cells.map(([, text]) => text)).toEqual(expectedCells);
  expect(pages.flatMap(({ paragraphs }) => paragraphs).filter(({ text }) => /Hidden (row|cell)/.test(text))).toEqual([]);
  expect(new Set(cells.map(([pageIndex]) => pageIndex)).size).toBeGreaterThanOrEqual(2);
  const pageField = pages.flatMap(({ pageIndex, paragraphs }) => paragraphs.filter(({ text }) => /^\d+$/.test(text)).map(({ text }) => [pageIndex, text]));
  expect(pageField).toEqual([[pageField[0]![0], String(Number(pageField[0]![0]) + 1)]]);
  expect(Number(pageField[0]![0])).toBeGreaterThan(0);
  expect(plain.filter(({ text }) => !/^\d+$/.test(text)).at(-1)?.text).toBe('After the table');
  expect(pages.at(-1)!.paragraphs.at(-1)).toEqual({ text: 'Section two', level: 1 });
  expect(run(wasm).pages).toEqual(pages);
  for (const { pageIndex, paragraphs } of pages) {
    for (const { text, inTable } of paragraphs) {
      if (!inTable) expect(painted[pageIndex]).toContain(text.replace(/\s+/g, ''));
    }
  }
  const everything = pages.flatMap(({ paragraphs }) => paragraphs.map(({ text }) => text)).join('\n');
  expect(everything).not.toContain('Hiddenword');
  expect(painted.join('')).not.toContain('Hiddenword');
  expect(everything.includes('Struck')).toBe(painted.join('').includes('Struck'));
  expect(everything).toContain(`Of ${pages.length}`);

  const twoColumns = { columnCount: 2, columnSpace: 720 };
  const columned = run(native, twoColumns).pages;
  const columnedCells = columned.flatMap(({ paragraphs }) => paragraphs.filter(({ inTable }) => inTable).map(({ text }) => text));
  expect(columnedCells).toEqual(expectedCells);
  expect(run(wasm, twoColumns).pages).toEqual(columned);
}, 30_000);

test('native and wasm mark the same list markers as numbered', () => {
  const encoder = new TextEncoder();
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const level = (format: string, text: string) =>
    `<w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/></w:lvl>`;
  const item = (num: number, text: string) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${num}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const bytes = wasm.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'),
    '_rels/.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/_rels/document.xml.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/numbering" Target="numbering.xml"/></Relationships>`),
    'word/numbering.xml': encoder.encode(`<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0">${level('decimal', 'Part %1.')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${level('bullet', 'q')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`),
    'word/document.xml': encoder.encode(`<w:document xmlns:w="${w}"><w:body>${item(1, 'Alpha')}${item(2, 'Beta')}<w:sectPr/></w:body></w:document>`),
  });
  const run = (engine: Engine) => {
    const session = new engine.EditSession(79);
    try {
      session.open_docx(bytes, true);
      const font = session.register_measure_font(fontBytes);
      const fontChains = { 'calibri|0|0': [font], 'calibri|1|0': [font] };
      const layout = session.layout_document_with_regions_json(JSON.stringify({
        bodyStory: 'body',
        regions: { sections: [{ sectionId: 'main', properties: {
          pageWidth: 12240, pageHeight: 15840, marginTop: 1440, marginRight: 1440, marginBottom: 1440, marginLeft: 1440,
        } }] },
        measurement: { fontChains, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
        renderEnv: { themeColors: (JSON.parse(session.materialize_docx() ?? '{}') as {
          document?: { package?: { theme?: { colorScheme?: Record<string, string> } } };
        }).document?.package?.theme?.colorScheme ?? {} },
      }));
      const list = JSON.parse(session.build_display_list_json(JSON.stringify({ ...JSON.parse(layout), fontChains }))) as {
        pages: { primitives: { listMarker?: boolean; listMarkerNumbered?: boolean; listMarkerBullet?: boolean; text?: string }[] }[];
      };
      const markers = list.pages.flatMap((page) => page.primitives).filter((primitive) => primitive.listMarker);
      return {
        numbered: markers.filter((marker) => marker.listMarkerNumbered).map((marker) => marker.text).join(''),
        bullets: markers.filter((marker) => marker.listMarkerBullet).map((marker) => marker.text).join(''),
        neither: markers.filter((marker) => !marker.listMarkerNumbered && !marker.listMarkerBullet).length,
      };
    } finally {
      session.free();
    }
  };
  const markers = run(native);
  expect(markers.numbered.trim()).toBe('Part 3.');
  expect(markers.bullets.trim()).not.toBe('');
  expect(markers.neither).toBe(0);
  expect(run(wasm)).toEqual(markers);
});

test('broken packages open or fail the same on native and wasm, and a failure reads as damaged', () => {
  const encoder = new TextEncoder();
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const types = encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  const rels = encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`);
  const document = encoder.encode(`<w:document xmlns:w="${w}"><w:body><w:p/><w:sectPr/></w:body></w:document>`);
  const good = { '[Content_Types].xml': types, '_rels/.rels': rels, 'word/document.xml': document };
  const without = (name: keyof typeof good) => Object.fromEntries(Object.entries(good).filter(([part]) => part !== name));
  const zipped = wasm.rezip_docx(good);
  const cases: [string, Uint8Array][] = [
    ['not a zip', encoder.encode('this is not a zip file')],
    ['a truncated zip', zipped.slice(0, zipped.length >> 1)],
    ['no content types', wasm.rezip_docx(without('[Content_Types].xml'))],
    ['no package relationships', wasm.rezip_docx(without('_rels/.rels'))],
    ['no main part', wasm.rezip_docx(without('word/document.xml'))],
    ['a main part that is not XML', wasm.rezip_docx({ ...good, 'word/document.xml': encoder.encode('<w:document') })],
    ['a main part of another kind', wasm.rezip_docx({ ...good, 'word/document.xml': encoder.encode('<other/>') })],
  ];
  const failure = (engine: Engine, bytes: Uint8Array) => {
    const session = new engine.EditSession(80);
    try {
      session.open_docx(bytes, true);
      return null;
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    } finally {
      session.free();
    }
  };
  const outcomes = cases.map(([name, bytes]) => {
    const message = failure(native, bytes);
    return {
      name, sameOnWasm: failure(wasm, bytes) === message,
      kind: message === null ? 'opens' : openFailureKind(message),
    };
  });
  expect(outcomes).toEqual(cases.map(([name]) => ({
    name, sameOnWasm: true, kind: name === 'no content types' || name === 'no package relationships' ? 'opens' : 'damaged',
  })));
});

test('native and wasm carry the same chart data in the display list', () => {
  const encoder = new TextEncoder();
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const drawing = '<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    + '<wp:inline><wp:extent cx="2743200" cy="1828800"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    + '<c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>';
  const chart = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>'
    + '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx>'
    + '<c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat>'
    + '<c:val><c:numRef><c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>'
    + '</c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>'
    + '<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>';
  const bytes = wasm.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'),
    '_rels/.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/_rels/document.xml.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="${rel}/chart" Target="charts/chart1.xml"/></Relationships>`),
    'word/document.xml': encoder.encode(`<w:document xmlns:w="${w}" xmlns:r="${rel}"><w:body><w:p><w:r>${drawing}</w:r></w:p><w:sectPr/></w:body></w:document>`),
    'word/charts/chart1.xml': encoder.encode(chart),
  });
  const run = (engine: Engine) => {
    const session = new engine.EditSession(78);
    try {
      session.open_docx(bytes, true);
      const font = session.register_measure_font(fontBytes);
      const fontChains = { 'calibri|0|0': [font], 'calibri|1|0': [font] };
      const layout = session.layout_document_with_regions_json(JSON.stringify({
        bodyStory: 'body',
        regions: { sections: [{ sectionId: 'main', properties: {
          pageWidth: 12240, pageHeight: 15840, marginTop: 1440, marginRight: 1440, marginBottom: 1440, marginLeft: 1440,
        } }] },
        measurement: { fontChains, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
        renderEnv: { themeColors: (JSON.parse(session.materialize_docx() ?? '{}') as {
          document?: { package?: { theme?: { colorScheme?: Record<string, string> } } };
        }).document?.package?.theme?.colorScheme ?? {} },
      }));
      const list = JSON.parse(session.build_display_list_json(JSON.stringify({ ...JSON.parse(layout), fontChains }))) as {
        pages: { primitives: { chart?: { label: string; data?: unknown } }[] }[];
      };
      const charted = list.pages.flatMap((page) => page.primitives).filter((primitive) => primitive.chart);
      expect(charted.length).toBeGreaterThan(1);
      expect(charted.filter((primitive) => primitive.chart!.data !== undefined)).toHaveLength(1);
      expect(charted[0]!.chart!.data).toEqual({ categories: ['A', 'B', 'C'], series: [{ name: 'Sales', values: ['1.0', null, '3.0'] }] });
      return charted[0]!.chart;
    } finally {
      session.free();
    }
  };
  expect(run(native)).toEqual(run(wasm));
});

test('native and wasm leave hidden table rows out of the display list unless hidden text is shown', () => {
  const encoder = new TextEncoder();
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const border = 'w:val="single" w:sz="24" w:space="0" w:color="000000"';
  const row = (text: string, hidden: boolean) => `<w:tr>${hidden ? '<w:trPr><w:hidden/></w:trPr>' : ''}<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r>${hidden ? '<w:rPr><w:vanish/></w:rPr>' : ''}<w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`;
  const table = (rows: string) => `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblBorders><w:top ${border}/><w:bottom ${border}/><w:insideH ${border}/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`;
  const bytes = wasm.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': encoder.encode(`<w:document xmlns:w="${w}"><w:body>${table(row('Alpha', false) + row('Bravo', true) + row('Charlie', false))}<w:p/>${table(row('Delta', true) + row('Echo', true))}<w:p><w:r><w:t>Tail</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
  });
  type Primitive = { kind: string; text?: string; role?: string; y1?: number; cell?: { row: number } };
  const run = (engine: Engine, showHiddenText: boolean) => {
    const session = new engine.EditSession(79);
    try {
      session.open_docx(bytes, true);
      const font = session.register_measure_font(fontBytes);
      const fontChains = { 'calibri|0|0': [font], 'calibri|1|0': [font] };
      const layout = session.layout_document_with_regions_json(JSON.stringify({
        bodyStory: 'body',
        regions: { sections: [{ sectionId: 'main', properties: {
          pageWidth: 12240, pageHeight: 15840, marginTop: 1440, marginRight: 1440, marginBottom: 1440, marginLeft: 1440,
        } }] },
        measurement: { fontChains, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
        renderEnv: { showHiddenText },
      }));
      const list = JSON.parse(session.build_display_list_json(JSON.stringify({ ...JSON.parse(layout), fontChains }))) as {
        pages: { primitives: Primitive[] }[];
      };
      const primitives = list.pages.flatMap((page) => page.primitives);
      return {
        text: primitives.map((p) => p.text ?? '').join('').replace(/\s/g, ''),
        rows: [...new Set(primitives.flatMap((p) => (p.cell ? [p.cell.row] : [])))].sort(),
        borders: primitives.filter((p) => p.role === 'table-border').length,
      };
    } finally {
      session.free();
    }
  };
  const hidden = run(native, false);
  expect(hidden.text).toBe('AlphaCharlieTail');
  expect(hidden.rows).toEqual([0, 2]);
  expect(hidden).toEqual(run(wasm, false));
  const shown = run(native, true);
  expect(shown.text).toBe('AlphaBravoCharlieDeltaEchoTail');
  expect(shown.borders).toBeGreaterThan(hidden.borders);
  expect(shown).toEqual(run(wasm, true));
});
