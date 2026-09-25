import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  DocxExportError,
  exportDocxStructured,
  exportDocxStructuredWithPages,
  renderDocxMarkdown,
  renderDocxMarkdownWithPages,
  type DocxHeadlessLayoutOptions,
  type DocxPageExportOptions,
  type DocxStorySelection,
} from '../core';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
  )
);
const OTHER_FONT = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/docx-raster/tests/assets/Carlito-Regular.ttf')
  )
);
const PAGES = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/docx-edit/tests/fixtures/page-fragments/pages.docx')
  )
);
const LAID_OUT: DocxStorySelection[] = ['body', 'headers', 'footers', 'footnotes', 'endnotes'];
const FONTS: DocxHeadlessLayoutOptions = {
  fonts: [{ key: 'liberation', data: Buffer.from(FONT).toString('base64') }],
  defaultChain: ['liberation'],
};
const MARKUP: DocxPageExportOptions = { revisionView: 'markup', stories: LAID_OUT };

let nextClientId = 97100;

/** A session holding the fixture, laid out as the editor lays it out with the pinned font. */
async function laidOut(): Promise<{ session: YrsSession; request: string }> {
  const session = await createYrsSession({ clientId: nextClientId++ });
  const { document } = session.openDocx(PAGES, true);
  const font = session.registerFont(FONT);
  const request = buildResidentRegionLayoutRequest(document, 24, {});
  const requirements = JSON.parse(
    session.layoutFontRequirementsJson(JSON.stringify(request))
  ) as Array<{ key: string }>;
  request.measurement = {
    fontChains: Object.fromEntries(requirements.map((requirement) => [requirement.key, [font]])),
    defaults: { fontSize: 11, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
  const json = JSON.stringify(request);
  session.layoutDocumentWithRegionsRetainedJson(json);
  return { session, request: json };
}

describe('paged structured export', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('exports bytes with a deterministic snapshot page map', async () => {
    const first = await exportDocxStructuredWithPages(PAGES, MARKUP, FONTS);
    const second = await exportDocxStructuredWithPages(
      PAGES,
      MARKUP,
      JSON.parse(JSON.stringify(FONTS)) as DocxHeadlessLayoutOptions
    );
    expect(JSON.stringify(second)).toBe(JSON.stringify(first));
    expect(first.structured).toEqual(await exportDocxStructured(PAGES, MARKUP));
    const map = first.layout;
    expect(map.pages.map((page) => page.displayedLabel)).toEqual([
      'i',
      'ii',
      'iii',
      'iv',
      '1',
      '2',
      'C',
      '4',
    ]);
    expect(map.pages.filter((page) => page.parityFiller).map((page) => page.pageIndex)).toEqual([5]);
    expect(map.layoutRevisionView).toBe('markup');
    expect('documentVersion' in map || 'layoutVersion' in map).toBe(false);
    expect(map.snapshotFingerprint).toMatch(/^[0-9a-f]{64}$/);
    expect(map.fragments.every((fragment) => fragment.geometry === null)).toBe(true);
    const occurrences = new Set(map.occurrences.map((occurrence) => occurrence.id));
    expect(map.fragments.every((fragment) => occurrences.has(fragment.occurrenceId))).toBe(true);
    expect(
      map.occurrences
        .filter((occurrence) => occurrence.story === 'hf:rIdHeader2')
        .map((occurrence) => occurrence.pageIndex)
    ).toEqual([1, 2, 3, 6, 7]);
    const split = map.fragments.filter(
      (fragment) => fragment.slice.kind === 'table' && fragment.slice.rows.some((row) => row.repeatedHeader)
    );
    expect(split).toHaveLength(2);
  });

  it('adds page-local geometry only on request', async () => {
    const plain = await exportDocxStructuredWithPages(PAGES, MARKUP, FONTS);
    const measured = await exportDocxStructuredWithPages(
      PAGES,
      { ...MARKUP, includeGeometry: true },
      FONTS
    );
    expect(measured.layout.fragments.map((fragment) => fragment.slice)).toEqual(
      plain.layout.fragments.map((fragment) => fragment.slice)
    );
    const body = measured.layout.fragments.find(
      (fragment) => fragment.slice.kind === 'block' && fragment.occurrenceId === 'p0.body.s0'
    );
    expect(body?.geometry?.unit).toBe('cssPx');
    expect(body?.geometry?.origin).toBe('pageTopLeft');
    for (const fragment of measured.layout.fragments) {
      const page = measured.layout.pages[fragment.pageIndex]!;
      for (const rect of fragment.geometry?.rects ?? []) {
        expect(rect.x + rect.width).toBeLessThanOrEqual(page.size.width + 0.5);
        expect(rect.y + rect.height).toBeLessThanOrEqual(page.size.height + 0.5);
      }
    }
  });

  it('refuses revision views the markup pages do not show, and bad fonts', async () => {
    const refused = await exportDocxStructuredWithPages(
      PAGES,
      { revisionView: 'accepted' },
      FONTS
    ).catch((error: unknown) => error);
    expect(refused).toBeInstanceOf(DocxExportError);
    expect((refused as DocxExportError).failure.code).toBe('unsupported-revision-layout');
    const unknown = await exportDocxStructuredWithPages(PAGES, MARKUP, {
      ...FONTS,
      defaultChain: ['missing'],
    }).catch((error: unknown) => error);
    expect(unknown).toBeInstanceOf(TypeError);
    const limited = await exportDocxStructuredWithPages(
      PAGES,
      { ...MARKUP, maxFragments: 0 },
      FONTS
    ).catch((error: unknown) => error);
    expect((limited as DocxExportError).failure.code).toBe('invalid-options');
    const undecodable = await exportDocxStructuredWithPages(PAGES, MARKUP, {
      fonts: [{ key: 'liberation', data: '%%' }],
      defaultChain: ['liberation'],
    }).catch((error: unknown) => error);
    expect(undecodable).toBeInstanceOf(TypeError);
    const fontless = await exportDocxStructuredWithPages(PAGES, MARKUP, {
      fonts: [],
      defaultChain: [],
    }).catch((error: unknown) => error);
    expect(fontless).toBeInstanceOf(DocxExportError);
    expect((fontless as DocxExportError).failure.code).toBe('layout-unavailable');
  });

  it('measures text naming no font in the default family it is given', async () => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const parts = new Map<string, Uint8Array>([
      [
        '[Content_Types].xml',
        toBytes(
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        ),
      ],
      [
        '_rels/.rels',
        toBytes(
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
        ),
      ],
      [
        'word/document.xml',
        toBytes(
          `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Plain text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
        ),
      ],
    ]);
    const bytes = new Uint8Array(rezipPartsToArrayBuffer(parts));
    const custom: DocxHeadlessLayoutOptions = {
      fonts: FONTS.fonts,
      fontChains: { 'liberation sans': ['liberation'] },
      defaultChain: [],
      measurementDefaults: { fontFamily: 'Liberation Sans' },
    };
    const paged = await exportDocxStructuredWithPages(bytes, MARKUP, custom);
    expect(paged.layout.fragments.length).toBeGreaterThan(0);
    const unmeasured = await exportDocxStructuredWithPages(bytes, MARKUP, {
      ...custom,
      fontChains: { calibri: ['liberation'] },
    }).catch((error: unknown) => error);
    expect(unmeasured).toBeInstanceOf(DocxExportError);
    expect((unmeasured as DocxExportError).failure.code).toBe('layout-unavailable');
  });

  it('fingerprints the measurement defaults and render environment it is given', async () => {
    const base = (await exportDocxStructuredWithPages(PAGES, MARKUP, FONTS)).layout.provenance;
    const larger = (
      await exportDocxStructuredWithPages(PAGES, MARKUP, {
        ...FONTS,
        measurementDefaults: { fontSize: 12 },
      })
    ).layout.provenance;
    expect(larger.fontSetFingerprint).not.toBe(base.fontSetFingerprint);
    expect(larger.optionsFingerprint).toBe(base.optionsFingerprint);
    const hidden = (
      await exportDocxStructuredWithPages(PAGES, MARKUP, {
        ...FONTS,
        renderEnvironment: { showHiddenText: true },
      })
    ).layout.provenance;
    expect(hidden.optionsFingerprint).not.toBe(base.optionsFingerprint);
    expect(hidden.fontSetFingerprint).toBe(base.fontSetFingerprint);
  });

  it('leaves the fonts of a live session alone', async () => {
    const { session, request } = await laidOut();
    try {
      const live = session.exportStructuredWithPages(MARKUP);
      if (!live.ok) throw new Error(live.failure.message);
      const fingerprints = new Set<string>();
      for (const data of [OTHER_FONT, FONT]) {
        const layout = {
          fonts: [{ key: 'face', data: Buffer.from(data).toString('base64') }],
          defaultChain: ['face'],
        };
        const paged = await exportDocxStructuredWithPages(PAGES, MARKUP, layout).catch(
          (error: unknown) => error
        );
        if (!(paged instanceof DocxExportError)) {
          if (paged instanceof Error) throw paged;
          fingerprints.add(
            (paged as { layout: { provenance: { fontSetFingerprint: string } } }).layout.provenance
              .fontSetFingerprint
          );
        }
        expect(session.exportStructuredWithPages(MARKUP)).toEqual(live);
      }
      expect(fingerprints.has(live.content.layout.provenance.fontSetFingerprint)).toBe(true);
      session.layoutDocumentWithRegionsRetainedJson(request);
      const relaid = session.exportStructuredWithPages(MARKUP);
      if (!relaid.ok) throw new Error(relaid.failure.message);
      expect(relaid.content.layout.fragments).toEqual(live.content.layout.fragments);
      expect(relaid.content.layout.provenance.fontSetFingerprint).toBe(
        live.content.layout.provenance.fontSetFingerprint
      );
    } finally {
      session.destroy();
    }
  });

  it('exports a live session against its retained layout and refuses stale ones', async () => {
    const empty = await createYrsSession({ clientId: nextClientId++ });
    try {
      empty.openDocx(PAGES, true);
      expect(empty.exportStructuredWithPages(MARKUP)).toMatchObject({
        ok: false,
        failure: { code: 'layout-unavailable', target: null },
      });
    } finally {
      empty.destroy();
    }
    const { session, request } = await laidOut();
    try {
      session.beginUndoCapture();
      const version = session.version();
      const state = session.encodeState();
      const saved = session.materializeDocx();
      const history = [session.canUndo(), session.canRedo()];
      const read = session.exportStructuredWithPages(MARKUP);
      if (!read.ok) throw new Error(read.failure.message);
      expect(read.version).toBe(version);
      expect(read.content.layout.documentVersion).toBe(version);
      expect(read.content.structured.anchorScope).toBe('session');
      expect('frameEpoch' in read.content.layout.provenance).toBe(false);
      expect(session.version()).toBe(version);
      expect(session.encodeState()).toEqual(state);
      expect(session.materializeDocx()).toEqual(saved);
      expect([session.canUndo(), session.canRedo()]).toEqual(history);
      const layoutVersion = read.content.layout.layoutVersion;
      expect(
        session.exportStructuredWithPages({ ...MARKUP, expectLayoutVersion: layoutVersion }).ok
      ).toBe(true);
      expect(
        session.exportStructuredWithPages({ ...MARKUP, expectLayoutVersion: `${layoutVersion}x` })
      ).toMatchObject({ ok: false, failure: { code: 'stale-layout' } });
      expect(session.exportStructuredWithPages({ revisionView: 'original' })).toMatchObject({
        ok: false,
        failure: { code: 'unsupported-revision-layout' },
      });

      const applied = session.applyEdits({
        expectVersion: session.version(),
        steps: [
          {
            op: 'replaceText',
            target: { kind: 'paragraph', story: 'body', paraId: '00000001' },
            text: 'Edited title',
          },
        ],
      });
      expect(applied.ok).toBe(true);
      expect(session.exportStructuredWithPages(MARKUP)).toMatchObject({
        ok: false,
        version: session.version(),
        failure: { code: 'stale-document' },
      });
      session.layoutDocumentWithRegionsRetainedJson(request);
      const relaid = session.exportStructuredWithPages(MARKUP);
      if (!relaid.ok) throw new Error(relaid.failure.message);
      expect(relaid.content.layout.layoutVersion).not.toBe(layoutVersion);

      const other = await createYrsSession({ clientId: nextClientId++ });
      try {
        other.registerFont(FONT);
        expect(session.exportStructuredWithPages(MARKUP)).toMatchObject({
          ok: false,
          failure: { code: 'stale-layout' },
        });
      } finally {
        other.destroy();
      }
    } finally {
      session.destroy();
    }
  });

  it('renders page markers from a matching map only', async () => {
    const paged = await exportDocxStructuredWithPages(PAGES, MARKUP, FONTS);
    const plain = await renderDocxMarkdown(paged.structured);
    expect(await renderDocxMarkdownWithPages(paged)).toEqual(plain);
    const marked = await renderDocxMarkdownWithPages(paged, { pageMarkers: true });
    expect(marked.markdown).toContain('<!-- docx-export:0 --><!-- docx-pages: 0=i -->');
    expect(marked.anchors).toEqual(plain.anchors);
    const mismatched = await renderDocxMarkdownWithPages({
      ...paged,
      structured: { ...paged.structured, truncated: !paged.structured.truncated },
    }).catch((error: unknown) => error);
    expect(mismatched).toBeInstanceOf(DocxExportError);
    expect((mismatched as DocxExportError).failure.code).toBe('invalid-options');
  });
});
