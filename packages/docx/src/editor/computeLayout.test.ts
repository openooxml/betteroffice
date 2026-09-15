import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsRenderEnv, type YrsSession } from '../yrs';
import { documentToYrs } from '../yrs/documentToYrs';
import { yrsToDocument } from '../yrs/yrsToDocument';
import type { Document, Paragraph } from '../types/document';
import {
  buildResidentRegionLayoutRequest,
  computeLayout,
  getLayoutKernelInputs,
  type ComputeLayoutInputs,
} from './computeLayout';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

type LoweredParagraph = {
  attrs?: { styleId?: string; effectiveStyleId?: string };
  runs?: unknown[];
};

function stylesDoc(
  styles: Array<{ styleId: string; type: string; default?: boolean; name?: string }>
): Document {
  return {
    package: { styles: { styles }, document: { content: [] } },
  } as unknown as Document;
}

function textParagraph(text: string, styleId?: string): Paragraph {
  return {
    type: 'paragraph',
    ...(styleId === undefined ? {} : { formatting: { styleId } }),
    content: [{ type: 'run', content: [{ type: 'text', text }] } as never],
  } as Paragraph;
}

function bodyDoc(
  paragraphs: Paragraph[],
  styles: Array<{ styleId: string; type: string; default?: boolean; name?: string }>
): Document {
  return {
    package: { document: { content: paragraphs }, styles: { styles } },
  } as unknown as Document;
}

function lowered(session: Pick<YrsSession, 'yrsBlocksForStory'>, env: YrsRenderEnv) {
  return session.yrsBlocksForStory('body', env) as LoweredParagraph[];
}

describe('computeLayout retained kernel inputs', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  test('the measured arena is fetched lazily, once, and only on demand', () => {
    const layout = { pages: [] };
    let kernelFetches = 0;
    const session = {
      layoutDocumentWithRegionsRetainedJson: () =>
        JSON.stringify({ layout, notesConverged: true }),
      residentWorkerProbe: () => ({ layoutRevision: 1 }),
      retainedKernelInputsJson: (expectedLayoutRevision: number) => {
        expect(expectedLayoutRevision).toBe(1);
        kernelFetches += 1;
        return JSON.stringify({ measured: [{ block: { kind: 'paragraph' } }], options: { pageGap: 24 } });
      },
    };
    const computation = computeLayout({
      document: null,
      pageGap: 24,
      session: session as never,
      renderEnv: {},
      measurement: { fontChains: {}, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true } as never,
    });
    expect(computation.notesConverged).toBe(true);
    const inputs = getLayoutKernelInputs(computation.layout);
    expect(inputs).toBeDefined();
    expect(kernelFetches).toBe(0);
    expect(inputs!.measured.length).toBe(1);
    expect(inputs!.options).toEqual({ pageGap: 24 });
    expect(kernelFetches).toBe(1);
  });

  test('an older layout cannot fetch a newer retained arena', async () => {
    const session = await createYrsSession({ clientId: 247 });
    try {
      const { paraId } = session.createStory('body', 'first');
      const fontId = session.registerFont(new Uint8Array(readFileSync(FONT)));
      const inputs: ComputeLayoutInputs = {
        document: null,
        pageGap: 24,
        session,
        renderEnv: {},
        measurement: {
          fontChains: { 'calibri|0|0': [fontId] },
          defaults: { fontSize: 11, fontFamily: 'Calibri' },
          compat: { noLeading: false, doNotExpandShiftReturn: false },
          authoritativeShaping: true,
        },
      };

      const first = computeLayout(inputs);
      session.insertText({ story: 'body', paraId, offset: 5 }, ' second');
      const second = computeLayout(inputs);
      const firstKernel = getLayoutKernelInputs(first.layout);
      const secondKernel = getLayoutKernelInputs(second.layout);

      expect(() => firstKernel!.measured).toThrow(
        'retained layout revision mismatch: expected 1, current 2'
      );
      expect(secondKernel!.measured).toHaveLength(1);
    } finally {
      session.destroy();
    }
  });
});

describe('buildResidentRegionLayoutRequest defaultParagraphStyleId', () => {
  test('empty without styles and without document', () => {
    expect(buildResidentRegionLayoutRequest(null, 24, {}).renderEnv.defaultParagraphStyleId).toBeUndefined();
    const empty = stylesDoc([]);
    expect(buildResidentRegionLayoutRequest(empty, 24, {}).renderEnv.defaultParagraphStyleId).toBeUndefined();
  });

  test('derives Normal when present without explicit default', () => {
    const doc = stylesDoc([{ styleId: 'Normal', type: 'paragraph' }]);
    expect(buildResidentRegionLayoutRequest(doc, 24, {}).renderEnv.defaultParagraphStyleId).toBe('Normal');
  });

  test('custom BodyDefault wins over existing Normal', () => {
    const doc = stylesDoc([
      { styleId: 'BodyDefault', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
    ]);
    expect(buildResidentRegionLayoutRequest(doc, 24, {}).renderEnv.defaultParagraphStyleId).toBe(
      'BodyDefault'
    );
  });

  test('explicit caller override wins over derived default', () => {
    const doc = stylesDoc([
      { styleId: 'BodyDefault', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
    ]);
    const request = buildResidentRegionLayoutRequest(doc, 24, {
      defaultParagraphStyleId: 'Override',
    });
    expect(request.renderEnv.defaultParagraphStyleId).toBe('Override');
    const bare = stylesDoc([]);
    const preserved = buildResidentRegionLayoutRequest(bare, 24, {
      defaultParagraphStyleId: 'Normal',
    });
    expect(preserved.renderEnv.defaultParagraphStyleId).toBe('Normal');
  });

  test('ignores character defaults and empty ids', () => {
    const characterOnly = stylesDoc([{ styleId: 'BodyDefault', type: 'character', default: true }]);
    expect(
      buildResidentRegionLayoutRequest(characterOnly, 24, {}).renderEnv.defaultParagraphStyleId
    ).toBeUndefined();
    const emptyId = stylesDoc([
      { styleId: '', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
    ]);
    expect(buildResidentRegionLayoutRequest(emptyId, 24, {}).renderEnv.defaultParagraphStyleId).toBe(
      'Normal'
    );
  });

  test('react-shaped env derives without mutation and keeps toc merge', () => {
    const doc = stylesDoc([
      { styleId: 'BodyDefault', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
      { styleId: 'TOC1', type: 'paragraph', name: 'TOC 1' },
    ]);
    const renderEnv: YrsRenderEnv = {
      themeColors: { accent1: '4472C4' },
      defaultTabStopTwips: 720,
      numericIds: {},
      showHiddenText: false,
    };
    const snapshot = structuredClone(renderEnv);
    const request = buildResidentRegionLayoutRequest(doc, 24, renderEnv);
    expect(renderEnv).toEqual(snapshot);
    expect(request.renderEnv.defaultParagraphStyleId).toBe('BodyDefault');
    expect(request.renderEnv.tocStyleIds).toContain('TOC1');
    expect(request.renderEnv.themeColors).toEqual({ accent1: '4472C4' });
    expect(request.renderEnv).not.toBe(renderEnv);
  });

  test('incremental document change updates derived default', () => {
    const normal = stylesDoc([{ styleId: 'Normal', type: 'paragraph' }]);
    const custom = stylesDoc([
      { styleId: 'BodyDefault', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
    ]);
    const first = buildResidentRegionLayoutRequest(normal, 24, {});
    const second = buildResidentRegionLayoutRequest(custom, 24, {});
    expect(first.renderEnv.defaultParagraphStyleId).toBe('Normal');
    expect(second.renderEnv.defaultParagraphStyleId).toBe('BodyDefault');
    expect(first.renderEnv).not.toBe(second.renderEnv);
    const repeat = buildResidentRegionLayoutRequest(custom, 24, {});
    expect(repeat.renderEnv.defaultParagraphStyleId).toBe('BodyDefault');
    expect(repeat.renderEnv).not.toBe(second.renderEnv);
  });
});

describe('computeLayout default style forwarding', () => {
  test('forwards derived default and preserves caller override', () => {
    const doc = stylesDoc([
      { styleId: 'BodyDefault', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
    ]);
    const seen: unknown[] = [];
    const session = {
      layoutDocumentWithRegionsRetainedJson: (input: string) => {
        seen.push(JSON.parse(input).renderEnv);
        return JSON.stringify({ layout: { pages: [] }, notesConverged: true });
      },
      residentWorkerProbe: () => ({ layoutRevision: 1 }),
      retainedKernelInputsJson: () => JSON.stringify({ measured: [], options: {} }),
    };
    computeLayout({
      document: doc,
      pageGap: 24,
      session: session as never,
      renderEnv: {},
      measurement: { fontChains: {}, defaults: { fontSize: 11, fontFamily: 'Calibri' } } as never,
    });
    expect((seen[0] as YrsRenderEnv).defaultParagraphStyleId).toBe('BodyDefault');
    computeLayout({
      document: doc,
      pageGap: 24,
      session: session as never,
      renderEnv: { defaultParagraphStyleId: 'Override' },
      measurement: { fontChains: {}, defaults: { fontSize: 11, fontFamily: 'Calibri' } } as never,
    });
    expect((seen[1] as YrsRenderEnv).defaultParagraphStyleId).toBe('Override');
  });

  test('legacy empty env without document still succeeds', () => {
    const session = {
      layoutDocumentWithRegionsRetainedJson: () =>
        JSON.stringify({ layout: { pages: [] }, notesConverged: true }),
      residentWorkerProbe: () => ({ layoutRevision: 1 }),
      retainedKernelInputsJson: () => JSON.stringify({ measured: [], options: {} }),
    };
    const computation = computeLayout({
      document: null,
      pageGap: 24,
      session: session as never,
      renderEnv: {},
      measurement: { fontChains: {}, defaults: { fontSize: 11, fontFamily: 'Calibri' } } as never,
    });
    expect(computation.layout).toMatchObject({ pages: [] });
  });
});

describe('frontend derived default lowering', () => {
  test('implicit Normal matches explicit both orders with raw preserved', async () => {
    const styles = [{ styleId: 'Normal', type: 'paragraph' }];
    for (const explicitFirst of [false, true]) {
      const paragraphs = explicitFirst
        ? [textParagraph('FIRST', 'Normal'), textParagraph('SECOND')]
        : [textParagraph('FIRST'), textParagraph('SECOND', 'Normal')];
      const document = bodyDoc(paragraphs, styles);
      const session = await createYrsSession({ clientId: explicitFirst ? 24811 : 24810 });
      try {
        documentToYrs(session, document);
        const env = buildResidentRegionLayoutRequest(document, 24, {}).renderEnv;
        expect(env.defaultParagraphStyleId).toBe('Normal');
        const blocks = lowered(session, env);
        const legacy = lowered(session, {});
        expect(legacy[0]?.attrs?.effectiveStyleId).toBeUndefined();
        expect(blocks[0]?.attrs?.styleId).toBe(explicitFirst ? 'Normal' : undefined);
        expect(blocks[1]?.attrs?.styleId).toBe(explicitFirst ? undefined : 'Normal');
        const implicitIndex = explicitFirst ? 1 : 0;
        const explicitIndex = explicitFirst ? 0 : 1;
        expect(blocks[implicitIndex]?.attrs?.effectiveStyleId).toBe('Normal');
        expect(blocks[explicitIndex]?.attrs?.effectiveStyleId).toBeUndefined();
        const raw = session.paragraphs('body').map((paragraph) => paragraph.properties['pStyle'] ?? null);
        expect(raw).toEqual(explicitFirst ? ['Normal', null] : [null, 'Normal']);
        const state = session.encodeState();
        lowered(session, env);
        expect(session.encodeState()).toEqual(state);
      } finally {
        session.destroy();
      }
    }
  });

  test('custom default keeps Normal distinct and matches BodyDefault both orders', async () => {
    const styles = [
      { styleId: 'BodyDefault', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
      { styleId: 'Different', type: 'paragraph' },
    ];
    const cases: Array<{ paragraphs: Paragraph[]; implicit: number; explicit?: string }> = [
      { paragraphs: [textParagraph('FIRST'), textParagraph('SECOND', 'Normal')], implicit: 0, explicit: 'Normal' },
      { paragraphs: [textParagraph('FIRST'), textParagraph('SECOND', 'BodyDefault')], implicit: 0, explicit: 'BodyDefault' },
      { paragraphs: [textParagraph('FIRST', 'BodyDefault'), textParagraph('SECOND')], implicit: 1, explicit: 'BodyDefault' },
      { paragraphs: [textParagraph('FIRST'), textParagraph('SECOND', 'Different')], implicit: 0, explicit: 'Different' },
    ];
    for (const [index, entry] of cases.entries()) {
      const document = bodyDoc(entry.paragraphs, styles);
      const session = await createYrsSession({ clientId: 24820 + index });
      try {
        documentToYrs(session, document);
        const env = buildResidentRegionLayoutRequest(document, 24, {}).renderEnv;
        expect(env.defaultParagraphStyleId).toBe('BodyDefault');
        const blocks = lowered(session, env);
        expect(blocks[entry.implicit]?.attrs?.effectiveStyleId).toBe('BodyDefault');
        expect(blocks[entry.implicit]?.attrs?.styleId).toBeUndefined();
        const other = entry.implicit === 0 ? 1 : 0;
        expect(blocks[other]?.attrs?.styleId).toBe(entry.explicit);
        const same = entry.explicit === 'BodyDefault';
        expect((blocks[entry.implicit]?.attrs?.effectiveStyleId ?? blocks[entry.implicit]?.attrs?.styleId) ===
          (blocks[other]?.attrs?.effectiveStyleId ?? blocks[other]?.attrs?.styleId)).toBe(same);
      } finally {
        session.destroy();
      }
    }
  });

  test('explicit override changes derived identity without touching raw', async () => {
    const document = bodyDoc([textParagraph('FIRST'), textParagraph('SECOND', 'Normal')], [
      { styleId: 'BodyDefault', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
    ]);
    const session = await createYrsSession({ clientId: 24830 });
    try {
      documentToYrs(session, document);
      const derived = buildResidentRegionLayoutRequest(document, 24, {}).renderEnv;
      const overridden = buildResidentRegionLayoutRequest(document, 24, {
        defaultParagraphStyleId: 'Normal',
      }).renderEnv;
      expect(derived.defaultParagraphStyleId).toBe('BodyDefault');
      expect(overridden.defaultParagraphStyleId).toBe('Normal');
      const derivedBlocks = lowered(session, derived);
      const overriddenBlocks = lowered(session, overridden);
      expect(derivedBlocks[0]?.attrs?.effectiveStyleId).toBe('BodyDefault');
      expect(overriddenBlocks[0]?.attrs?.effectiveStyleId).toBe('Normal');
      expect(derivedBlocks[0]?.attrs?.styleId).toBeUndefined();
      expect(overriddenBlocks[0]?.attrs?.styleId).toBeUndefined();
      expect(derivedBlocks[1]?.attrs?.styleId).toBe('Normal');
    } finally {
      session.destroy();
    }
  });

  test('save projection keeps raw pStyle after derived lowering', async () => {
    const document = bodyDoc([textParagraph('FIRST'), textParagraph('SECOND', 'Normal')], [
      { styleId: 'BodyDefault', type: 'paragraph', default: true },
      { styleId: 'Normal', type: 'paragraph' },
    ]);
    const session = await createYrsSession({ clientId: 24840 });
    try {
      documentToYrs(session, document);
      const env = buildResidentRegionLayoutRequest(document, 24, {}).renderEnv;
      lowered(session, env);
      lowered(session, {});
      const saved = yrsToDocument(session, document);
      const savedParagraphs = saved.package.document.content.filter(
        (block): block is Paragraph => block.type === 'paragraph'
      );
      expect(savedParagraphs).toHaveLength(2);
      expect(savedParagraphs[0]?.formatting?.styleId).toBeUndefined();
      expect(savedParagraphs[1]?.formatting?.styleId).toBe('Normal');
    } finally {
      session.destroy();
    }
  });

  test('legacy direct callers without env remain compatible', async () => {
    const session = await createYrsSession({ clientId: 24850 });
    try {
      session.loadStories([{ storyId: 'body', paragraphs: [{ text: 'legacy' }] }]);
      expect(() => session.yrsBlocksForStory('body')).not.toThrow();
      expect(() => session.yrsBlocksForStory('body', {})).not.toThrow();
      const blocks = session.yrsBlocksForStory('body') as LoweredParagraph[];
      expect(blocks).toHaveLength(1);
      expect(blocks[0]?.attrs?.effectiveStyleId).toBeUndefined();
    } finally {
      session.destroy();
    }
  });
});
