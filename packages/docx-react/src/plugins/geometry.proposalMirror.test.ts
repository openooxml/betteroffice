import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  computeAnchorTargetGeometry,
  computeProposalGeometryMirror,
  createYrsPositionProjection,
  createYrsSession,
  proposalSetIdentity,
  yrsLocToProjectedDisplayPosition,
  type AnchorGeometryTarget,
  type DocxProposalSnapshot,
  type ProposalGeometryMirror,
  type ProposalGeometryTarget,
  type YrsLoc,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { createProposalRegistry } from '@betteroffice/docx/yrs/proposals';
import { createResidentEngineSession } from '@betteroffice/docx/yrs/residentEngineSession';
import { stampRevisionPreviewKey } from '../components/DocxEditor/internals/layoutProvenance';
import { createPluginGeometry } from './geometry';
import type { DocxAnchorGeometryResult } from './types';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

beforeAll(() =>
  preloadEditWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
  )))
);

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function mirroredGeometry() {
  const pages = document.createElement('div');
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  pages.appendChild(canvas);
  const layer = document.createElement('div');
  const bounds = {
    x: 0,
    y: 0,
    left: 0,
    top: 0,
    width: 100,
    height: 200,
    right: 100,
    bottom: 200,
    toJSON() {},
  } as DOMRect;
  pages.getBoundingClientRect = () => bounds;
  canvas.getBoundingClientRect = () => bounds;
  layer.getBoundingClientRect = () => bounds;
  const anchorPositions: number[] = [];
  const queries = {
    pageSize: () => ({ width: 100, height: 200 }),
    pageCount: () => 1,
    rangeRects: () => [],
    anchorRect: (position: number) => {
      anchorPositions.push(position);
      return { pageIndex: 0, x: position, y: 10, width: 1, height: 12 };
    },
    pageBounds: () => ({ pageIndex: 0, x: 0, y: 0, width: 100, height: 200 }),
  } as unknown as DisplayListQueries;
  stampRevisionPreviewKey(queries, '');
  let snapshot: DocxProposalSnapshot = {
    version: 'v1',
    previewVersion: 0,
    proposals: [
      {
        id: 'proposal',
        state: 'proposed',
        paragraph: { kind: 'session', sessionId: 'session', story: 'body', paraId: 'first' },
        revisionIds: [],
        changed: false,
      },
    ],
  };
  const session = {
    version: () => 'v1',
    getProposals: () => snapshot,
  } as unknown as YrsSession;
  let mirror: ProposalGeometryMirror = {
    version: 'v1',
    previewVersion: 0,
    proposals: proposalSetIdentity(snapshot),
    targets: {
      proposal: { ok: true, ranges: [{ from: 4, to: 8 }], paragraph: 3 },
      missing: {
        ok: false,
        failure: { code: 'missing-target', message: 'The paragraph no longer exists' },
      },
    },
    hidden: [{ from: 4, to: 8 }],
  };
  let pending = false;
  let presented = true;
  const anchorTarget = mock((
    _target: Exclude<AnchorGeometryTarget, { kind: 'proposal' }>
  ): ProposalGeometryTarget | undefined => undefined);
  const geometry = createPluginGeometry(
    { id: 'layout', version: 'v1', previewVersion: 0, zoom: 1, pageCount: 1 },
    createRenderedDomContext(pages, 1),
    layer,
    () => true,
    () => null,
    queries,
    () => ({
      session,
      proposalGeometry: mirror,
      anchorTarget,
      presented,
      editor: {
        hasPendingInput: () => pending,
        yrsLocToDisplayPosition: () => {
          throw new Error('a mirrored proposal must not read replica positions');
        },
      },
    })
  );
  return {
    geometry,
    anchorPositions,
    anchorTarget,
    snapshot,
    mirror,
    setSnapshot: (value: DocxProposalSnapshot) => {
      snapshot = value;
    },
    setMirror: (value: ProposalGeometryMirror) => {
      mirror = value;
    },
    setPending: (value: boolean) => {
      pending = value;
    },
    setPresented: (value: boolean) => {
      presented = value;
    },
  };
}

test('serves proposal and revision geometry from a mirror while the replica has no anchor reads', () => {
  const { geometry, anchorPositions, anchorTarget } = mirroredGeometry();
  expect(geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toMatchObject({
    ok: true,
    rects: [],
    anchor: { pageIndex: 0, width: 0 },
  });
  expect(anchorPositions).toEqual([3]);
  expect(geometry.getAnchorGeometry({ kind: 'proposal', id: 'missing' })).toEqual({
    ok: false,
    failure: { code: 'missing-target', message: 'The paragraph no longer exists' },
  });
  expect(geometry.getAnchorGeometry({ kind: 'proposal', id: 'unknown' })).toMatchObject({
    ok: false,
    failure: { code: 'unknown-proposal' },
  });
  expect(geometry.getAnchorGeometry({ kind: 'proposal', id: '__proto__' })).toMatchObject({
    ok: false,
    failure: { code: 'unknown-proposal' },
  });
  expect(anchorTarget).not.toHaveBeenCalled();
  anchorTarget.mockReturnValue({ ok: true, ranges: [{ from: 10, to: 14 }], paragraph: 9 });
  const target = { kind: 'revision', revisionId: 'r1' } as const;
  expect(geometry.getAnchorGeometry(target)).toMatchObject({
    ok: true,
    rects: [],
    anchor: { pageIndex: 0, x: 9, width: 0 },
  });
  expect(anchorTarget).toHaveBeenCalledTimes(1);
  expect(anchorTarget).toHaveBeenCalledWith(target);
  anchorTarget.mockClear();
  expect(geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toMatchObject({
    ok: true,
  });
  expect(anchorTarget).not.toHaveBeenCalled();
});

test('waits for a non-proposal target until the authority resolves it', () => {
  const { geometry, anchorTarget } = mirroredGeometry();
  expect(geometry.getAnchorGeometry({ kind: 'revision', revisionId: 'r1' })).toEqual({
    ok: false,
    failure: { code: 'layout-unavailable', message: 'No rendered layout shows this target yet' },
  });
  const failure: ProposalGeometryTarget = {
    ok: false,
    failure: { code: 'missing-target', message: 'The revision no longer exists' },
  };
  anchorTarget.mockReturnValue(failure);
  expect(geometry.getAnchorGeometry({ kind: 'revision', revisionId: 'r1' })).toEqual(failure);
});

test('matches default anchor geometry exactly across lines, pages and table cells', async () => {
  const paragraph = (paraId: string, text: string) =>
    `<w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${paragraph('00000001', 'same 😀same same tail')}${paragraph('00000002', 'Other paragraph')}<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${paragraph('0000C001', 'cell same same')}</w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`)],
  ]);
  const main = await createYrsSession({ clientId: 79105 });
  const resident = await createResidentEngineSession();
  try {
    const bytes = new Uint8Array(rezipPartsToArrayBuffer(parts));
    main.openDocx(bytes, true);
    resident.openDocx(bytes);
    const identities = main.paragraphIdentities();
    const anchors = ['00000001', '00000002', '0000C001'].map((paraId) =>
      identities.paragraphs.find(({ ooxmlParaId }) => ooxmlParaId === paraId)!.persisted!
    );
    const request = (expectVersion: string) => ({
      expectVersion,
      proposals: [{
        id: 'tail', paragraph: anchors[0]!, op: 'replaceText' as const, search: 'tail', replaceWith: 'ending',
        suggest: { author: 'Reviewer', date: '2026-09-29T12:00:00Z' },
      }],
    });
    const proposed = main.proposeChanges(request(main.version()));
    if (!proposed.ok) throw new Error(proposed.failure.message);
    const registry = createProposalRegistry(resident.proposalEngine);
    const residentProposed = registry.propose(request(resident.proposalEngine.version()));
    if (!residentProposed.ok) throw new Error(residentProposed.failure.message);
    const snapshot = main.getProposals();
    const residentSnapshot = registry.snapshot();
    const mirror = computeProposalGeometryMirror(resident.geometryReader, residentSnapshot);
    const stub = {
      version: () => mirror.version,
      getProposals: () => residentSnapshot,
    } as unknown as YrsSession;
    const pages = document.createElement('div');
    const layer = document.createElement('div');
    pages.getBoundingClientRect = () => new DOMRect(20, 30, 200, 824);
    layer.getBoundingClientRect = () => new DOMRect(5, 10, 200, 824);
    for (let pageIndex = 0; pageIndex < 2; pageIndex += 1) {
      const canvas = document.createElement('canvas');
      canvas.dataset.pageIndex = String(pageIndex);
      canvas.getBoundingClientRect = () => new DOMRect(20, 30 + pageIndex * 424, 200, 400);
      pages.append(canvas);
    }
    const rectAt = (position: number): DisplayListRect => {
      const line = Math.floor(position / 8);
      return {
        pageIndex: line < 2 ? 0 : 1,
        x: 10 + position % 8 * 5,
        y: 10 + (line < 2 ? line : line - 2) * 18,
        width: 5, height: 12,
      };
    };
    const queries = {
      pageCount: () => 2,
      pageSize: () => ({ width: 200, height: 400 }),
      pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, width: 200, height: 400 }),
      rangeRects: (from: number, to: number) => {
        const rects: DisplayListRect[] = [];
        for (let position = from; position < to;) {
          const end = Math.min(to, (Math.floor(position / 8) + 1) * 8);
          rects.push({ ...rectAt(position), width: (end - position) * 5 });
          position = end;
        }
        return rects;
      },
      hitTestRegions: (pageIndex: number, x: number, y: number) => ({
        region: 'body', target: 'text',
        pos: (pageIndex === 0 ? 0 : 16) + Math.floor((y - 10) / 18) * 8 + Math.round((x - 10) / 5),
      }),
      anchorRect: (position: number) => ({ ...rectAt(position), width: 1 }),
    } as unknown as DisplayListQueries;
    stampRevisionPreviewKey(queries, '');
    const dom = createRenderedDomContext(pages, 1);
    const projection = createYrsPositionProjection(main, 'body');
    const positionFor = (loc: YrsLoc) => projection?.positionForLoc(loc) ??
      (loc.story === 'body' ? yrsLocToProjectedDisplayPosition(main, () => projection, loc) : null);
    const defaultGeometry = createPluginGeometry(
      {
        id: 'default-layout', version: main.version(), previewVersion: snapshot.previewVersion,
        zoom: 1, pageCount: 2,
      },
      dom, layer, () => true, () => null, queries,
      () => ({
        session: main,
        presented: true,
        editor: { hasPendingInput: () => false, yrsLocToDisplayPosition: positionFor },
      })
    );
    const mirrorGeometry = createPluginGeometry(
      {
        id: 'mirror-layout', version: mirror.version, previewVersion: mirror.previewVersion,
        zoom: 1, pageCount: 2,
      },
      dom, layer, () => true, () => null, queries,
      () => ({
        session: stub,
        presented: true,
        proposalGeometry: mirror,
        anchorTarget: (target) => computeAnchorTargetGeometry(resident.geometryReader, [target])[0],
        editor: {
          hasPendingInput: () => false,
          yrsLocToDisplayPosition: () => {
            throw new Error('mirrored geometry must not read replica positions');
          },
        },
      })
    );
    const range = {
      story: 'body', view: 'accepted',
      start: { paraId: '00000001', offset: 0 }, end: { paraId: '00000001', offset: 4 },
    } as const;
    const targets: Exclude<AnchorGeometryTarget, { kind: 'proposal' }>[] = [
      ...anchors.map((paragraph) => ({ kind: 'paragraph' as const, paragraph })),
      { kind: 'search', paragraph: anchors[0]!, text: 'same', occurrence: 'first' },
      { kind: 'search', paragraph: anchors[0]!, text: 'same', occurrence: 'all' },
      { kind: 'search', paragraph: anchors[0]!, text: 'same', occurrence: 2 },
      { kind: 'search', paragraph: anchors[2]!, text: 'same', occurrence: 'all' },
      { kind: 'range', version: main.version(), range },
      { kind: 'revision', revisionId: snapshot.proposals[0]!.revisionIds[0]! },
    ];
    const results: DocxAnchorGeometryResult[] = [];
    for (const target of targets) {
      const expected = defaultGeometry.getAnchorGeometry(target);
      const residentTarget = target.kind === 'range'
        ? { ...target, version: mirror.version }
        : target.kind === 'revision'
          ? { ...target, revisionId: residentSnapshot.proposals[0]!.revisionIds[0]! }
          : target;
      const actual = mirrorGeometry.getAnchorGeometry(residentTarget);
      expect(expected).toMatchObject({ ok: true });
      expect(actual).toMatchObject({ ok: true });
      if (!expected.ok || !actual.ok) throw new Error('The target has no geometry');
      expect(expected.rects.length).toBeGreaterThan(0);
      expect(actual).toEqual({ ...expected, version: mirror.version, layoutId: 'mirror-layout' });
      results.push(actual);
    }
    const first = results[0]!;
    if (!first.ok) throw new Error(first.failure.message);
    expect(first.rects.length).toBeGreaterThan(2);
    expect([...new Set(first.rects.map(({ pageIndex }) => pageIndex))]).toEqual([0, 1]);
    expect(first.anchor).toMatchObject({ pageIndex: 1, width: 0 });
    for (const target of [
      { kind: 'search', paragraph: anchors[0]!, text: 'absent' },
      { kind: 'revision', revisionId: 'unknown' },
      { kind: 'range', version: 'stale', range },
    ] as const) {
      const expected = defaultGeometry.getAnchorGeometry(target);
      expect(expected).toMatchObject({
        ok: false,
        failure: { code: target.kind === 'range' ? 'stale-version' : 'missing-target' },
      });
      expect(mirrorGeometry.getAnchorGeometry(target)).toEqual(expected);
    }
  } finally {
    resident.destroy();
    main.destroy();
  }
});

test('refuses mirrored geometry after withdrawal without version changes', () => {
  const state = mirroredGeometry();
  state.setSnapshot({ ...state.snapshot, proposals: [] });
  expect(state.geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toEqual({
    ok: false,
    failure: { code: 'layout-unavailable', message: 'No rendered layout shows this target yet' },
  });
  expect(state.anchorPositions).toEqual([]);
});

test('waits for a matching mirror when the same proposal id moves to another paragraph', () => {
  const state = mirroredGeometry();
  const snapshot: DocxProposalSnapshot = {
    ...state.snapshot,
    proposals: state.snapshot.proposals.map((proposal) => ({
      ...proposal,
      paragraph: { ...proposal.paragraph, paraId: 'second' },
    })),
  };
  state.setSnapshot(snapshot);
  expect(state.geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toEqual({
    ok: false,
    failure: { code: 'layout-unavailable', message: 'No rendered layout shows this target yet' },
  });
  expect(state.anchorPositions).toEqual([]);

  state.setMirror({
    ...state.mirror,
    proposals: proposalSetIdentity(snapshot),
    targets: { proposal: { ok: true, ranges: [], paragraph: 9 } },
  });
  expect(state.geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toMatchObject({
    ok: true,
    version: state.mirror.version,
    previewVersion: state.mirror.previewVersion,
    anchor: { pageIndex: 0, x: 9, width: 0 },
  });
  expect(state.anchorPositions).toEqual([9]);
});

test('keeps presented and pending-input checks on mirrored proposal geometry', () => {
  const state = mirroredGeometry();
  state.setPending(true);
  expect(state.geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
  state.setPending(false);
  state.setPresented(false);
  expect(state.geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' })).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
});
