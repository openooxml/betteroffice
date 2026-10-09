import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  proposalSetIdentity,
  type DocxProposalSnapshot,
  type ProposalGeometryMirror,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { stampRevisionPreviewKey } from '../components/DocxEditor/internals/layoutProvenance';
import { createPluginGeometry } from './geometry';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

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

test('serves proposal geometry from a mirror while the replica has no anchor reads', () => {
  const { geometry, anchorPositions } = mirroredGeometry();
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
  expect(geometry.getAnchorGeometry({ kind: 'revision', revisionId: 'r1' })).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
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
