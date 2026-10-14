import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxSelectionPointIntent } from './residentCommandProtocol';
import { resolveDocxSelectionIntent } from './residentCommandSelection';
import { acceptedOffsetOf, type DocxResolvedPointPosition } from './pointPosition';
import { retainedFramePageById, type RetainedFrame } from '../layout/render/frameDelta';
import { createDisplayListQueries } from '../layout/render/displayListQueries';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';

export async function readResidentPointer(engine: ResidentEngineSession, frame: RetainedFrame, layoutRevision: number,
  point: DocxSelectionPointIntent, version: string): Promise<DocxResolvedPointPosition | null> {
  if (engine.proposalEngine.version() !== version || point.frameEpoch !== frame.frameEpoch || point.layoutRevision !== layoutRevision ||
    !/^\d+$/u.test(point.page.pageId)) return null;
  const page = retainedFramePageById(frame, BigInt(point.page.pageId));
  if (!page || page.page.unbuilt || point.x < 0 || point.y < 0 || point.x > page.page.width || point.y > page.page.height) return null;
  const native = await loadRustDisplayListQueryEngine();
  const queries = createDisplayListQueries(frame.displayList, native);
  try {
    const pageIndex = frame.pages.indexOf(page);
    const hit = queries.hitTestRegions(pageIndex, point.x, point.y);
    if (!hit || hit.pos == null) return null;
    const basis = { epoch: 1, version: 0 };
    const resolved = resolveDocxSelectionIntent({ engine,
      state: { readOnly: true, selection: null, storedMarks: null, stickyPositions: new Map() },
      command: { session: 'point-read', incarnation: 1, seq: 0, basis, kind: 'selectPoint',
        payload: { ...point, clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false } } },
      context: { basis, currentVersion: 0, replayPreviouslyApplied: false }, readLayout: () => ({ frame, layoutRevision }) });
    const loc = resolved.selection.focus;
    const offset = acceptedOffsetOf({ ...engine.geometryReader, selectionText: engine.selectionText }, loc);
    return offset === null ? null : { position: hit.pos, region: hit.region, pageIndex, version,
      ...(hit.rId ? { rId: hit.rId } : {}), ...(hit.noteId === undefined ? {} : { noteId: hit.noteId }),
      target: { kind: 'range', story: loc.story, start: { paraId: loc.paraId, offset }, end: { paraId: loc.paraId, offset }, view: 'accepted' } };
  } finally { queries.dispose(); }
}
