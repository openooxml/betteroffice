import { createDisplayListQueries, effectiveZoom, loadRustDisplayListQueryEngine, type DisplayListQueries, type DisplayListRect } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext, type DisplayListPageProjector } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { ProposalGeometryTarget } from '@betteroffice/docx/yrs';
import { placeRepeated, toOverlayRect } from '../../../plugins/geometry';
import type { DocxAnchorGeometryResult, DocxAnchorRect, DocxGeometryTarget, DocxPluginGeometry, DocxPluginLayout, DocxPluginNavigation, DocxPluginSelection } from '../../../plugins/types';
import type { createPeerlessShellController } from './shellController';
import type { PeerlessFrameGeometry } from './viewStore';
import { clientPointToPage } from './pagePoint';

export function createWorkerPluginGeometry(controller: ReturnType<typeof createPeerlessShellController>, pages: () => HTMLElement | null, layer: () => HTMLElement | null, changed: () => void,
  navigate?: (target: { story: string; paraId: string }, focus: boolean) => Promise<boolean>) {
  let current: { frame: PeerlessFrameGeometry; version: number; layout: DocxPluginLayout; geometry: DocxPluginGeometry; queries: DisplayListQueries } | null = null;
  let displaySelection: DocxPluginSelection['displayRange'] = null;
  let displayWindow: readonly [number, number] | null = null;
  const targets = new Map<string, ProposalGeometryTarget | null>();
  let generation = 0;
  const listeners = new Set<() => void>();
  const publish = () => { changed(); for (const listener of listeners) listener(); };
  const keyOf = (target: DocxGeometryTarget) => JSON.stringify(target);
  const valid = (layout: DocxPluginLayout, frame: PeerlessFrameGeometry) => {
    const view = controller.getSnapshot().view;
    const review = controller.reviewSnapshot();
    return current?.layout === layout && view?.geometry === frame && view.presentedThroughSeq >= view.throughSeq && view.version === current.version && controller.ref.getZoom() === layout.zoom &&
      review?.version === layout.version && review.previewVersion === layout.previewVersion;
  };
  const requestTarget = (target: DocxGeometryTarget) => {
    const source = current;
    if (!source || target.kind === 'proposal') return;
    const key = keyOf(target);
    if (targets.has(key)) return;
    targets.set(key, null);
    void controller.readAt({ kind: 'anchorTargets', targets: [target] }, source.layout.version).then(results => {
      if (current !== source) return;
      targets.set(key, results[0] ?? null); publish();
    }).catch(() => { if (current === source) targets.delete(key); });
  };
  const update = async () => {
    const request = ++generation;
    const view = controller.getSnapshot().view;
    const frame = view?.geometry;
    const container = pages();
    const overlay = layer();
    if (!view || !frame || !container || !overlay || view.presentedThroughSeq < view.throughSeq) return;
    const zoom = controller.ref.getZoom();
    if (current?.frame === frame && current.layout.zoom === zoom) return;
    const native = await loadRustDisplayListQueryEngine();
    if (request !== generation || controller.getSnapshot().view?.geometry !== frame) return;
    const review = await controller.review();
    if (request !== generation || controller.getSnapshot().view?.geometry !== frame) return;
    const queries = createDisplayListQueries({ pages: frame.pages.map(page => page.displayPage!).filter(Boolean) }, native);
    const layout: DocxPluginLayout = { id: `${view.identity.session}:${view.version}:${frame.frameEpoch}`, version: review.version, previewVersion: review.previewVersion, zoom, pageCount: frame.pages.length };
    const projector: DisplayListPageProjector = {
      projectRect(rect) {
        const page = frame.pages[rect.pageIndex]?.bounds;
        const host = container.querySelector<HTMLElement>('[data-peerless-pages-host]')?.getBoundingClientRect();
        if (!page || !host) return null;
        const origin = container.getBoundingClientRect();
        const scale = zoom * effectiveZoom(container);
        return { x: (host.left - origin.left) / scale + page.x + rect.x, y: (host.top - origin.top) / scale + page.y + rect.y, width: rect.width, height: rect.height };
      },
      getPageBounds(pageIndex) {
        const bounds = queries.pageBounds(pageIndex);
        return bounds && projector.projectRect(bounds);
      },
    };
    const dom = createRenderedDomContext(container, zoom, { displayListQueries: queries, projector });
    const shown = () => valid(layout, frame);
    const unavailable = () => ({ ok: false as const, failure: { code: 'layout-unavailable' as const, message: 'No rendered layout shows this target yet' } });
    const project = (rect: DisplayListRect): DocxAnchorRect | null => { const value = projector.projectRect(rect); return value ? { ...toOverlayRect(container, overlay, zoom, value), pageIndex: rect.pageIndex } : null; };
    const repeated = {
      queries, unavailable,
      complete: () => !frame.layoutProvisional && frame.pages.every(({ displayPage }) => displayPage?.unbuilt !== true),
      pending: (pageIndex: number) => !!displayWindow && pageIndex >= displayWindow[0] && pageIndex < displayWindow[1] &&
        queries.displayList?.pages[pageIndex]?.unbuilt === true,
      respond(rects: readonly DisplayListRect[], end: DisplayListRect | null): DocxAnchorGeometryResult {
        const projected = rects.map(project);
        const anchor = end && project(end);
        const bounds = anchor && projector.getPageBounds(anchor.pageIndex);
        if (!anchor || !bounds || projected.some(rect => rect === null)) return unavailable();
        return { ok: true, version: layout.version, previewVersion: layout.previewVersion, layoutId: layout.id, rects: projected as DocxAnchorRect[], anchor, pageRect: toOverlayRect(container, overlay, zoom, bounds) };
      },
    };
    const geometry: DocxPluginGeometry = {
      layout, dom,
      toOverlayRect: rect => current?.geometry === geometry ? toOverlayRect(container, overlay, zoom, rect) : null,
      getPositionAtPoint: () => null,
      async readPositionAtPoint(x, y) {
        if (!shown()) return null;
        const host = container.querySelector<HTMLElement>('[data-peerless-pages-host]');
        const bounds = host?.getBoundingClientRect();
        if (!bounds?.width || !bounds.height) return null;
        const point = clientPointToPage(x, y, bounds, layout.zoom, frame.pages);
        if (!point || point.x < 0 || point.y < 0 || point.x > point.page.bounds.width || point.y > point.page.bounds.height) return null;
        const result = await controller.readAt({ kind: 'pointPosition', expectVersion: review.version,
          hit: { position: -1, pageIndex: point.page.page.pageIndex, region: 'body' }, nativePoint: {
            page: point.page.page, x: point.x, y: point.y, frameEpoch: frame.frameEpoch, layoutRevision: frame.layoutRevision,
            clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false },
          } }, review.version);
        return shown() && result ? { ...result, layoutId: layout.id } : null;
      },
      getAnchorGeometry(target) {
        if (!shown()) return unavailable();
        const expected = target.kind === 'range' ? target.version : (target.kind === 'paragraph' || target.kind === 'search') && 'version' in target.paragraph ? target.paragraph.version : layout.version;
        if (expected !== layout.version) return { ok: false, failure: { code: 'stale-version', message: 'The document changed after that version' } };
        if (target.kind !== 'proposal') requestTarget(target);
        const resolved = target.kind === 'proposal' ? review.proposalGeometry.targets[target.id] : targets.get(keyOf(target));
        if (target.kind === 'proposal' && !resolved) return { ok: false, failure: { code: 'unknown-proposal', message: 'No proposal is registered' } };
        if (!resolved) return unavailable();
        if (!resolved.ok) return resolved;
        if (resolved.root) {
          return resolved.rIds ? placeRepeated(repeated, resolved.rIds, [...resolved.ranges], resolved.hidden ?? [], resolved.paragraph)
            : { ok: false, failure: { code: 'unsupported', message: 'The target has no display position' } };
        }
        const raw = target.kind === 'proposal' ? review.proposalHighlights[target.id] ?? [] : resolved.ranges.flatMap(range => queries.rangeRects(range.from, range.to));
        const rects = raw.map(project).filter(value => value !== null);
        const end = raw.at(-1);
        const anchorRect = end ? { ...end, x: end.x + end.width, width: 0 } : (resolved.paragraph === null ? null : queries.anchorRect(resolved.paragraph));
        const anchor = anchorRect && project(anchorRect);
        const bounds = anchor && projector.getPageBounds(anchor.pageIndex);
        if (!anchor || !bounds) return unavailable();
        return { ok: true, version: layout.version, previewVersion: layout.previewVersion, layoutId: layout.id, rects, anchor, pageRect: toOverlayRect(container, overlay, zoom, bounds) };
      },
    };
    current?.queries.dispose();
    current = { frame, version: view.version, layout, geometry, queries }; targets.clear(); displaySelection = null; publish();
    const selection = view.workerSelection;
    if (selection) {
      const accepted = async (point: typeof selection.anchor) => ({ paraId: point.paraId,
        offset: (await controller.readAt({ kind: 'rawSelectionInfo', selection: { anchor: point, focus: point } }, review.version)).before.length });
      const start = await accepted(selection.anchor);
      const end = await accepted(selection.focus);
      if (current?.layout !== layout) return;
      const result = await controller.readAt({ kind: 'anchorTargets', targets: [{ kind: 'range', version: review.version, range: {
        story: selection.anchor.story, start, end, view: 'accepted',
      } }] }, review.version);
      const mapped = result[0];
      if (current?.layout === layout && mapped?.ok && mapped.ranges[0]) {
        displaySelection = { story: selection.anchor.story, from: mapped.ranges[0].from, to: mapped.ranges.at(-1)!.to, layoutId: layout.id }; publish();
      }
    }
  };
  const navigation: DocxPluginNavigation = {
    async scrollToParagraph(target, options) {
      const token = await controller.ref.readParagraphs({ view: 'accepted', paraIds: [] });
      if (!token.ok) return { ok: false, failure: { code: 'missing-target', message: token.failure.message } };
      if (token.version !== options.expectVersion) return { ok: false, failure: { code: 'stale-version', message: 'The document changed after that version' } };
      const resolved = await controller.readAt({ kind: 'navigationTarget', ...target }, token.version);
      if (typeof resolved === 'string') return { ok: false, failure: { code: resolved === 'ambiguous-target' ? 'ambiguous-target' : resolved === 'unsupported' ? 'unsupported' : 'missing-target', message: resolved } };
      if (navigate) {
        if (!await navigate(target, options.focus === true)) return { ok: false, failure: { code: 'layout-unavailable', message: 'The paragraph could not be presented' } };
        if (options.focus) controller.ref.focus();
        return { ok: true };
      }
      await controller.ref.whenLayoutComplete();
      await update();
      const source = current;
      if (!source || !valid(source.layout, source.frame)) return { ok: false, failure: { code: 'layout-unavailable', message: 'No rendered layout shows the paragraph' } };
      if (source.layout.version !== token.version) return { ok: false, failure: { code: 'stale-version', message: 'The document changed after that version' } };
      const rect = source.queries.anchorRect(resolved.position);
      const page = rect && source.frame.pages.find(entry => entry.page.pageIndex === rect.pageIndex);
      if (!rect || !page) return { ok: false, failure: { code: 'layout-unavailable', message: 'The paragraph page is not laid out' } };
      const container = source.geometry.dom.pagesContainer;
      const scroller = container.closest<HTMLElement>('.docx-editor__scroll-container');
      if (!scroller) return { ok: false, failure: { code: 'layout-unavailable', message: 'The scroll container is unavailable' } };
      scroller.scrollTop += container.getBoundingClientRect().top - scroller.getBoundingClientRect().top + (page.bounds.y + rect.y) * source.layout.zoom - 24;
      if (options.focus) controller.ref.focus();
      return { ok: true };
    },
  };
  return {
    update,
    displaySelection: () => displaySelection,
    geometry: () => current?.geometry ?? null,
    layout: () => current && valid(current.layout, current.frame) ? current.layout : null,
    requestTarget,
    navigation,
    /** Pages in `window` are shown; an unbuilt one there holds the anchors it paints until it is built. */
    setDisplayWindow(window: readonly [number, number]) {
      if (displayWindow?.[0] === window[0] && displayWindow[1] === window[1]) return;
      displayWindow = window; publish();
    },
    subscribe(listener: () => void) { listeners.add(listener); return () => { listeners.delete(listener); }; },
    dispose() { generation++; current?.queries.dispose(); current = null; displaySelection = null; targets.clear(); listeners.clear(); },
  };
}
