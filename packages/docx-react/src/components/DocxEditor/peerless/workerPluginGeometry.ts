import { createDisplayListQueries, effectiveZoom, loadRustDisplayListQueryEngine, type DisplayListQueries, type DisplayListRect } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext, type DisplayListPageProjector } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { ProposalGeometryTarget } from '@betteroffice/docx/yrs';
import { overlayProjector, pageBuilt, pageRectOf, placeRepeated, toOverlayRect, unbuiltPagesOf } from '../../../plugins/geometry';
import type { DocxAnchorGeometryResult, DocxAnchorRect, DocxGeometryTarget, DocxPluginGeometry, DocxPluginRect, DocxPluginLayout, DocxPluginNavigation, DocxPluginSelection } from '../../../plugins/types';
import { WorkerReadSupersededError, type createPeerlessShellController } from './shellController';
import type { PeerlessFrameGeometry, PeerlessViewIdentity } from './viewStore';
import { clientPointToPage } from './pagePoint';

export function createWorkerPluginGeometry(controller: ReturnType<typeof createPeerlessShellController>, pages: () => HTMLElement | null, layer: () => HTMLElement | null, changed: () => void,
  navigate?: (target: { story: string; paraId: string }, focus: boolean) => Promise<boolean>) {
  type Read = Exclude<DocxGeometryTarget, { kind: 'proposal' }>;
  type ReadAll = (targets: readonly DocxGeometryTarget[], retry: boolean) => Promise<DocxAnchorGeometryResult[]>;
  let current: { frame: PeerlessFrameGeometry; version: number; layout: DocxPluginLayout; geometry: DocxPluginGeometry; queries: DisplayListQueries; readAll: ReadAll } | null = null;
  let displaySelection: DocxPluginSelection['displayRange'] = null;
  let displayWindow: readonly [number, number] | null = null;
  interface Answers { key: string; version: string; resolved: Map<string, ProposalGeometryTarget>; requested: Map<string, Promise<boolean>>; failed: Set<string>;
    batch: { targets: Map<string, Read>; read: Promise<boolean> } | null }
  let answers: Answers | null = null;
  let generation = 0;
  const listeners = new Set<() => void>();
  const publish = () => { changed(); for (const listener of listeners) listener(); };
  const keyOf = (target: DocxGeometryTarget) => JSON.stringify(target);
  const sameView = (a: PeerlessViewIdentity, b: PeerlessViewIdentity) => a.session === b.session && a.incarnation === b.incarnation && a.epoch === b.epoch;
  const pinnedVersion = (target: DocxGeometryTarget) => target.kind === 'range' ? target.version
    : (target.kind === 'paragraph' || target.kind === 'search') && 'version' in target.paragraph ? target.paragraph.version : undefined;
  /** Worker answers are display positions of one document version and preview, whatever the zoom. */
  const answersOf = ({ identity }: PeerlessFrameGeometry, layout: DocxPluginLayout) => {
    const key = `${identity.session}|${identity.incarnation}|${identity.epoch}|${layout.version}|${layout.previewVersion}`;
    if (answers?.key !== key) answers = { key, version: layout.version, resolved: new Map(), requested: new Map(), failed: new Set(), batch: null };
    return answers;
  };
  const valid = (layout: DocxPluginLayout, frame: PeerlessFrameGeometry) => {
    const view = controller.getSnapshot().view;
    const review = controller.reviewSnapshot();
    return current?.layout === layout && view?.geometry === frame && view.presentedThroughSeq >= view.throughSeq && view.version === current.version && controller.ref.getZoom() === layout.zoom &&
      review?.version === layout.version && review.previewVersion === layout.previewVersion;
  };
  const resolveBatch = (entry: Answers) => {
    const { targets } = entry.batch!;
    entry.batch = null;
    const keys = [...targets.keys()];
    const asked = controller.getSnapshot().view;
    /** A read the view moved past was superseded, not failed, so the next ask reads again. */
    const unmoved = () => { const view = controller.getSnapshot().view; return !!view && !!asked && view.version === asked.version && sameView(view.identity, asked.identity); };
    return controller.readAt({ kind: 'anchorTargets', targets: [...targets.values()] }, entry.version).then(results => {
      keys.forEach((key, index) => { if (results[index]) entry.resolved.set(key, results[index]); });
      return true;
    }, error => { if (!(error instanceof WorkerReadSupersededError) && unmoved()) for (const key of keys) entry.failed.add(key); return false; }).finally(() => {
      for (const key of keys) entry.requested.delete(key);
      if (answers === entry) publish();
    });
  };
  /**
   * Starts the worker read of `target` with the others asked in this tick; resolves once it lands. A target pinned to
   * another version is never read, and one whose read failed is read again only on `retry`.
   */
  const requestTarget = (target: DocxGeometryTarget, retry = false): Promise<boolean> | undefined => {
    const source = current;
    if (!source || target.kind === 'proposal') return;
    const entry = answersOf(source.frame, source.layout);
    const pinned = pinnedVersion(target);
    if (pinned !== undefined && pinned !== entry.version) return;
    const key = keyOf(target);
    if (entry.resolved.has(key) || entry.requested.has(key)) return entry.requested.get(key);
    if (entry.failed.has(key) && !retry) return;
    entry.failed.delete(key);
    entry.batch ??= { targets: new Map(), read: Promise.resolve().then(() => resolveBatch(entry)) };
    entry.batch.targets.set(key, target);
    entry.requested.set(key, entry.batch.read);
    return entry.batch.read;
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
    /** Projections over the pages host and overlay layer, each measured at most once. */
    const measure = () => {
      let host: { x: number; y: number } | null | undefined;
      let layerOf: ((rect: DocxPluginRect) => DocxPluginRect) | undefined;
      const local = (rect: DisplayListRect) => {
        const page = frame.pages[rect.pageIndex]?.bounds;
        if (host === undefined) {
          const hostRect = container.querySelector<HTMLElement>('[data-peerless-pages-host]')?.getBoundingClientRect();
          const origin = hostRect && container.getBoundingClientRect();
          const scale = zoom * effectiveZoom(container);
          host = hostRect && origin ? { x: (hostRect.left - origin.left) / scale, y: (hostRect.top - origin.top) / scale } : null;
        }
        return page && host ? { x: host.x + page.x + rect.x, y: host.y + page.y + rect.y, width: rect.width, height: rect.height } : null;
      };
      const toLayer = (rect: DocxPluginRect) => (layerOf ??= overlayProjector(container, overlay, zoom))(rect);
      const pageBounds = (pageIndex: number) => { const bounds = queries.pageBounds(pageIndex); return bounds && local(bounds); };
      const project = (rect: DisplayListRect): DocxAnchorRect | null => { const value = local(rect); return value ? { ...toLayer(value), pageIndex: rect.pageIndex } : null; };
      return { local, toLayer, pageBounds, project };
    };
    const projector: DisplayListPageProjector = { projectRect: rect => measure().local(rect), getPageBounds: pageIndex => measure().pageBounds(pageIndex) };
    const dom = createRenderedDomContext(container, zoom, { displayListQueries: queries, projector });
    const shown = () => valid(layout, frame);
    const unavailable = () => ({ ok: false as const, failure: { code: 'layout-unavailable' as const, message: 'No rendered layout shows this target yet' } });
    const pending = { ok: false as const, failure: { code: 'pending' as const, message: 'The target geometry is being read' } };
    const readFailed = { ok: false as const, failure: { code: 'read-failed' as const, message: 'The worker could not read the target geometry' } };
    const repeatedOf = ({ project, pageBounds, toLayer }: ReturnType<typeof measure>) => ({
      queries, unavailable,
      complete: () => !frame.layoutProvisional && frame.pages.every(({ displayPage }) => displayPage?.unbuilt !== true),
      pending: (pageIndex: number) => !!displayWindow && pageIndex >= displayWindow[0] && pageIndex < displayWindow[1] &&
        queries.displayList?.pages[pageIndex]?.unbuilt === true,
      respond(rects: readonly DisplayListRect[], end: DisplayListRect | null, unbuiltPages: readonly number[] = []): DocxAnchorGeometryResult {
        const projected = rects.map(project);
        const anchor = end && project(end);
        const bounds = anchor && pageBounds(anchor.pageIndex);
        if (!anchor || !bounds || projected.some(rect => rect === null)) return unavailable();
        return { ok: true, version: layout.version, previewVersion: layout.previewVersion, layoutId: layout.id, rects: projected as DocxAnchorRect[], anchor,
          pageRect: toLayer(bounds), ...(unbuiltPages.length ? { unbuiltPages } : {}) };
      },
    });
    const answer = (target: DocxGeometryTarget, request: boolean, measured = measure()): DocxAnchorGeometryResult => {
      if (!shown()) return unavailable();
      if ((pinnedVersion(target) ?? layout.version) !== layout.version) return { ok: false, failure: { code: 'stale-version', message: 'The document changed after that version' } };
      if (request) requestTarget(target);
      const entry = answersOf(frame, layout);
      const resolved = target.kind === 'proposal' ? review.proposalGeometry.targets[target.id] : entry.resolved.get(keyOf(target));
      if (target.kind === 'proposal' && !resolved) return { ok: false, failure: { code: 'unknown-proposal', message: 'No proposal is registered' } };
      if (!resolved) return entry.failed.has(keyOf(target)) ? readFailed : request ? pending : unavailable();
      if (!resolved.ok) return resolved;
      if (resolved.root) {
        return resolved.rIds ? placeRepeated(repeatedOf(measured), resolved.rIds, [...resolved.ranges], resolved.hidden ?? [], resolved.paragraph)
          : { ok: false, failure: { code: 'unsupported', message: 'The target has no display position' } };
      }
      const raw = target.kind === 'proposal' ? review.proposalHighlights[target.id] ?? [] : resolved.ranges.flatMap(range => queries.rangeRects(range.from, range.to));
      const rects = raw.map(measured.project).filter(value => value !== null);
      const end = raw.at(-1);
      const anchorRect = end ? { ...end, x: end.x + end.width, width: 0 } : (resolved.paragraph === null ? null : queries.anchorRect(resolved.paragraph));
      const anchor = anchorRect && measured.project(anchorRect);
      const bounds = anchor && measured.pageBounds(anchor.pageIndex);
      if (!anchor || !bounds) return unavailable();
      const unbuiltPages = unbuiltPagesOf(queries, resolved.ranges);
      return { ok: true, version: layout.version, previewVersion: layout.previewVersion, layoutId: layout.id, rects, anchor, pageRect: measured.toLayer(bounds),
        ...(unbuiltPages.length ? { unbuiltPages } : {}) };
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
      getAnchorGeometry: target => answer(target, true),
      readAnchorGeometries: targets => readAll(targets, true),
      getPageRect(pageIndex) {
        if (!shown()) return null;
        const { pageBounds, toLayer } = measure();
        return pageRectOf(queries, pageIndex, pageBounds(pageIndex), toLayer);
      },
      isPageBuilt: pageIndex => shown() && pageBuilt(queries, pageIndex),
    };
    /** Answers `targets` once their reads land, on the layout then current, until the document is replaced. */
    const readAll: ReadAll = async (targets, retry) => {
      const sameDocument = () => !!current && sameView(current.frame.identity, frame.identity);
      if (!sameDocument()) return targets.map(unavailable);
      await Promise.all(targets.map(target => requestTarget(target, retry)));
      if (!sameDocument()) return targets.map(unavailable);
      if (current!.geometry !== geometry) return current!.readAll(targets, false);
      const measured = measure();
      return targets.map(target => answer(target, false, measured));
    };
    current?.queries.dispose();
    current = { frame, version: view.version, layout, geometry, queries, readAll }; displaySelection = null; publish();
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
    dispose() { generation++; current?.queries.dispose(); current = null; displaySelection = null; answers = null; listeners.clear(); },
  };
}
