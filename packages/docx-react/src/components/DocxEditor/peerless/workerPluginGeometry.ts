import { createDisplayListQueries, effectiveZoom, loadRustDisplayListQueryEngine, type DisplayListQueries, type DisplayListRect } from '@betteroffice/docx/layout/render';
import { createRenderedDomContext, type DisplayListPageProjector } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import { anchorFailure, type DocxReservationsPayload, type ProposalGeometryTarget } from '@betteroffice/docx/yrs';
import { hiddenBoundary, overlayProjector, pageBuilt, pageRectOf, placeRepeated, placeReservation, registerReservationReader, toOverlayRect, unbuiltPagesOf } from '../../../plugins/geometry';
import type { DocxAnchorGeometryResult, DocxAnchorRect, DocxGeometryTarget, DocxPluginGeometry, DocxPluginRect, DocxPluginLayout, DocxPluginNavigation, DocxPluginRefusal, DocxPluginSelection } from '../../../plugins/types';
import { pluginRefusal } from '../../../plugins/createPluginClients';
import { WorkerReadSupersededError, type createPeerlessShellController } from './shellController';
import type { PeerlessFrameGeometry, PeerlessViewIdentity, PeerlessViewSnapshot } from './viewStore';
import { clientPointToPage } from './pagePoint';
import { reviewVersionOf } from './workerReview';

export function workerDocumentVersion(view: PeerlessViewSnapshot | null | undefined): string {
  return view?.editState?.documentVersion ?? view?.geometry?.documentVersion ?? String(view?.version ?? 0);
}

/** The refusal a plugin read or navigation answers when the worker read rejected with `error`, or null to rethrow it. */
export function workerReadRefusal(controller: Pick<ReturnType<typeof createPeerlessShellController>, 'getSnapshot'>, error: unknown): DocxPluginRefusal | null {
  if (error instanceof WorkerReadSupersededError) return controller.getSnapshot().status === 'ready'
    ? pluginRefusal('input-failed', error.message) : pluginRefusal('document-replaced');
  return error instanceof Error && error.message === 'Peer-less editor is not ready'
    ? pluginRefusal('document-replaced') : null;
}

/** `reservations` is the set the editor sends, which a frame's boxes answer for only once it was laid out with it. */
export function createWorkerPluginGeometry(controller: ReturnType<typeof createPeerlessShellController>, pages: () => HTMLElement | null, layer: () => HTMLElement | null, changed: () => void,
  navigate?: (target: { story: string; paraId: string }, focus: boolean, expectVersion: string) => Promise<boolean>, reservations: () => DocxReservationsPayload | null = () => null) {
  type Read = Exclude<DocxGeometryTarget, { kind: 'proposal' | 'reservation' }>;
  type ReadAll = (targets: readonly DocxGeometryTarget[], retry: boolean, owner?: string) => Promise<DocxAnchorGeometryResult[]>;
  let current: { frame: PeerlessFrameGeometry; version: number; pixelRatio: number; layout: DocxPluginLayout; geometry: DocxPluginGeometry; queries: DisplayListQueries;
    readAll: ReadAll } | null = null;
  let repaint: { frame: PeerlessFrameGeometry; next: PeerlessFrameGeometry; same: boolean } | null = null;
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
  const sameLayout = (frame: PeerlessFrameGeometry, next: PeerlessFrameGeometry) => next.documentVersion === frame.documentVersion &&
    next.layoutEpoch !== undefined && next.layoutEpoch === frame.layoutEpoch && next.layoutRevision === frame.layoutRevision && next.documentWidth === frame.documentWidth && !next.layoutProvisional === !frame.layoutProvisional &&
    next.reservations?.version === frame.reservations?.version && next.reservations?.key === frame.reservations?.key && next.pages.length === frame.pages.length &&
    next.pages.every(({ page, bounds }, index) => {
      const built = frame.pages[index]!;
      return page.pageId === built.page.pageId && bounds.x === built.bounds.x && bounds.y === built.bounds.y && bounds.width === built.bounds.width && bounds.height === built.bounds.height;
    });
  /**
   * A selection ack or a repaint of the built frame's document version, preview and layout keeps answering from it until its
   * rebuild lands; the review read waits for that frame, so the worker's reported versions stand in for it.
   */
  const repaints = (layout: DocxPluginLayout, frame: PeerlessFrameGeometry, view: PeerlessViewSnapshot) => {
    const next = view.geometry;
    const reported = reviewVersionOf(view);
    if (!next || frame.documentVersion === undefined || workerDocumentVersion(view) !== frame.documentVersion ||
      reported.version !== layout.version || reported.previewVersion !== layout.previewVersion) return false;
    if (next === frame) return true;
    if (repaint?.frame !== frame || repaint.next !== next) repaint = { frame, next, same: sameLayout(frame, next) };
    return repaint.same && globalThis.devicePixelRatio === current?.pixelRatio;
  };
  const valid = (layout: DocxPluginLayout, frame: PeerlessFrameGeometry, presented = false) => {
    const snapshot = controller.getSnapshot();
    const view = snapshot.view;
    const review = controller.describedReview();
    const described = review?.version === layout.version && review.previewVersion === layout.previewVersion;
    return snapshot.status === 'ready' && !!view && sameView(view.identity, frame.identity) && current?.layout === layout && view.version === current.version &&
      controller.ref.getZoom() === layout.zoom && (view.geometry === frame && view.presentedThroughSeq >= view.throughSeq ? described
        : !presented && (described || !review) && repaints(layout, frame, view));
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
    if (!source || target.kind === 'proposal' || target.kind === 'reservation') return;
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
  const clear = () => {
    generation++;
    const hadGeometry = !!current || !!answers;
    current?.queries.dispose(); current = null; answers = null; displaySelection = null; repaint = null;
    if (hadGeometry) publish();
  };
  const update = async () => {
    const snapshot = controller.getSnapshot();
    const view = snapshot.view;
    if (snapshot.status !== 'ready' || !view || (current && !sameView(current.frame.identity, view.identity))) clear();
    const request = ++generation;
    const frame = view?.geometry;
    const container = pages();
    const overlay = layer();
    if (snapshot.status !== 'ready' || !view || !frame || !container || !overlay || !sameView(view.identity, frame.identity) || view.presentedThroughSeq < view.throughSeq) return;
    const zoom = controller.ref.getZoom();
    if (current?.frame === frame && current.layout.zoom === zoom) return;
    const native = await loadRustDisplayListQueryEngine();
    if (request !== generation || controller.getSnapshot().view?.geometry !== frame) return;
    const review = controller.describedReview() ?? await controller.review();
    if (request !== generation || controller.getSnapshot().view?.geometry !== frame) return;
    const queries = createDisplayListQueries({ pages: frame.pages.map(page => page.displayPage!).filter(Boolean) }, native);
    const layout: DocxPluginLayout = { id: `${view.identity.session}:${view.version}:${frame.frameEpoch}`, version: review.version, previewVersion: review.previewVersion,
      reservationsVersion: frame.reservations?.version ?? null, zoom, pageCount: frame.pages.length };
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
    /** `owner`'s reservation `id` as this frame lays it out: its boxes, or the anchor a `hide` leaves a boundary at. */
    const reserved = (owner: string, id: string, measured: ReturnType<typeof measure>): { result: DocxAnchorGeometryResult } | { hide: Read } => {
      if (!shown()) return { result: unavailable() };
      const sent = reservations();
      if (sent ? sent.key === undefined || frame.reservations?.key !== sent.key : frame.reservations !== undefined) return { result: unavailable() };
      const item = sent?.sets.find(set => set.owner === owner)?.items.find(entry => entry.id === id);
      if (!item) return { result: anchorFailure('missing-target', 'The reservation did not resolve') };
      return item.kind === 'hide' ? { hide: item.anchor } : { result: placeReservation(repeatedOf(measured), owner, id) };
    };
    const answer = (target: DocxGeometryTarget, request: boolean, measured = measure(), owner?: string, hide = false): DocxAnchorGeometryResult => {
      if (target.kind === 'reservation') {
        if (owner === undefined) return anchorFailure('unsupported', 'Reservations are read through a plugin context');
        const item = reserved(owner, target.id, measured);
        return 'hide' in item ? answer(item.hide, request, measured, undefined, true) : item.result;
      }
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
      if (hide) {
        const placement = repeatedOf(measured);
        const at = hiddenBoundary(queries, resolved.ranges, [...(review.proposalGeometry?.hidden ?? []), ...resolved.ranges], resolved.paragraph);
        return at && placement.pending(at.pageIndex) ? unavailable() : placement.respond([], at, unbuiltPagesOf(queries, resolved.ranges));
      }
      const raw = resolved.ranges.flatMap(range => queries.rangeRects(range.from, range.to));
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
        if (!valid(layout, frame, true)) return null;
        const host = container.querySelector<HTMLElement>('[data-peerless-pages-host]');
        const bounds = host?.getBoundingClientRect();
        if (!bounds?.width || !bounds.height) return null;
        const point = clientPointToPage(x, y, bounds, layout.zoom, frame.pages);
        if (!point || point.x < 0 || point.y < 0 || point.x > point.page.bounds.width || point.y > point.page.bounds.height) return null;
        try {
          const result = await controller.readAt({ kind: 'pointPosition', expectVersion: review.version,
            hit: { position: -1, pageIndex: point.page.page.pageIndex, region: 'body' }, nativePoint: {
              page: point.page.page, x: point.x, y: point.y, frameEpoch: frame.frameEpoch, layoutRevision: frame.layoutRevision,
              clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false },
            } }, review.version);
          return valid(layout, frame, true) && result ? { ...result, layoutId: layout.id } : null;
        } catch (error) {
          if (workerReadRefusal(controller, error)) return null;
          throw error;
        }
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
    const readAll: ReadAll = async (targets, retry, owner) => {
      const sameDocument = () => !!current && sameView(current.frame.identity, frame.identity);
      if (!sameDocument()) return targets.map(unavailable);
      const asked = measure();
      await Promise.all(targets.map(target => {
        const item = target.kind === 'reservation' && owner !== undefined ? reserved(owner, target.id, asked) : null;
        return item && 'hide' in item ? requestTarget(item.hide, retry) : requestTarget(target, retry);
      }));
      if (!sameDocument()) return targets.map(unavailable);
      if (current!.geometry !== geometry) return current!.readAll(targets, false, owner);
      const measured = measure();
      return targets.map(target => answer(target, false, measured, owner));
    };
    registerReservationReader(geometry, {
      get: (owner, id) => answer({ kind: 'reservation', id }, true, measure(), owner),
      read: (owner, targets) => readAll(targets, true, owner),
    });
    current?.queries.dispose();
    current = { frame, version: view.version, pixelRatio: globalThis.devicePixelRatio, layout, geometry, queries, readAll }; displaySelection = null; publish();
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
      const identity = controller.getSnapshot().view?.identity;
      const check = () => {
        const snapshot = controller.getSnapshot();
        const view = snapshot.view;
        if (snapshot.status !== 'ready' || !identity || !view || !sameView(identity, view.identity)) throw new WorkerReadSupersededError('Document was replaced');
        if (workerDocumentVersion(view) !== options.expectVersion) throw new WorkerReadSupersededError('Worker read was superseded');
      };
      const token = await controller.ref.readParagraphs({ view: 'accepted', paraIds: [] });
      if (!token.ok) return { ok: false, failure: { code: 'missing-target', message: token.failure.message } };
      if (token.version !== options.expectVersion) return { ok: false, failure: { code: 'stale-version', message: 'The document changed after that version' } };
      check();
      const resolved = await controller.readAt({ kind: 'navigationTarget', ...target }, token.version);
      check();
      if (typeof resolved === 'string') return { ok: false, failure: { code: resolved === 'ambiguous-target' ? 'ambiguous-target' : resolved === 'unsupported' ? 'unsupported' : 'missing-target', message: resolved } };
      if (navigate) {
        const revealed = await navigate(target, options.focus === true, token.version);
        check();
        if (!revealed) return { ok: false, failure: { code: 'layout-unavailable', message: 'The paragraph could not be presented' } };
        if (options.focus) controller.ref.focus();
        return { ok: true };
      }
      await controller.ref.whenLayoutComplete();
      check();
      await update();
      check();
      const source = current;
      if (!source || !valid(source.layout, source.frame)) return { ok: false, failure: { code: 'layout-unavailable', message: 'No rendered layout shows the paragraph' } };
      if (source.layout.version !== token.version) return { ok: false, failure: { code: 'stale-version', message: 'The document changed after that version' } };
      const rect = source.queries.anchorRect(resolved.position);
      const page = rect && source.frame.pages.find(entry => entry.page.pageIndex === rect.pageIndex);
      if (!rect || !page) return { ok: false, failure: { code: 'layout-unavailable', message: 'The paragraph page is not laid out' } };
      const container = source.geometry.dom.pagesContainer;
      const scroller = container.closest<HTMLElement>('.docx-editor__scroll-container');
      if (!scroller) return { ok: false, failure: { code: 'layout-unavailable', message: 'The scroll container is unavailable' } };
      check();
      scroller.scrollTop += container.getBoundingClientRect().top - scroller.getBoundingClientRect().top + (page.bounds.y + rect.y) * source.layout.zoom - 24;
      if (options.focus) controller.ref.focus();
      return { ok: true };
    },
  };
  return {
    update,
    clear,
    displaySelection: () => displaySelection,
    geometry: () => {
      const snapshot = controller.getSnapshot();
      return current && snapshot.status === 'ready' && snapshot.view && sameView(current.frame.identity, snapshot.view.identity) ? current.geometry : null;
    },
    layout: () => current && valid(current.layout, current.frame) ? current.layout : null,
    requestTarget,
    navigation,
    /** Pages in `window` are shown; an unbuilt one there holds the anchors it paints until it is built. */
    setDisplayWindow(window: readonly [number, number]) {
      if (displayWindow?.[0] === window[0] && displayWindow[1] === window[1]) return;
      displayWindow = window; publish();
    },
    subscribe(listener: () => void) { listeners.add(listener); return () => { listeners.delete(listener); }; },
    dispose() { generation++; current?.queries.dispose(); current = null; displaySelection = null; answers = null; repaint = null; listeners.clear(); },
  };
}
