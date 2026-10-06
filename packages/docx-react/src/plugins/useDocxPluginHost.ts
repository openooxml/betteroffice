import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import {
  createT,
  deepMerge,
  en,
  type LocaleStrings,
  type Translations,
} from '@betteroffice/docx-i18n';
import {
  effectiveZoom,
  resolveDisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import type { RenderedDomContext } from '@betteroffice/docx/plugin-api';
import type { ResidentDocumentReadValues } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import type { ResidentEngineWorkerClient, YrsSession } from '@betteroffice/docx/yrs';
import type { DocxCommandController } from '../commands/createDocxCommandStore';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';
import {
  isPresented,
  onPresented,
  sourceVersionOf,
  presentedWorkerVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import { displayWindowOf } from '../components/DocxEditor/internals/displayWindow';
import { resolvePointPosition } from '../components/DocxEditor/internals/pointPosition';
import { workerProposalAuthority, workerProposalRoundAuthority, hasEditorWorkerProposalRounds, subscribeEditorWorkerProposalAuthority } from '../components/DocxEditor/internals/workerProposalAuthority';
import { workerOpenReplicaReady } from '../components/DocxEditor/internals/workerOpenReplica';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { SelectionState } from '../components/DocxEditor/types';
import type { ViewerSelectionChange } from '../components/DocxEditor/internals/viewerSelectionController';
import type { ReactSidebarItem } from '../plugin-api/types';
import {
  createDocxPluginHost,
  type DocxPluginActivation,
  type DocxPluginHost,
} from './createDocxPluginHost';
import { resolveParagraph } from './createPluginClients';
import { createPluginGeometry, pluginLayout, readPluginPositionAtPoint } from './geometry';
import { managedSidebarItems } from './PluginSidebarItems';
import { currentPreviewKey } from './proposalPreview';
import type {
  DocxEditorPluginProps,
  DocxPluginGeometry,
  DocxPluginLayout,
  DocxPluginSidebarItem,
} from './types';

const NO_ACTIVATIONS: readonly DocxPluginActivation[] = Object.freeze([]);

type RenderedDom = { context: RenderedDomContext; queries: DisplayListQueries };

export interface UseDocxPluginHostOptions extends DocxEditorPluginProps {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  writeModeRef: React.RefObject<EditorMode>;
  mode: EditorMode;
  /** Host `readOnly` or viewing mode. */
  readOnly: boolean;
  viewerSelection?: boolean;
  commands: DocxCommandController;
  /** The authoritative session once a document is ready, else null. */
  session: YrsSession | null;
  /** Changes whenever a new document load starts. */
  loadGeneration: number;
  queries: DisplayListQueries | null;
  viewerDocumentRead?: ResidentEngineWorkerClient['documentRead'];
  layoutError: Error | null;
  zoom: number;
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  overlayTarget: HTMLElement | null;
  selectionChangeSubscribersRef: React.RefObject<Set<(state: SelectionState | null) => void>>;
  i18n: Translations | undefined;
  onRenderedDomContextReady: ((context: RenderedDomContext) => void) | undefined;
}

export interface DocxPluginHostBinding {
  host: DocxPluginHost;
  /** Whether any plugin is installed; raw geometry overrides step aside then. */
  managed: boolean;
  activations: readonly DocxPluginActivation[];
  sidebarItems: ReactSidebarItem[];
  /** The editor's own rendered-DOM context. */
  renderedDomContext: RenderedDomContext | null;
  /** The geometry overlays draw with while the host's own layout is behind or not yet built. */
  heldGeometry: DocxPluginGeometry | null;
  overlayLayerRef: (element: HTMLDivElement | null) => void;
  /** Receives each rendered-DOM context the paged editor builds. */
  onRenderedDomContext(context: RenderedDomContext, queries: DisplayListQueries): void;
  /** Ends every plugin activation as a new document starts loading. */
  beginLoad(): void;
  /** Publishes the current selection to plugins. */
  publishSelection(): void;
  publishViewerSelection(selection: ViewerSelectionChange): void;
}

/** Owns the plugin host of one `DocxEditor` and binds it to the editor's authority. */
export function useDocxPluginHost(options: UseDocxPluginHostOptions): DocxPluginHostBinding {
  const latest = useRef(options);
  latest.current = options;
  const translate = useMemo(() => {
    const merged = deepMerge(
      en as Record<string, unknown>,
      options.i18n as Record<string, unknown> | undefined
    ) as LocaleStrings;
    return createT(merged, typeof options.i18n?._lang === 'string' ? options.i18n._lang : 'en');
  }, [options.i18n]);
  const translateRef = useRef(translate);
  translateRef.current = translate;
  const geometryRef = useRef<DocxPluginGeometry | null>(null);
  const adoptedRef = useRef<{ session: YrsSession; geometry: DocxPluginGeometry } | null>(null);
  const queriesCurrentRef = useRef(false);
  const heldCandidate = (): DocxPluginGeometry | null => {
    const held = adoptedRef.current;
    const { session, layoutError } = latest.current;
    return held && held.session === session && !layoutError ? held.geometry : null;
  };
  const layoutRef = useRef<DocxPluginLayout | null>(null);
  const formattingRef = useRef<SelectionState | null>(null);
  const viewerSelectionRef = useRef<ViewerSelectionChange | null>(null);
  const layoutListeners = useRef(new Set<() => void>());
  const detachAuthority = useRef<(() => void) | null>(null);

  const [host] = useState<DocxPluginHost>(() =>
    createDocxPluginHost({
      pagedEditorRef: options.pagedEditorRef,
      writeMode: () => latest.current.writeModeRef.current ?? 'viewing',
      viewer: () => latest.current.viewerDocumentRead !== undefined,
      commands: () => latest.current.commands,
      translate: (key) => translateRef.current(key),
      geometry: () => geometryRef.current,
      layout() {
        const queries = latest.current.queries;
        const layout = latest.current.pagedEditorRef.current?.getLayout();
        return {
          queries,
          complete:
            !!queries &&
            !!layout &&
            !layout.partial &&
            sourceVersionOf(layout) === sourceVersionOf(queries) &&
            queries.pageCount() === layout.pages.length,
          failed: latest.current.layoutError !== null,
        };
      },
      subscribeLayout(listener) {
        layoutListeners.current.add(listener);
        const unsubscribe = host.subscribe(listener);
        return () => {
          layoutListeners.current.delete(listener);
          unsubscribe();
        };
      },
    })
  );

  const managed = (options.plugins?.length ?? 0) > 0;

  useLayoutEffect(() => {
    for (const listener of layoutListeners.current) listener();
  });

  useEffect(() => {
    const queries = options.queries;
    if (!queries || queries.sourceState().status !== 'loading') return;
    let cancelled = false;
    const ready = () => {
      if (cancelled) return;
      for (const listener of layoutListeners.current) listener();
    };
    void queries.whenReady().then(ready, ready);
    return () => {
      cancelled = true;
    };
  }, [options.queries]);

  useLayoutEffect(() => {
    host.setReporter(options.onPluginError);
  });

  useLayoutEffect(() => {
    host.modeChanged(options.mode, options.readOnly);
  }, [host, options.mode, options.readOnly]);

  useLayoutEffect(() => {
    host.setGrants(options.pluginGrants);
  });

  useLayoutEffect(() => {
    host.setPlugins(options.plugins);
  }, [host, options.plugins]);

  useEffect(() => () => host.close('unmounted'), [host]);

  useEffect(() => {
    if (!options.session) return;
    host.open(options.session);
    const update = () => {
      host.geometryChanged();
      if (layoutRef.current) host.layoutPresented(layoutRef.current);
    };
    const unsubscribe = workerProposalAuthority(options.session)?.subscribe(update);
    const unsubscribeEditor = subscribeEditorWorkerProposalAuthority(options.session, update);
    let attached = true;
    const detach = () => {
      if (!attached) return;
      attached = false;
      unsubscribe?.();
      unsubscribeEditor();
      if (detachAuthority.current === detach) detachAuthority.current = null;
    };
    detachAuthority.current = detach;
    return () => {
      detach();
      host.close('document-replaced');
    };
  }, [host, options.session, options.loadGeneration]);

  useEffect(() => {
    host.syncCommands();
    return host.subscribe(() => host.syncCommands());
  }, [host]);

  const activations = useSyncExternalStore(host.subscribe, host.activations, () => NO_ACTIVATIONS);

  const publishSelection = useCallback(() => {
    if (!latest.current.plugins?.length) return;
    if (latest.current.viewerSelection) {
      host.selectionChanged({ formatting: null, displayRange: viewerSelectionRef.current?.displayRange ?? null });
      return;
    }
    const editor = latest.current.pagedEditorRef.current;
    const range = editor?.getSelectionRange() ?? null;
    const layout = layoutRef.current;
    const story = latest.current.viewerDocumentRead
      ? 'body'
      : editor?.getYrsSession()?.selection()?.head.story ?? 'body';
    host.selectionChanged({
      formatting: formattingRef.current,
      displayRange:
        range && layout ? { story, from: range.from, to: range.to, layoutId: layout.id } : null,
    });
  }, [host]);

  const publishViewerSelection = useCallback((selection: ViewerSelectionChange) => {
    viewerSelectionRef.current = selection;
    if (!latest.current.plugins?.length) return;
    host.selectionChanged({ formatting: null, displayRange: selection.displayRange });
  }, [host]);

  useEffect(() => {
    if (!managed) return;
    const subscribers = options.selectionChangeSubscribersRef.current;
    const listener = (state: SelectionState | null) => {
      formattingRef.current = state;
      publishSelection();
    };
    subscribers.add(listener);
    return () => {
      subscribers.delete(listener);
    };
  }, [managed, options.selectionChangeSubscribersRef, publishSelection]);

  const domRef = useRef<RenderedDom | null>(null);
  const [dom, setDom] = useState<RenderedDom | null>(null);
  const onRenderedDomContext = useCallback(
    (context: RenderedDomContext, queries: DisplayListQueries) => {
      domRef.current = { context, queries };
      setDom(domRef.current);
      for (const listener of layoutListeners.current) listener();
    },
    []
  );
  useEffect(() => {
    if (!dom) return;
    try {
      latest.current.onRenderedDomContextReady?.(dom.context);
    } catch (error) {
      console.error('[DocxEditor] onRenderedDomContextReady threw', error);
    }
  }, [dom]);

  const [layer, setLayer] = useState<HTMLDivElement | null>(null);
  const [moved, setMoved] = useState(0);
  useEffect(() => {
    const pages = options.canvasHostRef.current;
    const target = options.overlayTarget;
    if (!managed || !pages || !target) return;
    let frame = 0;
    const bump = () => {
      if (frame) return;
      frame = requestAnimationFrame(() => {
        frame = 0;
        setMoved((value) => value + 1);
      });
    };
    const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(bump);
    observer?.observe(pages);
    observer?.observe(target);
    pages.addEventListener('transitionend', bump);
    window.addEventListener('resize', bump);
    return () => {
      if (frame) cancelAnimationFrame(frame);
      observer?.disconnect();
      pages.removeEventListener('transitionend', bump);
      window.removeEventListener('resize', bump);
    };
  }, [managed, options.canvasHostRef, options.overlayTarget, options.queries]);

  const version = useSyncExternalStore(host.subscribe, host.version, host.version);
  const previewVersion = useSyncExternalStore(
    host.subscribe,
    host.previewVersion,
    host.previewVersion
  );
  const previewKey = currentPreviewKey(options.session);
  const layout = useMemo(
    () =>
      pluginLayout(options.queries, managed ? version : null, options.zoom, {
        key: previewKey,
        previewVersion,
      }),
    [managed, options.queries, options.zoom, version, previewKey, previewVersion]
  );
  const layoutStable = useRef<DocxPluginLayout | null>(null);
  if (
    layout?.id !== layoutStable.current?.id ||
    layout?.version !== layoutStable.current?.version ||
    layout?.previewVersion !== layoutStable.current?.previewVersion ||
    layout?.zoom !== layoutStable.current?.zoom
  ) {
    layoutStable.current = layout;
  }
  const currentLayout = layoutStable.current;
  layoutRef.current = currentLayout;

  const geometry = useMemo(() => {
    if (!currentLayout || !dom || dom.queries !== options.queries || !layer) return null;
    const shownList = dom.queries.displayList;
    let proposalTarget = false;
    const created: DocxPluginGeometry = createPluginGeometry(
      currentLayout,
      dom.context,
      layer,
      () =>
        host.layoutId() === currentLayout.id &&
        host.previewVersion() === currentLayout.previewVersion &&
        latest.current.zoom === currentLayout.zoom &&
        domRef.current === dom &&
        dom.context.pagesContainer.isConnected,
      (hit) =>
        resolvePointPosition(
          latest.current.pagedEditorRef.current,
          hit,
          dom.context.pagesContainer,
          dom.queries
        ),
      dom.queries,
      () => {
        const editor = latest.current.pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const active = session && hasEditorWorkerProposalRounds(session);
        const proposalGeometry = session ? (active ? proposalTarget ? workerProposalRoundAuthority(session)?.geometry() : null : workerProposalAuthority(session)?.geometry()) : null;
        return editor && session && (!active || proposalTarget || workerOpenReplicaReady(session))
          ? {
              session,
              editor,
              presented: isPresented(dom.context.pagesContainer, dom.queries.displayList),
              ...(proposalGeometry ? { proposalGeometry } : {}),
            }
          : null;
      },
      () =>
        heldCandidate() === created &&
        latest.current.zoom === currentLayout.zoom &&
        dom.context.pagesContainer.isConnected &&
        (isPresented(dom.context.pagesContainer, shownList) || queriesCurrentRef.current),
      (clientX, clientY) => readPluginPositionAtPoint(latest.current.pagedEditorRef, clientX, clientY)
    );
    const resolveAnchor = created.getAnchorGeometry;
    created.getAnchorGeometry = (target) => {
      const previous = proposalTarget;
      proposalTarget = target.kind === 'proposal';
      try { return resolveAnchor(target); }
      finally { proposalTarget = previous; }
    };
    return created;
    // `moved` rebuilds the geometry when its elements move without a new frame.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [host, currentLayout, dom, options.queries, layer, moved]);
  geometryRef.current = geometry;
  // Overlays keep the geometry the host last adopted until geometry for its next layout exists.
  const adopted = geometry && host.layoutId() === geometry.layout.id ? geometry : null;
  const heldGeometry = adopted ?? heldCandidate();
  useLayoutEffect(() => {
    if (adopted && options.session) {
      adoptedRef.current = { session: options.session, geometry: adopted };
    } else if (!managed || !heldCandidate()) {
      adoptedRef.current = null;
    }
    // Pages that show a layout of the current version get its geometry next.
    queriesCurrentRef.current = layout !== null;
  });
  // A frame painted during the hold may retire the held geometry.
  const holding = !adopted && heldGeometry !== null;
  useEffect(() => {
    if (!holding) return;
    return onPresented(() => host.geometryChanged());
  }, [host, holding]);

  useEffect(() => {
    host.layoutChanged(currentLayout);
    publishSelection();
  }, [host, currentLayout, publishSelection]);

  useLayoutEffect(() => {
    if (!managed) return;
    return displayWindowOf(options.queries)?.subscribe(() => {
      host.geometryChanged();
      if (layoutRef.current) host.layoutPresented(layoutRef.current);
    });
  }, [host, managed, options.queries]);

  useEffect(() => {
    host.geometryChanged();
    if (!geometry || !dom) return;
    const shown = () => isPresented(dom.context.pagesContainer, dom.queries.displayList);
    if (shown()) {
      host.geometryPresented(geometry.layout);
      return;
    }
    let frame = 0;
    const settle = () => {
      if (!shown()) {
        frame = requestAnimationFrame(settle);
        return;
      }
      host.geometryChanged();
      host.layoutPresented(geometry.layout);
    };
    frame = requestAnimationFrame(settle);
    return () => cancelAnimationFrame(frame);
  }, [host, geometry, dom]);

  const viewerTargets = useRef<{
    read: ResidentEngineWorkerClient['documentRead'];
    version: string;
    targets: Map<string, ResidentDocumentReadValues['navigationTarget']>;
    pending: Set<string>;
  } | null>(null);
  const place = useCallback(
    (anchor: DocxPluginSidebarItem<unknown>['anchor']) => {
      const { queries, zoom, canvasHostRef, overlayTarget, viewerDocumentRead } = latest.current;
      const pages = canvasHostRef.current;
      if (!queries || !pages || !overlayTarget) return null;
      let resolved: ResidentDocumentReadValues['navigationTarget'];
      if (viewerDocumentRead) {
        const version = presentedWorkerVersion(queries);
        if (version === null) return null;
        if (viewerTargets.current?.version !== version || viewerTargets.current.read !== viewerDocumentRead) {
          viewerTargets.current = { read: viewerDocumentRead, version, targets: new Map(), pending: new Set() };
        }
        if (anchor.version !== host.version() || anchor.version !== version) return null;
        const cache = viewerTargets.current;
        const key = `${anchor.version}\u0000${anchor.story}\u0000${anchor.paraId}`;
        const target = cache.targets.get(key);
        if (target === undefined) {
          if (!cache.pending.has(key)) {
            cache.pending.add(key);
            void viewerDocumentRead({ kind: 'navigationTarget', story: anchor.story, paraId: anchor.paraId })
              .then((reply) => {
                if (viewerTargets.current !== cache || latest.current.viewerDocumentRead !== viewerDocumentRead ||
                  presentedWorkerVersion(latest.current.queries) !== version || host.version() !== version ||
                  reply.version !== version) return;
                if (typeof reply.value === 'object' && reply.value !== null) {
                  cache.targets.set(key, reply.value);
                  setMoved((value) => value + 1);
                }
              }, () => undefined)
              .finally(() => cache.pending.delete(key));
          }
          return null;
        }
        resolved = target;
      } else {
        const session = latest.current.pagedEditorRef.current?.getYrsSession() ?? null;
        if (!session || anchor.version !== host.version() || sourceVersionOf(queries) !== anchor.version) return null;
        resolved = resolveParagraph(session, anchor);
      }
      if (typeof resolved === 'string') return null;
      const rect = queries.anchorRect(resolved.position);
      if (!rect) return null;
      const pageRect = resolveDisplayPageClientRect(pages, queries, rect.pageIndex);
      const pageSize = queries.pageSize(rect.pageIndex);
      if (!pageRect || !pageSize || pageSize.height <= 0) return null;
      const y =
        pageRect.top -
        overlayTarget.getBoundingClientRect().top +
        rect.y * (pageRect.height / pageSize.height);
      const scale = (zoom > 0 ? zoom : 1) * effectiveZoom(overlayTarget);
      return { position: resolved.position, y: y / scale };
    },
    [host]
  );

  const sidebarItems = useMemo(
    () => (activations.length > 0 ? managedSidebarItems(host, activations, place) : []),
    // `moved` re-places cards when the pages move without a new frame.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [host, activations, place, options.queries, options.zoom, version, moved]
  );

  const beginLoad = useCallback(() => {
    viewerSelectionRef.current = null;
    detachAuthority.current?.();
    host.close('document-replaced');
  }, [host]);

  return {
    host,
    managed,
    activations: managed ? activations : NO_ACTIVATIONS,
    sidebarItems,
    renderedDomContext: dom?.context ?? null,
    heldGeometry,
    overlayLayerRef: setLayer,
    onRenderedDomContext,
    beginLoad,
    publishSelection,
    publishViewerSelection,
  };
}
