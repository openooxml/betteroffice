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
  resolveDisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import type { RenderedDomContext } from '@betteroffice/docx/plugin-api';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { DocxCommandController } from '../commands/createDocxCommandStore';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';
import { sourceVersionOf } from '../components/DocxEditor/internals/layoutProvenance';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { SelectionState } from '../components/DocxEditor/types';
import type { ReactSidebarItem } from '../plugin-api/types';
import {
  createDocxPluginHost,
  type DocxPluginActivation,
  type DocxPluginHost,
} from './createDocxPluginHost';
import { resolveParagraph } from './createPluginClients';
import { createPluginGeometry, pluginLayout } from './geometry';
import { managedSidebarItems } from './PluginSidebarItems';
import type {
  DocxEditorPluginProps,
  DocxPluginGeometry,
  DocxPluginLayout,
  DocxPluginSidebarItem,
} from './types';

const NO_ACTIVATIONS: readonly DocxPluginActivation[] = Object.freeze([]);
const LAYOUT_WAIT_MS = 1000;

type RenderedDom = { context: RenderedDomContext; queries: DisplayListQueries };

export interface UseDocxPluginHostOptions extends DocxEditorPluginProps {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  writeModeRef: React.RefObject<EditorMode>;
  mode: EditorMode;
  /** Host `readOnly` or viewing mode. */
  readOnly: boolean;
  commands: DocxCommandController;
  /** The authoritative session once a document is ready, else null. */
  session: YrsSession | null;
  /** Changes whenever a new document load starts. */
  loadGeneration: number;
  queries: DisplayListQueries | null;
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
  overlayLayerRef: (element: HTMLDivElement | null) => void;
  /** Receives each rendered-DOM context the paged editor builds. */
  onRenderedDomContext(context: RenderedDomContext, queries: DisplayListQueries): void;
  /** Ends every plugin activation as a new document starts loading. */
  beginLoad(): void;
  /** Publishes the current selection to plugins. */
  publishSelection(): void;
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
  const layoutRef = useRef<DocxPluginLayout | null>(null);
  const formattingRef = useRef<SelectionState | null>(null);

  const [host] = useState(() =>
    createDocxPluginHost({
      pagedEditorRef: options.pagedEditorRef,
      writeMode: () => latest.current.writeModeRef.current ?? 'viewing',
      commands: () => latest.current.commands,
      translate: (key) => translateRef.current(key),
      geometry: () => geometryRef.current,
      async settledLayout(version) {
        const deadline = Date.now() + LAYOUT_WAIT_MS;
        while (sourceVersionOf(latest.current.queries) !== version) {
          if (Date.now() >= deadline) return false;
          await new Promise((resolve) => setTimeout(resolve, 16));
        }
        return true;
      },
    })
  );

  const managed = (options.plugins?.length ?? 0) > 0;

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
    return () => host.close('document-replaced');
  }, [host, options.session, options.loadGeneration]);

  useEffect(() => {
    host.syncCommands();
    return host.subscribe(() => host.syncCommands());
  }, [host]);

  const activations = useSyncExternalStore(host.subscribe, host.activations, () => NO_ACTIVATIONS);

  const publishSelection = useCallback(() => {
    if (!latest.current.plugins?.length) return;
    const editor = latest.current.pagedEditorRef.current;
    const range = editor?.getSelectionRange() ?? null;
    const layout = layoutRef.current;
    const story = editor?.getYrsSession()?.selection()?.head.story ?? 'body';
    host.selectionChanged({
      formatting: formattingRef.current,
      displayRange:
        range && layout ? { story, from: range.from, to: range.to, layoutId: layout.id } : null,
    });
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
  const layout = useMemo(
    () => pluginLayout(options.queries, managed ? version : null, options.zoom),
    [managed, options.queries, options.zoom, version]
  );
  const layoutStable = useRef<DocxPluginLayout | null>(null);
  if (
    layout?.id !== layoutStable.current?.id ||
    layout?.version !== layoutStable.current?.version ||
    layout?.zoom !== layoutStable.current?.zoom
  ) {
    layoutStable.current = layout;
  }
  const currentLayout = layoutStable.current;
  layoutRef.current = currentLayout;

  const geometry = useMemo(
    () =>
      currentLayout && dom && dom.queries === options.queries && layer
        ? createPluginGeometry(
            currentLayout,
            dom.context,
            layer,
            () =>
              host.layoutId() === currentLayout.id &&
              domRef.current === dom &&
              dom.context.pagesContainer.isConnected
          )
        : null,
    // `moved` rebuilds the geometry when its elements move without a new frame.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [host, currentLayout, dom, options.queries, layer, moved]
  );
  geometryRef.current = geometry;

  useEffect(() => {
    host.layoutChanged(currentLayout);
    publishSelection();
  }, [host, currentLayout, publishSelection]);

  useEffect(() => {
    host.geometryChanged();
  }, [host, geometry]);

  const place = useCallback(
    (anchor: DocxPluginSidebarItem<unknown>['anchor']) => {
      const { queries, zoom, canvasHostRef, overlayTarget } = latest.current;
      const session = latest.current.pagedEditorRef.current?.getYrsSession() ?? null;
      const pages = canvasHostRef.current;
      if (!session || !queries || !pages || !overlayTarget) return null;
      if (anchor.version !== host.version() || sourceVersionOf(queries) !== anchor.version) {
        return null;
      }
      const resolved = resolveParagraph(session, anchor);
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
      return { position: resolved.position, y: y / (zoom > 0 ? zoom : 1) };
    },
    [host]
  );

  const sidebarItems = useMemo(
    () => (activations.length > 0 ? managedSidebarItems(host, activations, place) : []),
    // `moved` re-places cards when the pages move without a new frame.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [host, activations, place, options.queries, options.zoom, version, moved]
  );

  const beginLoad = useCallback(() => host.close('document-replaced'), [host]);

  return {
    host,
    managed,
    activations: managed ? activations : NO_ACTIVATIONS,
    sidebarItems,
    renderedDomContext: dom?.context ?? null,
    overlayLayerRef: setLayer,
    onRenderedDomContext,
    beginLoad,
    publishSelection,
  };
}
