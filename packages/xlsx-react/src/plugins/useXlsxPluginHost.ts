import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import type { RefObject } from 'react';
import type { TFunction } from '@betteroffice/xlsx-i18n';
import type { Proposal } from '@betteroffice/xlsx';
import type { XlsxCommandController } from '../commands/createXlsxCommandStore';
import type { XlsxPluginEditorAccess } from './createPluginClients';
import type { XlsxPluginDocumentAccess } from './createPluginDocumentAccess';
import type { createPeerlessPluginGeometry } from './createPeerlessPluginDocument';
import type { createPluginClients } from './createPluginClients';
import {
  createXlsxPluginHost,
  type XlsxPluginActivation,
  type XlsxPluginHost,
} from './createXlsxPluginHost';
import {
  createPluginGeometry,
  pluginLayout,
  type EditorPointQuery,
  type PaintedGrid,
} from './geometry';
import { rowDecorationIndex, type RowDecorationSet, type RowGutterIndex } from './rowDecorations';
import type { XlsxEditorPluginProps, XlsxPluginGeometry, XlsxPluginSelection } from './types';

const NO_ACTIVATIONS: readonly XlsxPluginActivation[] = Object.freeze([]);
const NO_DECORATIONS: readonly RowDecorationSet[] = Object.freeze([]);

/** The frame the canvas shows, for chrome that follows every paint. */
export interface PaintedGridStore {
  get(): PaintedGrid | null;
  subscribe(listener: () => void): () => void;
}

export interface UseXlsxPluginHostOptions extends XlsxEditorPluginProps {
  createClients?: typeof createPluginClients;
  createGeometry?: typeof createPeerlessPluginGeometry;
  /** Stable; its members read the editor at call time. */
  access: XlsxPluginEditorAccess | XlsxPluginDocumentAccess;
  /** False when the editor cannot gate typed edits; read once. */
  editGate?: boolean;
  commands: XlsxCommandController;
  readOnly: boolean;
  proposals?: readonly Proposal[];
  staleFor?: Readonly<Record<string, readonly string[]>>;
  proposalsPending?(): boolean;
  /** The workbook once it is open, else null. Plugins open it while the editor still holds it. */
  handle: object | null;
  selection: XlsxPluginSelection;
  /** The canvas the grid paints into. */
  canvasRef: RefObject<HTMLCanvasElement | null>;
  /** The editor's hit test that plugin point queries go through; read at call time. */
  positionAtPoint?: EditorPointQuery;
  t: TFunction;
}

export interface XlsxPluginHostBinding {
  host: XlsxPluginHost;
  /** Whether any plugin is installed. */
  managed: boolean;
  activations: readonly XlsxPluginActivation[];
  overlayLayerRef: (element: HTMLDivElement | null) => void;
  /** Ends every plugin activation as a new workbook starts loading. */
  beginLoad(): void;
  /**
   * Records the frame the canvas now shows. Its layout and geometry replace the previous ones
   * in the same step, so retained geometry of an earlier frame refuses at once.
   */
  presentGrid(painted: PaintedGrid | null): void;
  /** Current row decorations of the active sheet by row, or null when the gutter is hidden. */
  rowGutter: RowGutterIndex | null;
  painted: PaintedGridStore;
}

function proposalSignature(options: UseXlsxPluginHostOptions, version: string): string {
  const proposals = options.proposals ?? [];
  return JSON.stringify({ proposals, staleFor: options.staleFor ?? {}, version: proposals.length ? version : null });
}

/** Owns the plugin host of one `XlsxEditor` and binds it to the editor's authority. */
export function useXlsxPluginHost(options: UseXlsxPluginHostOptions): XlsxPluginHostBinding {
  const latest = useRef(options);
  latest.current = options;
  const geometryRef = useRef<XlsxPluginGeometry | null>(null);
  const paintedRef = useRef<PaintedGrid | null>(null);
  const layerRef = useRef<HTMLDivElement | null>(null);
  const [paintedStore] = useState(() => {
    const listeners = new Set<() => void>();
    return {
      get: () => paintedRef.current,
      subscribe(listener: () => void) {
        listeners.add(listener);
        return () => { listeners.delete(listener); };
      },
      emit() {
        for (const listener of [...listeners]) listener();
      },
    };
  });

  const [host] = useState(() =>
    createXlsxPluginHost({
      ...options.access,
      translate: (key) => latest.current.t(key),
      geometry: () => geometryRef.current,
    }, options.createClients, { editGate: options.editGate })
  );

  const managed = (options.plugins?.length ?? 0) > 0;

  /** Publishes the layout and geometry of the painted frame for the current version. */
  const publish = useCallback(() => {
    const painted = paintedRef.current;
    const canvas = latest.current.canvasRef.current;
    const layer = layerRef.current;
    const installed = (latest.current.plugins?.length ?? 0) > 0;
    const layout =
      installed && host.generation() !== null ? pluginLayout(painted, host.version()) : null;
    const pointAt: EditorPointQuery | undefined = latest.current.positionAtPoint
      ? (clientX, clientY) => latest.current.positionAtPoint?.(clientX, clientY) ?? null
      : undefined;
    const proposals = () => {
      const options = latest.current;
      if (options.proposalsPending?.() || !options.selection || options.selection.sheetId !== layout?.sheetId) return null;
      return { proposals: options.proposals ?? [], sheetIndex: options.selection.sheetIndex,
        previewVersion: host.previewVersion() };
    };
    geometryRef.current =
      layout && painted && canvas && layer
        ? 'identity' in latest.current.access ? latest.current.createGeometry?.(
            layout,
            painted,
            canvas,
            layer,
            () => {
              const access = latest.current.access;
              const live = host.layoutId() === layout.id && canvas.isConnected && layer.isConnected;
              if (!('identity' in access)) return { version: null, painted: null };
              const identity = access.identity();
              return { version: identity ? access.version(identity) : null,
                painted: live ? paintedRef.current : null };
            },
            proposals,
            pointAt
          ) ?? null : createPluginGeometry(layout, painted, canvas, layer, () =>
              host.layoutId() === layout.id &&
              paintedRef.current === painted &&
              canvas.isConnected &&
              layer.isConnected, proposals, pointAt)
        : null;
    host.layoutChanged(layout);
    host.geometryChanged();
  }, [host]);

  useLayoutEffect(() => {
    host.setReporter(options.onPluginError);
  });

  useLayoutEffect(() => {
    if (!('identity' in options.access)) host.modeChanged(options.readOnly);
  }, [host, options.readOnly]);

  useLayoutEffect(() => {
    host.setGrants(options.pluginGrants);
  });

  useLayoutEffect(() => {
    host.setPlugins(options.plugins);
  }, [host, options.plugins]);

  useEffect(() => () => host.close('unmounted'), [host]);

  useEffect(() => {
    if (!options.handle) return;
    const access = options.access;
    const identity = 'identity' in access ? access.identity() : access.handle();
    if (identity !== options.handle) return;
    host.open(options.handle);
    host.proposalsChanged(proposalSignature(latest.current, host.version()));
    publish();
    return () => host.close('document-replaced');
  }, [host, options.access, options.handle, publish]);

  useEffect(() => {
    host.syncCommands();
    return host.subscribe(() => host.syncCommands());
  }, [host]);

  const proposalState = proposalSignature(options, host.version());
  useLayoutEffect(() => {
    if (options.handle && host.generation()) host.proposalsChanged(proposalState);
  }, [host, options.handle, proposalState]);

  const activations = useSyncExternalStore(host.subscribe, host.activations, () => NO_ACTIVATIONS);
  const decorations = managed && activations.length ? host.rowDecorations() : NO_DECORATIONS;
  const activeSheetId = options.selection?.sheetId ?? null;
  const t = options.t;
  const rowGutter = useMemo(() => rowDecorationIndex(decorations, activeSheetId,
    (row) => t('a11y.rowHeaderLabel', { row: row + 1 })), [decorations, activeSheetId, t]);

  useEffect(() => {
    if (managed) host.selectionChanged(options.selection);
  }, [host, managed, options.handle, options.selection]);

  useEffect(() => {
    publish();
  }, [managed, publish]);

  const presentGrid = useCallback(
    (next: PaintedGrid | null) => {
      paintedRef.current = next;
      if (!next || (latest.current.plugins?.length ?? 0) > 0) publish();
      paintedStore.emit();
    },
    [paintedStore, publish]
  );

  const overlayLayerRef = useCallback(
    (element: HTMLDivElement | null) => {
      layerRef.current = element;
      publish();
    },
    [publish]
  );

  const beginLoad = useCallback(() => {
    presentGrid(null);
    host.close('document-replaced');
  }, [host, presentGrid]);

  return {
    host,
    managed,
    activations: managed ? activations : NO_ACTIVATIONS,
    overlayLayerRef,
    beginLoad,
    presentGrid,
    rowGutter,
    painted: paintedStore,
  };
}
