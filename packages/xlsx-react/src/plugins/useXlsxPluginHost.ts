import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import type { RefObject } from 'react';
import type { WorkbookHandle } from '@betteroffice/xlsx';
import type { TFunction } from '@betteroffice/xlsx-i18n';
import type { XlsxCommandController } from '../commands/createXlsxCommandStore';
import type { XlsxPluginEditorAccess } from './createPluginClients';
import {
  createXlsxPluginHost,
  type XlsxPluginActivation,
  type XlsxPluginHost,
} from './createXlsxPluginHost';
import { createPluginGeometry, pluginLayout, type PaintedGrid } from './geometry';
import type { XlsxEditorPluginProps, XlsxPluginGeometry, XlsxPluginSelection } from './types';

const NO_ACTIVATIONS: readonly XlsxPluginActivation[] = Object.freeze([]);

export interface UseXlsxPluginHostOptions extends XlsxEditorPluginProps {
  /** Stable; its members read the editor at call time. */
  access: XlsxPluginEditorAccess;
  commands: XlsxCommandController;
  readOnly: boolean;
  /** The workbook once it is open, else null. Plugins open it while the editor still holds it. */
  handle: WorkbookHandle | null;
  selection: XlsxPluginSelection;
  /** The canvas the grid paints into. */
  canvasRef: RefObject<HTMLCanvasElement | null>;
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
}

/** Owns the plugin host of one `XlsxEditor` and binds it to the editor's authority. */
export function useXlsxPluginHost(options: UseXlsxPluginHostOptions): XlsxPluginHostBinding {
  const latest = useRef(options);
  latest.current = options;
  const geometryRef = useRef<XlsxPluginGeometry | null>(null);
  const paintedRef = useRef<PaintedGrid | null>(null);
  const layerRef = useRef<HTMLDivElement | null>(null);

  const [host] = useState(() =>
    createXlsxPluginHost({
      ...options.access,
      translate: (key) => latest.current.t(key),
      geometry: () => geometryRef.current,
    })
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
    geometryRef.current =
      layout && painted && canvas && layer
        ? createPluginGeometry(
            layout,
            painted,
            canvas,
            layer,
            () =>
              host.layoutId() === layout.id &&
              paintedRef.current === painted &&
              canvas.isConnected &&
              layer.isConnected
          )
        : null;
    host.layoutChanged(layout);
    host.geometryChanged();
  }, [host]);

  useLayoutEffect(() => {
    host.setReporter(options.onPluginError);
  });

  useLayoutEffect(() => {
    host.modeChanged(options.readOnly);
  }, [host, options.readOnly]);

  useLayoutEffect(() => {
    host.setGrants(options.pluginGrants);
  });

  useLayoutEffect(() => {
    host.setPlugins(options.plugins);
  }, [host, options.plugins]);

  useEffect(() => () => host.close('unmounted'), [host]);

  useEffect(() => {
    if (!options.handle || options.access.handle() !== options.handle) return;
    host.open(options.handle);
    publish();
    return () => host.close('document-replaced');
  }, [host, options.access, options.handle, publish]);

  useEffect(() => {
    host.syncCommands();
    return host.subscribe(() => host.syncCommands());
  }, [host]);

  const activations = useSyncExternalStore(host.subscribe, host.activations, () => NO_ACTIVATIONS);

  useEffect(() => {
    if (managed) host.selectionChanged(options.selection);
  }, [host, managed, options.selection]);

  useEffect(() => {
    publish();
  }, [managed, publish]);

  const presentGrid = useCallback(
    (painted: PaintedGrid | null) => {
      paintedRef.current = painted;
      if (!painted || (latest.current.plugins?.length ?? 0) > 0) publish();
    },
    [publish]
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
  };
}
