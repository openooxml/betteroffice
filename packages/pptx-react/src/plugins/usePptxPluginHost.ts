import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import type { PresentationHandle, SlideDisplayList } from '@betteroffice/pptx';
import type { TFunction } from '@betteroffice/pptx-i18n';
import type { PptxCommandController } from '../commands/createPptxCommandStore';
import type { PptxPointPosition } from '../PptxEditor';
import type { PptxPluginEditorAccess } from './createPluginClients';
import {
  createPptxPluginHost,
  type PptxPluginActivation,
  type PptxPluginHost,
} from './createPptxPluginHost';
import { createPluginGeometry, pluginLayout, type PresentedSlide } from './geometry';
import type { PptxEditorPluginProps, PptxPluginGeometry, PptxPluginSelection } from './types';

const NO_ACTIVATIONS: readonly PptxPluginActivation[] = Object.freeze([]);

/** A presented slide and the canvas that shows it. */
export interface PresentedCanvas extends PresentedSlide {
  canvas: HTMLCanvasElement;
}

export interface UsePptxPluginHostOptions extends PptxEditorPluginProps {
  /** Stable; its members read the editor at call time. */
  access: PptxPluginEditorAccess;
  commands: PptxCommandController;
  readOnly: boolean;
  /** The presentation once it is open and laid out, else null. Plugins open it while it is current. */
  handle: PresentationHandle | null;
  /** Proposal review paints a preview over the slide, so no layout is published meanwhile. */
  reviewing: boolean;
  selection: PptxPluginSelection;
  /** Hit-tests client coordinates while `frame` is the editor's current frame. */
  pointAt(frame: SlideDisplayList, clientX: number, clientY: number): PptxPointPosition | null;
  t: TFunction;
}

export interface PptxPluginHostBinding {
  host: PptxPluginHost;
  /** Whether any plugin is installed. */
  managed: boolean;
  activations: readonly PptxPluginActivation[];
  overlayLayerRef: (element: HTMLDivElement | null) => void;
  /** Ends every plugin activation as a new presentation starts loading. */
  beginLoad(): void;
  /** Records what the canvas shows: null once a paint starts, the slide once it is painted. */
  presentFrame(presented: PresentedCanvas | null): void;
}

/** Owns the plugin host of one `PptxEditor` and binds it to the editor's authority. */
export function usePptxPluginHost(options: UsePptxPluginHostOptions): PptxPluginHostBinding {
  const latest = useRef(options);
  latest.current = options;
  const geometryRef = useRef<PptxPluginGeometry | null>(null);

  const [host] = useState(() =>
    createPptxPluginHost({
      ...options.access,
      translate: (key) => latest.current.t(key),
      geometry: () => geometryRef.current,
    })
  );

  const managed = (options.plugins?.length ?? 0) > 0;

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
    return () => host.close('document-replaced');
  }, [host, options.access, options.handle]);

  useEffect(() => {
    host.syncCommands();
    return host.subscribe(() => host.syncCommands());
  }, [host]);

  const activations = useSyncExternalStore(host.subscribe, host.activations, () => NO_ACTIVATIONS);

  useEffect(() => {
    if (managed) host.selectionChanged(options.selection);
  }, [host, managed, options.selection]);

  const presentedRef = useRef<PresentedCanvas | null>(null);
  const [presented, setPresented] = useState<PresentedCanvas | null>(null);
  const presentFrame = useCallback(
    (next: PresentedCanvas | null) => {
      presentedRef.current = next;
      if (!next) host.layoutChanged(null);
      if (latest.current.plugins?.length) setPresented(next);
    },
    [host]
  );
  useEffect(() => {
    if (managed) setPresented(presentedRef.current);
  }, [managed]);

  const [layer, setLayer] = useState<HTMLDivElement | null>(null);
  const [moved, setMoved] = useState(0);
  const canvas = presented?.canvas ?? null;
  useEffect(() => {
    if (!managed || !canvas || !layer) return;
    let frame = 0;
    const bump = () => {
      if (frame) return;
      frame = requestAnimationFrame(() => {
        frame = 0;
        setMoved((value) => value + 1);
      });
    };
    const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(bump);
    observer?.observe(canvas);
    observer?.observe(layer);
    window.addEventListener('resize', bump);
    return () => {
      if (frame) cancelAnimationFrame(frame);
      observer?.disconnect();
      window.removeEventListener('resize', bump);
    };
  }, [managed, canvas, layer]);

  const version = useSyncExternalStore(host.subscribe, host.version, host.version);
  const layout = useMemo(
    () => (managed && !options.reviewing ? pluginLayout(presented, version) : null),
    [managed, options.reviewing, presented, version]
  );

  const geometry = useMemo(
    () =>
      layout && presented && layer
        ? createPluginGeometry(
            layout,
            presented,
            presented.canvas,
            layer,
            () =>
              host.layoutId() === layout.id &&
              presentedRef.current === presented &&
              presented.canvas.isConnected,
            (frame, clientX, clientY) => latest.current.pointAt(frame, clientX, clientY)
          )
        : null,
    // `moved` rebuilds the geometry when its elements move without a new frame.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [host, layout, presented, layer, moved]
  );
  geometryRef.current = geometry;

  useEffect(() => {
    host.layoutChanged(layout);
  }, [host, layout]);

  useEffect(() => {
    host.geometryChanged();
  }, [host, geometry]);

  const beginLoad = useCallback(() => {
    presentFrame(null);
    host.close('document-replaced');
  }, [host, presentFrame]);

  return {
    host,
    managed,
    activations: managed ? activations : NO_ACTIVATIONS,
    overlayLayerRef: setLayer,
    beginLoad,
    presentFrame,
  };
}
