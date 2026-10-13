import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { RefObject } from 'react';
import type { DeckSnapshot, ShapeSnapshot } from '@betteroffice/pptx';
import { indexGeometry, transformedRect } from '@betteroffice/pptx/peerless';
import { findShape } from '../interactions';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import type { PptxCommandStore } from '../commands/types';
import { createPptxPluginHost } from '../plugins/createPptxPluginHost';
import type { PptxEditorPluginProps, PptxPluginGeometry, PptxPluginSelection } from '../plugins/types';
import { EMU_PER_PIXEL, slideRectToOverlay } from '../plugins/geometry';
import { useTranslation } from '../i18n';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { peerlessPositionAtPoint, type PptxAsyncEditorApi } from './createPeerlessEditorApi';
import { peerlessPluginClients } from './peerlessPluginClients';
import { peerlessPluginDocument } from './peerlessPluginDocument';

export function usePeerlessPluginHost(run: PeerlessPresentation | null, commands: PptxCommandStore,
  canvas: RefObject<HTMLCanvasElement | null>, api: RefObject<PptxAsyncEditorApi | null>,
  props: PptxEditorPluginProps & { readOnly?: boolean }, reviewing: boolean) {
  const { t } = useTranslation();
  const latest = useRef({ commands, t, reviewing });
  latest.current = { commands, t, reviewing };
  const geometry = useRef<PptxPluginGeometry | null>(null);
  const [layer, layerRef] = useState<HTMLDivElement | null>(null);
  const managed = Boolean(props.plugins?.length);
  const binding = useRef<{ run: PeerlessPresentation; document: ReturnType<typeof peerlessPluginDocument> } | null>(null);
  const host = useMemo(() => createPptxPluginHost({
    commands: () => pptxCommandController(latest.current.commands),
    translate: (key) => latest.current.t(key), geometry: () => geometry.current,
  }, (invocation, grant, store) => {
    const current = binding.current!;
    return peerlessPluginClients(invocation, current.run, () => api.current, grant, store,
      (intent) => current.document.dispatch(invocation, intent));
  }), [api]);

  useLayoutEffect(() => { host.setReporter(props.onPluginError); host.setGrants(props.pluginGrants); });
  useLayoutEffect(() => { host.modeChanged(Boolean(props.readOnly || run?.failure)); }, [host, props.readOnly, run?.failure]);
  useLayoutEffect(() => { host.setPlugins(props.plugins); }, [host, props.plugins]);
  useLayoutEffect(() => () => host.close('unmounted'), [host]);
  useLayoutEffect(() => {
    if (!run?.current || run.failure) return;
    const current = { run, document: peerlessPluginDocument(run) };
    binding.current = current;
    host.open(current.document.document);
    return () => {
      host.close('document-replaced');
      if (binding.current === current) binding.current = null;
    };
  }, [host, run, run?.failure]);
  useLayoutEffect(() => {
    host.syncCommands();
    return host.subscribe(() => host.syncCommands());
  }, [host]);

  const subscribe = useCallback((listener: () => void) =>
    managed && run ? run.subscribeVisual(listener) : () => {}, [run, managed]);
  const snapshot = useCallback(() => managed ? run?.getVisualSnapshot() ?? null : null, [run, managed]);
  const visual = useSyncExternalStore(subscribe, snapshot, snapshot);
  const confirmed = visual?.view.confirmed;
  const fit = visual?.fit;
  const version = confirmed?.frame.version;
  const [shapes, setShapes] = useState<{ run: PeerlessPresentation; version: string; deck: DeckSnapshot } | null>(null);
  useEffect(() => {
    let current = true;
    if (run?.current && managed && version) {
      void run.client.access!('snapshot').then((deck) => {
        if (current && run.current && run.client.state.version === version) setShapes({ run, version, deck });
      }).catch(() => {});
    }
    return () => { current = false; };
  }, [run, version, managed]);

  const presentedGeometry = useMemo<PptxPluginGeometry | null>(() => {
    const surface = canvas.current;
    if (!managed || reviewing || !run?.current || !confirmed || !fit || !surface || !layer ||
      shapes?.run !== run || shapes.version !== confirmed.frame.version || confirmed.frame.version !== run.client.state.version) return null;
    const { frame, displayList } = confirmed;
    const layout = {
      id: `${frame.incarnation}:${frame.epoch}:${fit.generation}`, version: frame.version,
      slideId: frame.slideId, slide: frame.slideIndex + 1, zoom: fit.scale,
      width: displayList.width, height: displayList.height,
    };
    const current = () => run.current && !latest.current.reviewing && surface.isConnected &&
      host.layoutId() === layout.id && run.getVisualSnapshot()?.view.confirmed === confirmed &&
      run.getVisualSnapshot()?.fit === fit && run.store.getSnapshot().confirmed === confirmed && run.client.state.version === frame.version;
    const toOverlay = (rect: { x: number; y: number; width: number; height: number }) =>
      current() ? slideRectToOverlay(surface, layer, displayList, rect) : null;
    const indexed = indexGeometry(displayList);
    return {
      layout,
      toOverlayRect: (input) => {
        const rect = input?.rect;
        if (!rect || ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)) return null;
        if (input.space === 'slide-px') return toOverlay(rect);
        if (input.space !== 'slide-emu') return null;
        return toOverlay({ x: rect.x / EMU_PER_PIXEL, y: rect.y / EMU_PER_PIXEL,
          width: rect.width / EMU_PER_PIXEL, height: rect.height / EMU_PER_PIXEL });
      },
      getShapeRect: (id) => {
        const slide = shapes.deck.slides.find((slide) => slide.id === frame.slideId);
        const shape = slide && findShape(slide.shapes, id);
        if (!shape) return null;
        const ids = new Set<string>();
        const collect = (shape: ShapeSnapshot) => { ids.add(shape.id); shape.children.forEach(collect); };
        collect(shape);
        const bounds = indexed.hits.filter((hit) => hit.primitive.shapeId && ids.has(hit.primitive.shapeId))
          .map((hit) => transformedRect(hit.transforms, { x: hit.primitive.x, y: hit.primitive.y, width: hit.primitive.w, height: hit.primitive.h }));
        if (!bounds.length) return null;
        const x = Math.min(...bounds.map((rect) => rect.x));
        const y = Math.min(...bounds.map((rect) => rect.y));
        return toOverlay({ x, y, width: Math.max(...bounds.map((rect) => rect.x + rect.width)) - x,
          height: Math.max(...bounds.map((rect) => rect.y + rect.height)) - y });
      },
      getPositionAtPoint: (x, y) => {
        if (!current()) return null;
        const hit = peerlessPositionAtPoint(run, surface, x, y);
        return hit ? { ...hit, version: frame.version, layoutId: layout.id } : null;
      },
    };
  }, [host, managed, reviewing, run, confirmed, fit, canvas, layer, shapes]);
  geometry.current = presentedGeometry;
  useLayoutEffect(() => {
    host.layoutChanged(presentedGeometry?.layout ?? null);
    host.geometryChanged();
  }, [host, presentedGeometry]);
  useLayoutEffect(() => {
    if (!confirmed) return;
    const { frame } = confirmed;
    const selection = visual?.view.selection;
    const selected: PptxPluginSelection = { slideId: frame.slideId, slide: frame.slideIndex + 1,
      target: !selection ? { kind: 'slide' } : selection.kind === 'object' ? { kind: 'shape', shapeId: selection.shapeId }
        : { kind: 'text', shapeId: selection.shapeId, storyId: selection.storyId, anchor: selection.anchor, focus: selection.focus } };
    host.selectionChanged(selected);
  }, [host, confirmed, visual?.view.selection]);
  useEffect(() => {
    const surface = canvas.current;
    if (!managed || !surface || !layer) return;
    let frame = 0;
    const moved = () => {
      if (frame) return;
      frame = requestAnimationFrame(() => { frame = 0; host.geometryChanged(); });
    };
    const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(moved);
    observer?.observe(surface); observer?.observe(layer);
    window.addEventListener('resize', moved);
    return () => { if (frame) cancelAnimationFrame(frame); observer?.disconnect(); window.removeEventListener('resize', moved); };
  }, [host, managed, canvas, layer]);
  const activations = useSyncExternalStore(host.subscribe, host.activations, host.activations);
  return { host, activations, layerRef, managed };
}
