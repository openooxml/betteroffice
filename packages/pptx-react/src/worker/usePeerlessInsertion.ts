import { useCallback, useEffect, useRef, useState, type PointerEvent, type RefObject } from 'react';
import type { ShapeRect } from '@betteroffice/pptx';
import type { PresentationCommandAck } from '@betteroffice/pptx/worker';
import { DEFAULT_EDITOR_TEXT_STYLE } from '@betteroffice/pptx/worker';
import { shapePresetFromTool, type PptxEditorTool } from '../components/toolbarTypes';
import { useTranslation } from '../i18n';
import { slidePoint, type SlidePoint } from '../interactions';
import { setPeerlessGestureActive } from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export interface PeerlessInsertionOrigin { run: PeerlessPresentation; slideId: string }

interface TextBoxEcho { origin: PeerlessInsertionOrigin; rect: ShapeRect; seq?: number }

function displayCoordinate(emu: number): number {
  const pixels = Math.fround(Math.fround(emu) * Math.fround(1 / 9525));
  for (let digits = 1; digits <= 9; digits++) {
    const value = Number(pixels.toPrecision(digits));
    if (Math.fround(value) === pixels) return value;
  }
  return pixels;
}

const imageTypes: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
  gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', webp: 'image/webp', svg: 'image/svg+xml' };

export const peerlessImagePreparation = {
  async prepare(file: File, run: PeerlessPresentation): Promise<{ bytes: ArrayBuffer; contentType: string; width: number; height: number }> {
    const contentType = file.type === 'image/jpg' ? 'image/jpeg' : file.type || imageTypes[file.name.split('.').pop()?.toLowerCase() ?? ''];
    if (file.size === 0 || file.size > 8 * 1024 * 1024 || !Object.values(imageTypes).includes(contentType)) {
      throw new Error('Unsupported picture type or size');
    }
    const bytes = await file.arrayBuffer();
    const preview = contentType === 'image/tiff' ? await run.client.preparePicture?.(bytes) : null;
    if (contentType === 'image/tiff' && !preview) throw new Error('Picture conversion unavailable');
    const blob = preview ? new Blob([preview], { type: 'image/png' }) : file;
    const url = URL.createObjectURL(blob);
    try {
      const image = new Image();
      const size = await new Promise<{ width: number; height: number }>((resolve, reject) => {
        image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
        image.onerror = () => reject(new Error('Could not decode picture'));
        image.src = url;
      });
      if (!(size.width > 0 && size.height > 0)) throw new Error('Invalid picture dimensions');
      return { bytes, contentType, ...size };
    } finally { URL.revokeObjectURL(url); }
  },
};

export function peerlessInsertionOrigin(run: PeerlessPresentation | null): PeerlessInsertionOrigin | null {
  const slideId = run?.store.getSnapshot().confirmed?.frame.slideId;
  return run?.inputEnabled && slideId ? { run, slideId } : null;
}

function requireOrigin(origin: PeerlessInsertionOrigin): void {
  origin.run.require();
  if (!origin.run.inputEnabled || origin.run.store.getSnapshot().confirmed?.frame.slideId !== origin.slideId ||
    !origin.run.client.state.slides.some((slide) => slide.id === origin.slideId)) throw new Error('Picture target changed');
}

export async function insertPeerlessPicture(origin: PeerlessInsertionOrigin, file: File): Promise<PresentationCommandAck> {
  requireOrigin(origin);
  const prepared = await peerlessImagePreparation.prepare(file, origin.run);
  requireOrigin(origin);
  const run = origin.run;
  const size = run.client.state.size;
  const frame = run.store.getSnapshot().confirmed!.displayList;
  const scale = Math.min(frame.width * 0.5 / prepared.width, frame.height * 0.5 / prepared.height, 1);
  const width = prepared.width * scale * size.width / frame.width;
  const height = prepared.height * scale * size.height / frame.height;
  const rect = { x: Math.round((size.width - width) / 2), y: Math.round((size.height - height) / 2),
    width: Math.max(1, Math.round(width)), height: Math.max(1, Math.round(height)) };
  const ack = await run.store.dispatch({ kind: 'insertPicture', payload: { ...prepared,
    slideId: origin.slideId, name: file.name || 'Picture', rect, inputTime: performance.now(), undoGroup: 'picture' } });
  run.require();
  return ack;
}

export function usePeerlessInsertion(run: PeerlessPresentation | null, canvas: RefObject<HTMLCanvasElement | null>, focus: () => void) {
  const { t } = useTranslation();
  const [tool, updateTool] = useState<PptxEditorTool>('select');
  const [preview, setPreview] = useState<ShapeRect | null>(null);
  const [echo, setEcho] = useState<TextBoxEcho | null>(null);
  const pending = useRef<TextBoxEcho | null>(null);
  const gesture = useRef<{ origin: PeerlessInsertionOrigin; start: SlidePoint; end: SlidePoint;
    pointerId: number; element: HTMLElement; tool: PptxEditorTool } | null>(null);
  const cancel = useCallback(() => {
    const current = gesture.current;
    if (current) setPeerlessGestureActive(current.origin.run.store, 'insertion', false);
    gesture.current = null;
    current?.element.releasePointerCapture?.(current.pointerId);
    setPreview(null);
    pending.current = null;
    setEcho(null);
  }, []);
  const setTool = useCallback((value: PptxEditorTool) => { pending.current = null; setEcho(null); updateTool(value); }, []);
  useEffect(() => {
    setTool('select');
    return cancel;
  }, [run, cancel]);
  useEffect(() => {
    window.addEventListener('blur', cancel);
    const changed = () => {
      if (gesture.current && (!run?.inputEnabled || run.store.getSnapshot().confirmed?.frame.slideId !== gesture.current.origin.slideId)) cancel();
      const insertion = pending.current;
      if (!insertion) return;
      const displayed = run?.getVisualSnapshot()?.view.confirmed?.frame;
      if (!run?.inputEnabled || run.store.getSnapshot().confirmed?.frame.slideId !== insertion.origin.slideId ||
        displayed?.slideId !== insertion.origin.slideId ||
        insertion.seq !== undefined && displayed.coversThroughSeq >= insertion.seq) {
        pending.current = null;
        setEcho(null);
      }
    };
    const release = run?.store.subscribePresentation(changed);
    const visual = run?.subscribeVisual(changed);
    return () => { window.removeEventListener('blur', cancel); release?.(); visual?.(); };
  }, [run, cancel]);
  const point = (event: PointerEvent<HTMLElement>) => {
    const frame = run?.getVisualSnapshot()?.view.confirmed;
    return frame && canvas.current ? slidePoint(canvas.current.getBoundingClientRect(), frame.displayList, event.clientX, event.clientY) : null;
  };
  const onPointerDown = (event: PointerEvent<HTMLElement>): boolean => {
    if (tool === 'select') return false;
    const origin = peerlessInsertionOrigin(run);
    const start = point(event);
    if (!origin || !start || event.button !== 0 || event.isPrimary === false) return true;
    cancel();
    gesture.current = { origin, start, end: start, pointerId: event.pointerId, element: event.currentTarget, tool };
    setPeerlessGestureActive(origin.run.store, 'insertion', true);
    event.currentTarget.setPointerCapture?.(event.pointerId);
    event.preventDefault();
    return true;
  };
  const onPointerMove = (event: PointerEvent<HTMLElement>): boolean => {
    const current = gesture.current;
    if (!current || current.pointerId !== event.pointerId) return false;
    const end = point(event);
    if (end) {
      current.end = end;
      setPreview({ x: Math.min(current.start.x, end.x), y: Math.min(current.start.y, end.y),
        width: Math.abs(current.start.x - end.x), height: Math.abs(current.start.y - end.y) });
    }
    return true;
  };
  const onPointerUp = (event: PointerEvent<HTMLElement>): boolean => {
    const current = gesture.current;
    if (!current || current.pointerId !== event.pointerId) return false;
    const end = point(event) ?? current.end;
    cancel();
    try { requireOrigin(current.origin); } catch { return true; }
    const frame = current.origin.run.store.getSnapshot().confirmed!.displayList;
    const size = current.origin.run.client.state.size;
    const dragged = Math.abs(end.x - current.start.x) >= 6 || Math.abs(end.y - current.start.y) >= 6;
    const textbox = current.tool === 'textBox';
    const width = dragged ? Math.abs(end.x - current.start.x) : frame.width * (textbox ? 0.36 : 0.22);
    const height = dragged ? Math.abs(end.y - current.start.y) : frame.height * (textbox ? 0.12 : 0.18);
    const x = Math.max(0, Math.min(dragged ? Math.min(current.start.x, end.x) : current.start.x - (textbox ? 0 : width / 2), frame.width - width));
    const y = Math.max(0, Math.min(dragged ? Math.min(current.start.y, end.y) : current.start.y - (textbox ? 0 : height / 2), frame.height - height));
    const rect = { x: Math.round(x * size.width / frame.width), y: Math.round(y * size.height / frame.height),
      width: Math.round(Math.max(12, width) * size.width / frame.width), height: Math.round(Math.max(12, height) * size.height / frame.height) };
    const payload = { slideId: current.origin.slideId, inputTime: event.timeStamp, undoGroup: 'insertion' };
    const preset = shapePresetFromTool(current.tool);
    const insertion: TextBoxEcho | null = textbox ? { origin: current.origin, rect: {
      x: displayCoordinate(rect.x), y: displayCoordinate(rect.y),
      width: displayCoordinate(rect.width), height: displayCoordinate(rect.height),
    } } : null;
    pending.current = insertion;
    setEcho(insertion);
    const operation = textbox ? current.origin.run.store.dispatch({ kind: 'insertTextBox', payload: { ...payload,
      draft: { name: t('objects.defaultTextBoxName'), rect, text: '', style: DEFAULT_EDITOR_TEXT_STYLE } } })
      : preset ? current.origin.run.store.dispatch({ kind: 'insertShape', payload: { ...payload,
        draft: { name: preset, rect, geometry: preset, fill: '#d9eaf7' } } }) : null;
    updateTool('select');
    if (operation && textbox) focus();
    const clear = () => { if (pending.current === insertion) { pending.current = null; setEcho(null); } };
    void operation?.then((ack) => {
      if (!insertion || pending.current !== insertion) return;
      if (ack.outcomes.some((outcome) => outcome.seq === ack.throughSeq && outcome.status === 'refused')) { clear(); return; }
      insertion.seq = ack.throughSeq;
      const displayed = current.origin.run.getVisualSnapshot()?.view.confirmed?.frame;
      if (!current.origin.run.inputEnabled || displayed?.slideId !== insertion.origin.slideId ||
        displayed.coversThroughSeq >= insertion.seq) clear();
    }, clear);
    return true;
  };
  return { tool, preview, echo: echo?.origin.run === run ? echo : null, setTool, cancel, onPointerDown, onPointerMove, onPointerUp,
    onPointerCancel: (event: PointerEvent<HTMLElement>) => { if (gesture.current?.pointerId === event.pointerId) cancel(); } };
}
