import { useEffect, useRef } from 'react';
import { modelPointToCanvas } from '@betteroffice/vsdx';
import type { ModelPoint, PageDisplayList } from '@betteroffice/vsdx';
import { pageModelBounds } from '../../snap';

export const RULER_SIZE = 20;
const MARKER_STROKE = '#0f6cbd';

const minorSteps = [0.125, 0.25, 0.5, 1];
const majorSteps = [1, 2, 5, 10];

const pickStep = (steps: readonly number[], pxPerInch: number, minPx: number): number | null =>
  steps.find((step) => step * pxPerInch >= minPx) ?? null;

/** Inch ruler above the page; sticky so it tracks zoom, scroll and the pointer. */
export function RulerTop({ frame, zoom, marker }: { frame: PageDisplayList; zoom: number; marker: ModelPoint | null }) {
  const ref = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const context = canvas.getContext('2d');
    if (!context) return;
    const dpr = window.devicePixelRatio || 1;
    const width = Math.max(1, Math.round(frame.width * zoom));
    canvas.width = Math.round(width * dpr);
    canvas.height = Math.round(RULER_SIZE * dpr);
    canvas.style.width = `${width}px`;
    canvas.style.height = `${RULER_SIZE}px`;
    const bounds = pageModelBounds(frame);
    const pxPerInch = 96 * zoom;
    const minor = pickStep(minorSteps, pxPerInch, 5);
    const major = pickStep(majorSteps, pxPerInch, 40) ?? 10;
    context.setTransform(dpr, 0, 0, dpr, 0, 0);
    context.fillStyle = '#f3f5f8';
    context.fillRect(0, 0, width, RULER_SIZE);
    context.strokeStyle = '#8a94a6';
    context.fillStyle = '#424242';
    context.lineWidth = 1;
    context.font = '9px sans-serif';
    context.textBaseline = 'top';
    const span = bounds.maxX - bounds.minX;
    for (let inch = 0; inch <= span + 1e-9; inch += minor ?? major) {
      const value = Math.round(inch * 1000) / 1000;
      const x = modelPointToCanvas(frame.paintTransform, bounds.minX + value, (bounds.minY + bounds.maxY) / 2).x * zoom;
      const isMajor = Math.abs(value / major - Math.round(value / major)) < 1e-9;
      const tall = isMajor ? 12 : value * 2 % 1 === 0 ? 8 : 5;
      context.beginPath();
      context.moveTo(x + 0.5, RULER_SIZE - tall);
      context.lineTo(x + 0.5, RULER_SIZE);
      context.stroke();
      if (isMajor) context.fillText(String(Math.round(value)), x + 3, 2);
    }
    context.strokeStyle = '#c8c6c4';
    context.beginPath();
    context.moveTo(0, RULER_SIZE - 0.5);
    context.lineTo(width, RULER_SIZE - 0.5);
    context.stroke();
    if (marker) {
      context.strokeStyle = MARKER_STROKE;
      context.beginPath();
      context.moveTo(marker.x * zoom + 0.5, 0);
      context.lineTo(marker.x * zoom + 0.5, RULER_SIZE);
      context.stroke();
    }
  }, [frame, zoom, marker]);
  const width = Math.max(1, Math.round(frame.width * zoom));
  return <canvas ref={ref} aria-hidden="true" style={{ display: 'block', width: `${width}px`, height: `${RULER_SIZE}px` }} />;
}

/** Inch ruler left of the page; sticky so it tracks zoom, scroll and the pointer. */
export function RulerLeft({ frame, zoom, marker }: { frame: PageDisplayList; zoom: number; marker: ModelPoint | null }) {
  const ref = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const context = canvas.getContext('2d');
    if (!context) return;
    const dpr = window.devicePixelRatio || 1;
    const height = Math.max(1, Math.round(frame.height * zoom));
    canvas.width = Math.round(RULER_SIZE * dpr);
    canvas.height = Math.round(height * dpr);
    canvas.style.width = `${RULER_SIZE}px`;
    canvas.style.height = `${height}px`;
    const bounds = pageModelBounds(frame);
    const pxPerInch = 96 * zoom;
    const minor = pickStep(minorSteps, pxPerInch, 5);
    const major = pickStep(majorSteps, pxPerInch, 40) ?? 10;
    context.setTransform(dpr, 0, 0, dpr, 0, 0);
    context.fillStyle = '#f3f5f8';
    context.fillRect(0, 0, RULER_SIZE, height);
    context.strokeStyle = '#8a94a6';
    context.fillStyle = '#424242';
    context.lineWidth = 1;
    context.font = '9px sans-serif';
    context.textBaseline = 'top';
    const span = bounds.maxY - bounds.minY;
    const midX = (bounds.minX + bounds.maxX) / 2;
    for (let inch = 0; inch <= span + 1e-9; inch += minor ?? major) {
      const value = Math.round(inch * 1000) / 1000;
      const y = modelPointToCanvas(frame.paintTransform, midX, bounds.maxY - value).y * zoom;
      const isMajor = Math.abs(value / major - Math.round(value / major)) < 1e-9;
      const wide = isMajor ? 12 : value * 2 % 1 === 0 ? 8 : 5;
      context.beginPath();
      context.moveTo(RULER_SIZE - wide, y + 0.5);
      context.lineTo(RULER_SIZE, y + 0.5);
      context.stroke();
      if (isMajor && pxPerInch >= 48) context.fillText(String(Math.round(value)), 2, y + 2);
    }
    context.strokeStyle = '#c8c6c4';
    context.beginPath();
    context.moveTo(RULER_SIZE - 0.5, 0);
    context.lineTo(RULER_SIZE - 0.5, height);
    context.stroke();
    if (marker) {
      context.strokeStyle = MARKER_STROKE;
      context.beginPath();
      context.moveTo(0, marker.y * zoom + 0.5);
      context.lineTo(RULER_SIZE, marker.y * zoom + 0.5);
      context.stroke();
    }
  }, [frame, zoom, marker]);
  const height = Math.max(1, Math.round(frame.height * zoom));
  return <canvas ref={ref} aria-hidden="true" style={{ display: 'block', width: `${RULER_SIZE}px`, height: `${height}px` }} />;
}
