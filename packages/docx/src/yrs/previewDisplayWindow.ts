import type { ResidentEngineSession } from './residentEngineSession';

/** @internal */
export function setFinalPreviewDisplayWindow(
  engine: Pick<ResidentEngineSession, 'setDisplayWindow'>,
  window: [number, number] | undefined,
  partialDocument: boolean | undefined,
  provisional: boolean,
  pageCount: number | null
): void {
  if (partialDocument !== true || provisional) return;
  if (pageCount === null) {
    engine.setDisplayWindow(0, 0);
    return;
  }
  const [start, end] = window ?? [0, pageCount];
  engine.setDisplayWindow(start, Math.min(end, Math.max(start, pageCount - 2)));
}
