import { MIN_ZOOM, MAX_ZOOM } from './commands/evaluate';

export function initialPptxZoom(value: number | undefined): number | 'fit' {
  return value !== undefined && Number.isFinite(value) && value > 0
    ? Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, value)) : 'fit';
}
