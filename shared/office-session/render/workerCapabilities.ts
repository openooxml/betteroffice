export type WorkerRenderCapability =
  | 'Worker'
  | 'OffscreenCanvas'
  | 'HTMLCanvasElement.prototype.transferControlToOffscreen'
  | 'createImageBitmap';

export function missingWorkerRenderCapabilities(): readonly WorkerRenderCapability[] {
  const missing: WorkerRenderCapability[] = [];
  if (typeof Worker !== 'function') missing.push('Worker');
  if (typeof OffscreenCanvas !== 'function') missing.push('OffscreenCanvas');
  if (typeof HTMLCanvasElement !== 'function' || typeof HTMLCanvasElement.prototype.transferControlToOffscreen !== 'function') {
    missing.push('HTMLCanvasElement.prototype.transferControlToOffscreen');
  }
  if (typeof createImageBitmap !== 'function') missing.push('createImageBitmap');
  return missing;
}
