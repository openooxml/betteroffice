export interface CanvasPixelComparison {
  differingPixels: number;
  maxChannelDelta: number;
}

export function compareCanvasPixels(
  local: HTMLCanvasElement,
  worker: HTMLCanvasElement
): CanvasPixelComparison {
  if (local.width !== worker.width || local.height !== worker.height) {
    throw new Error(`Canvas sizes differ: ${local.width}x${local.height} vs ${worker.width}x${worker.height}`);
  }
  const a = local.getContext('2d')!.getImageData(0, 0, local.width, local.height).data;
  const b = worker.getContext('2d')!.getImageData(0, 0, worker.width, worker.height).data;
  let differingPixels = 0;
  let maxChannelDelta = 0;
  for (let pixel = 0; pixel < a.length; pixel += 4) {
    let differs = false;
    for (let channel = 0; channel < 4; channel += 1) {
      const delta = Math.abs(a[pixel + channel]! - b[pixel + channel]!);
      if (delta !== 0) differs = true;
      maxChannelDelta = Math.max(maxChannelDelta, delta);
    }
    if (differs) differingPixels += 1;
  }
  return { differingPixels, maxChannelDelta };
}
