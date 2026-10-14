export interface SlideFit { scale: number; dpr: number; generation: number }

export function fitSlideScale(width: number, height: number,
  slide: { width: number; height: number }, padding = 40): number {
  if (width <= 0 || height <= 0 || slide.width <= 0 || slide.height <= 0) return 0;
  return Math.min(1, (width > padding ? width - padding : width) / slide.width,
    (height > padding ? height - padding : height) / slide.height);
}
