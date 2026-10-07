import { MAX_TIFF_BYTES, isTiff } from '../../../../shared/media';
import { decodeTiffImage } from '../wasm/loader';
import { rasterImageSize, type ImageSize } from './imageSize';

const MAX_BITMAP_PIXELS = 33_554_432;
const SVG_MEDIA_TYPE = 'image/svg+xml';
/** Matches `MAX_SVG_BYTES` in pptx-raster: the largest SVG either backend decodes. */
const MAX_SVG_BYTES = 4_194_304;
const MAX_IMAGE_DIMENSION = 4096;
const decodeScales = new WeakMap<CanvasImageSource, { x: number; y: number }>();

export interface PresentationImageDecodeOptions {
  maxDimension?: number;
}

export function imageDecodeScale(source: CanvasImageSource): { x: number; y: number } {
  return decodeScales.get(source) ?? { x: 1, y: 1 };
}

export function setImageDecodeScale(source: CanvasImageSource, scale: { x: number; y: number }): void {
  decodeScales.set(source, scale);
}

/** Convert presentation image formats that browsers cannot decode. */
export function presentationImageBlob(bytes: Uint8Array): Blob {
  if (isTiff(bytes)) {
    if (bytes.byteLength > MAX_TIFF_BYTES) {
      throw new Error('TIFF image exceeds the browser transfer budget');
    }
    return new Blob([decodeTiffImage(bytes).slice()], { type: 'image/png' });
  }
  const bitmap = wmfBitmap(bytes) ?? emfBitmap(bytes);
  if (bitmap) return new Blob([bitmap], { type: 'image/bmp' });
  return svgBlob(bytes) ?? new Blob([bytes.slice()]);
}

/** Media the `<img>` element must decode, because `createImageBitmap` is not portable for it. */
export function needsElementDecode(blob: Blob): boolean {
  return blob.type === SVG_MEDIA_TYPE;
}

/**
 * Decode presentation media for the canvas backend, rejecting with
 * `errorMessage` when the browser will not decode it.
 */
export async function decodePresentationImage(
  bytes: Uint8Array,
  errorMessage: string,
  options: PresentationImageDecodeOptions = {}
): Promise<CanvasImageSource> {
  const maximum = options.maxDimension ?? MAX_IMAGE_DIMENSION;
  if (!Number.isSafeInteger(maximum) || maximum < 1) throw new Error('invalid image dimension limit');
  let blob = presentationImageBlob(bytes);
  const svg = needsElementDecode(blob) ? await blob.text() : undefined;
  const original = svg === undefined ? rasterImageSize(new Uint8Array(await blob.arrayBuffer())) : svgImageSize(svg);
  if (!original) throw new Error(errorMessage);
  const factor = Math.min(1, maximum / Math.max(original.width, original.height));
  const size = {
    width: factor === 1 ? original.width : Math.max(1, Math.floor(original.width * factor)),
    height: factor === 1 ? original.height : Math.max(1, Math.floor(original.height * factor)),
  };
  const remember = (source: CanvasImageSource) => {
    if (factor < 1) setImageDecodeScale(source, { x: original.width / size.width, y: original.height / size.height });
    return source;
  };
  if (typeof createImageBitmap === 'function' && svg === undefined) {
    return remember(await (factor < 1
      ? createImageBitmap(blob, { resizeWidth: size.width, resizeHeight: size.height })
      : createImageBitmap(blob)));
  }
  if (svg !== undefined) blob = new Blob([sizedSvg(svg, original, size)], { type: SVG_MEDIA_TYPE });
  else if (factor < 1) throw new Error(errorMessage);
  const url = URL.createObjectURL(blob);
  try {
    return remember(await new Promise<HTMLImageElement>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error(errorMessage));
      image.src = url;
    }));
  } finally {
    URL.revokeObjectURL(url);
  }
}

function svgImageSize(text: string): ImageSize | undefined {
  const tag = rootTag(text);
  if (!tag) return;
  const root = text.slice(tag.start, tag.end);
  const style = root.match(/\sstyle\s*=\s*(["'])(.*?)\1/s)?.[2] ?? '';
  const dimension = (name: string) => {
    const attribute = root.match(new RegExp(`\\s${name}\\s*=\\s*(["'])(.*?)\\1`, 's'))?.[2];
    const rules = [...style.matchAll(new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]+)`, 'gi'))];
    const value = (rules[rules.length - 1]?.[1] ?? attribute ?? '').replace(/\s*!important\s*$/, '').trim();
    const match = value.match(/^([+\d.eE-]+)\s*(px|pt|pc|in|cm|mm)?$/);
    const units: Record<string, number> = { px: 1, pt: 96 / 72, pc: 16, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 };
    const pixels = Number(match?.[1]) * (units[match?.[2] ?? 'px'] ?? 1);
    return Number.isFinite(pixels) && pixels > 0 ? pixels : undefined;
  };
  const box = root.match(/\sviewBox\s*=\s*(["'])\s*[-+.\deE]+[\s,]+[-+.\deE]+[\s,]+([-+.\deE]+)[\s,]+([-+.\deE]+)\s*\1/);
  const boxWidth = Number(box?.[2]);
  const boxHeight = Number(box?.[3]);
  let width = dimension('width');
  let height = dimension('height');
  if (boxWidth > 0 && boxHeight > 0) {
    if (width && !height) height = width * boxHeight / boxWidth;
    else if (height && !width) width = height * boxWidth / boxHeight;
  }
  width ??= 300;
  height ??= 150;
  return Number.isFinite(width) && Number.isFinite(height) ? { width, height } : undefined;
}

function sizedSvg(text: string, original: ImageSize, size: ImageSize): string {
  const tag = rootTag(text)!;
  let root = text.slice(tag.start, tag.end).replace(/\s(?:width|height)\s*=\s*(["']).*?\1/gs, '');
  const selfClosing = root.endsWith('/');
  if (selfClosing) root = root.slice(0, -1);
  const dimensions = `width:${size.width}px!important;height:${size.height}px!important`;
  if (/\sstyle\s*=/.test(root))
    root = root.replace(/(\sstyle\s*=\s*)(["'])(.*?)\2/s, (_, prefix, quote, style) => `${prefix}${quote}${style};${dimensions}${quote}`);
  else root += ` style="${dimensions}"`;
  if (!/\sviewBox\s*=/.test(root)) root += ` viewBox="0 0 ${original.width} ${original.height}"`;
  root += ` width="${size.width}" height="${size.height}"`;
  return text.slice(0, tag.start) + root + (selfClosing ? '/' : '') + text.slice(tag.end);
}

/** An SVG a browser will decode: typed, and sized where only a `viewBox` says how big it is. */
function svgBlob(bytes: Uint8Array): Blob | undefined {
  if (bytes.byteLength > MAX_SVG_BYTES) return;
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^\ufeff/, '');
  } catch {
    return;
  }
  if (!/^\s*</.test(text) || !text.slice(0, 1024).includes('<svg')) return;
  return new Blob([withIntrinsicSize(text)], { type: SVG_MEDIA_TYPE });
}

function withIntrinsicSize(text: string): string {
  const tag = rootTag(text);
  if (!tag || /\s(?:width|height)\s*=/.test(text.slice(tag.start, tag.end))) return text;
  const box = text
    .slice(tag.start, tag.end)
    .match(/\sviewBox\s*=\s*(["'])\s*[-+.\deE]+[\s,]+[-+.\deE]+[\s,]+([-+.\deE]+)[\s,]+([-+.\deE]+)\s*\1/);
  const [width, height] = [Number(box?.[2]), Number(box?.[3])];
  if (!(width > 0) || !(height > 0)) return text;
  const at = tag.start + 4;
  return `${text.slice(0, at)} width="${width}" height="${height}"${text.slice(at)}`;
}

/** The `<svg>` start tag's bounds, or nothing unless it is the first element. */
function rootTag(text: string): { start: number; end: number } | undefined {
  let start = 0;
  while (start < text.length) {
    if (/\s/.test(text[start])) start += 1;
    else if (text.startsWith('<!--', start)) start = past(text.indexOf('-->', start + 4), 3);
    else if (text.startsWith('<?', start)) start = past(text.indexOf('?>', start + 2), 2);
    else if (text.startsWith('<!', start)) start = past(declarationEnd(text, start + 2), 1);
    else break;
    if (start < 0) return;
  }
  if (!/^<svg[\s/>]/.test(text.slice(start, start + 5))) return;
  const end = declarationEnd(text, start + 4);
  return end < 0 ? undefined : { start, end };
}

function past(at: number, length: number): number {
  return at < 0 ? -1 : at + length;
}

/**
 * The `>` closing a markup declaration or start tag, past `>` inside a quoted
 * value or a DOCTYPE's internal subset, and past the comments and processing
 * instructions that subset may hold.
 */
function declarationEnd(text: string, from: number): number {
  let quote = '';
  let subset = false;
  for (let index = from; index < text.length; index += 1) {
    const character = text[index];
    if (quote) {
      if (character === quote) quote = '';
    } else if (subset && text.startsWith('<!--', index)) {
      index = past(text.indexOf('-->', index + 4), 2);
      if (index < 0) return -1;
    } else if (subset && text.startsWith('<?', index)) {
      index = past(text.indexOf('?>', index + 2), 1);
      if (index < 0) return -1;
    } else if (character === '"' || character === "'") quote = character;
    else if (character === '[') subset = true;
    else if (character === ']') subset = false;
    else if (character === '>' && !subset) return index;
  }
  return -1;
}

/** An EMF whose only ink is one unscaled `EMR_STRETCHDIBITS` filling its bounds. */
function emfBitmap(bytes: Uint8Array): Uint8Array<ArrayBuffer> | undefined {
  const data = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (
    data.byteLength < 108 ||
    data.getUint32(0, true) !== 1 ||
    data.getUint32(40, true) !== 0x464d4520
  ) return;
  const header = data.getUint32(4, true);
  if (header < 88 || header % 4 || header > data.byteLength - 8) return;
  const width = data.getInt32(16, true) - data.getInt32(8, true) + 1;
  const height = data.getInt32(20, true) - data.getInt32(12, true) + 1;
  const size = data.getUint32(header + 4, true);
  if (
    data.getUint32(header, true) !== 81 ||
    size < 80 ||
    size > data.byteLength - header - 8
  ) return;
  const at = (offset: number) => data.getInt32(header + offset, true);
  const bmi = data.getUint32(header + 48, true);
  const bits = data.getUint32(header + 56, true);
  const bitsBytes = data.getUint32(header + 60, true);
  if (
    at(24) !== data.getInt32(8, true) || at(28) !== data.getInt32(12, true) ||
    at(72) !== width || at(76) !== height ||
    at(32) !== 0 || at(36) !== 0 || at(40) !== width || at(44) !== height ||
    data.getUint32(header + 64, true) !== 0 ||
    data.getUint32(header + 68, true) !== 0x00cc0020 ||
    bmi < 80 || bmi + data.getUint32(header + 52, true) !== bits ||
    bits + bitsBytes !== size
  ) return;
  const end = header + size;
  if (
    data.getUint32(end, true) !== 14 ||
    data.getUint32(end + 4, true) !== data.byteLength - end
  ) return;
  return dibBitmap(bytes.subarray(header + bmi, end), width, height);
}

function wmfBitmap(bytes: Uint8Array): Uint8Array<ArrayBuffer> | undefined {
  const data = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (data.byteLength < 18) return;
  const start = data.getUint32(0, true) === 0x9ac6cdd7 ? 22 : 0;
  if (
    data.byteLength < start + 18 ||
    ![1, 2].includes(data.getUint16(start, true)) ||
    data.getUint16(start + 2, true) !== 9 ||
    data.getUint16(start + 4, true) !== 0x0300 ||
    data.getUint32(start + 6, true) * 2 !== data.byteLength - start
  ) return;
  let originX = 0;
  let originY = 0;
  let windowWidth = 0;
  let windowHeight = 0;
  let anisotropic = false;
  let bitmap: Uint8Array<ArrayBuffer> | undefined;
  for (let offset = start + 18; offset + 6 <= data.byteLength;) {
    const size = data.getUint32(offset, true) * 2;
    const command = data.getUint16(offset + 4, true);
    if (size < 6 || size > data.byteLength - offset) return;
    if (command === 0) return size === 6 && offset + size === data.byteLength ? bitmap : undefined;
    if (bitmap) return;
    if (command === 0x0103) {
      if (size < 8 || data.getUint16(offset + 6, true) !== 8) return;
      anisotropic = true;
    } else if (command === 0x0107) {
      if (size < 8 || ![1, 2, 3, 4].includes(data.getUint16(offset + 6, true))) return;
    } else if (command === 0x020b || command === 0x020c) {
      if (size !== 10) return;
      const y = data.getInt16(offset + 6, true);
      const x = data.getInt16(offset + 8, true);
      if (command === 0x020b) {
        originX = x;
        originY = y;
      } else {
        if (!anisotropic) return;
        windowWidth = x;
        windowHeight = y;
      }
    } else if (command === 0x0b41) {
      if (
        size < 66 ||
        data.getUint32(offset + 6, true) !== 0x00cc0020 ||
        data.getInt16(offset + 14, true) !== 0 ||
        data.getInt16(offset + 16, true) !== 0 ||
        data.getInt16(offset + 18, true) !== windowHeight ||
        data.getInt16(offset + 20, true) !== windowWidth ||
        data.getInt16(offset + 22, true) !== originY ||
        data.getInt16(offset + 24, true) !== originX ||
        !windowWidth || !windowHeight
      ) return;
      bitmap = dibBitmap(
        bytes.subarray(offset + 26, offset + size),
        data.getInt16(offset + 12, true),
        data.getInt16(offset + 10, true)
      );
      if (!bitmap) return;
    } else return;
    offset += size;
  }
}

function dibBitmap(
  bytes: Uint8Array,
  sourceWidth: number,
  sourceHeight: number
): Uint8Array<ArrayBuffer> | undefined {
  if (bytes.length < 40) return;
  const data = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = data.getInt32(4, true);
  const height = data.getInt32(8, true);
  const bits = data.getUint16(14, true);
  if (
    data.getUint32(0, true) !== 40 ||
    width < 1 || !height || width * Math.abs(height) > MAX_BITMAP_PIXELS ||
    sourceWidth !== width || sourceHeight !== Math.abs(height) ||
    data.getUint16(12, true) !== 1 ||
    ![24, 32].includes(bits) ||
    data.getUint32(16, true) !== 0 ||
    data.getUint32(32, true) !== 0
  ) return;
  const pixelBytes = Math.ceil(width * bits / 32) * 4 * Math.abs(height);
  const declaredBytes = data.getUint32(20, true);
  if (40 + pixelBytes !== bytes.length || (declaredBytes !== 0 && declaredBytes !== pixelBytes)) return;
  const bitmap = new Uint8Array(14 + bytes.length);
  const header = new DataView(bitmap.buffer);
  header.setUint16(0, 0x4d42, true);
  header.setUint32(2, bitmap.length, true);
  header.setUint32(10, 54, true);
  bitmap.set(bytes, 14);
  return bitmap;
}
