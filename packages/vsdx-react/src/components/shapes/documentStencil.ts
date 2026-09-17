import type { Affine, DocumentMaster, FormulaShapeDraft, PageDisplayList, PagePrimitive } from '@betteroffice/vsdx';
import type { StandardShape } from './shapeLibrary';

function placementCell(name: string, formula: string): FormulaShapeDraft['cells'][number] {
  return { locator: { cellName: name }, name, formula };
}

/** Writes placement only: size and local pin stay inherited from the master. */
export function documentMasterDraft(master: number, x: number, y: number): FormulaShapeDraft {
  return {
    master,
    cells: [
      placementCell('PinX', String(Number.isFinite(x) ? x : 0)),
      placementCell('PinY', String(Number.isFinite(y) ? y : 0)),
    ],
  };
}

function identity(): Affine {
  return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
}

function compose(outer: Affine, inner: Affine): Affine {
  return {
    a: outer.a * inner.a + outer.c * inner.b,
    b: outer.b * inner.a + outer.d * inner.b,
    c: outer.a * inner.c + outer.c * inner.d,
    d: outer.b * inner.c + outer.d * inner.d,
    e: outer.a * inner.e + outer.c * inner.f + outer.e,
    f: outer.b * inner.e + outer.d * inner.f + outer.f,
  };
}

function apply(transform: Affine, x: number, y: number): [number, number] {
  return [transform.a * x + transform.c * y + transform.e, transform.b * x + transform.d * y + transform.f];
}

function coordinate(value: unknown): number | null {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function formatNumber(value: number): string {
  return String(Math.abs(value) < 0.0005 ? 0 : Math.round(value * 1000) / 1000);
}

function collectPaths(primitives: readonly PagePrimitive[], transform: Affine, out: string[][]): void {
  for (const primitive of primitives) {
    if (primitive.kind === 'group') {
      collectPaths(primitive.primitives, compose(transform, primitive.transform ?? identity()), out);
      continue;
    }
    if (primitive.kind !== 'shape') continue;
    const local = compose(transform, primitive.transform ?? identity());
    const segments: string[] = [];
    for (const command of primitive.path) {
      if (command.type === 'close') { segments.push('Z'); continue; }
      const x = coordinate(command.x);
      const y = coordinate(command.y);
      if (x === null || y === null) continue;
      const [tx, ty] = apply(local, x, y);
      if (command.type === 'move') segments.push(`M ${formatNumber(tx)} ${formatNumber(ty)}`);
      else if (command.type === 'line') segments.push(`L ${formatNumber(tx)} ${formatNumber(ty)}`);
      else if (command.type === 'quad') {
        const cpx = coordinate(command.cpx);
        const cpy = coordinate(command.cpy);
        if (cpx === null || cpy === null) continue;
        const [cx, cy] = apply(local, cpx, cpy);
        segments.push(`Q ${formatNumber(cx)} ${formatNumber(cy)} ${formatNumber(tx)} ${formatNumber(ty)}`);
      } else if (command.type === 'cubic') {
        const cp1x = coordinate(command.cp1x);
        const cp1y = coordinate(command.cp1y);
        const cp2x = coordinate(command.cp2x);
        const cp2y = coordinate(command.cp2y);
        if (cp1x === null || cp1y === null || cp2x === null || cp2y === null) continue;
        const [c1x, c1y] = apply(local, cp1x, cp1y);
        const [c2x, c2y] = apply(local, cp2x, cp2y);
        segments.push(`C ${formatNumber(c1x)} ${formatNumber(c1y)} ${formatNumber(c2x)} ${formatNumber(c2y)} ${formatNumber(tx)} ${formatNumber(ty)}`);
      }
    }
    if (segments.length) out.push(segments);
  }
}

interface PreviewBounds { minX: number; minY: number; scale: number; offsetX: number; offsetY: number }

function pathBounds(paths: string[][]): PreviewBounds | null {
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const segments of paths) {
    for (const segment of segments) {
      const numbers = segment.split(' ').slice(1).map(Number).filter(Number.isFinite);
      for (let index = 0; index + 1 < numbers.length; index += 2) {
        const [x, y] = [numbers[index], numbers[index + 1]];
        if (x < minX) minX = x;
        if (y < minY) minY = y;
        if (x > maxX) maxX = x;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (!Number.isFinite(minX) || !Number.isFinite(minY) || maxX < minX || maxY < minY) return null;
  const width = maxX - minX;
  const height = maxY - minY;
  const scale = Math.max(width, height) || 1;
  return { minX, minY, scale, offsetX: (1 - width / scale) / 2, offsetY: (1 - height / scale) / 2 };
}

export function masterPreviewPath(list: PageDisplayList): string {
  const paint = list.paintTransform ?? identity();
  const paths: string[][] = [];
  collectPaths(list.primitives, paint, paths);
  const bounds = pathBounds(paths);
  if (!bounds) return '';
  return paths
    .map((segments) => segments.map((segment) => {
      if (segment === 'Z') return 'Z';
      const [command, ...rest] = segment.split(' ');
      const numbers: number[] = [];
      for (let index = 0; index + 1 < rest.length; index += 2) {
        numbers.push(
          (Number(rest[index]) - bounds.minX) / bounds.scale + bounds.offsetX,
          (Number(rest[index + 1]) - bounds.minY) / bounds.scale + bounds.offsetY,
        );
      }
      return `${command} ${numbers.map(formatNumber).join(' ')}`;
    }).join(' '))
    .join(' ');
}

const PIXELS_PER_INCH = 96;

function masterSize(display: PageDisplayList | null): { width: number; height: number } {
  const inches = (value: number | undefined) => (Number.isFinite(value) && (value as number) > 0 ? (value as number) / PIXELS_PER_INCH : 1);
  return { width: inches(display?.width), height: inches(display?.height) };
}

export function documentStencilEntries(masters: readonly DocumentMaster[]): StandardShape[] {
  return masters.map((master) => ({
    id: `document-master-${master.id}`,
    nameKey: 'shapesPanel.shape.documentShape',
    label: master.name ?? `#${master.id}`,
    preview: master.display ? masterPreviewPath(master.display) : '',
    defaultSize: masterSize(master.display),
    draft: (x, y) => documentMasterDraft(master.id, x, y),
  } satisfies StandardShape));
}
