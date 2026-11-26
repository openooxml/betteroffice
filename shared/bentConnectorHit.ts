type Point = { x: number; y: number };

const presets = new Set(['bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5']);

export function bentConnectorHit(shape: {
  geometry: string; x: number; y: number; w: number; h: number;
  stroke?: { width: number } | null;
  path: readonly { type: string; x?: number; y?: number }[];
}) {
  if (!presets.has(shape.geometry)) return undefined;
  const segments: [Point, Point][] = [];
  if (!shape.stroke || !Number.isFinite(shape.stroke.width) || shape.stroke.width <= 0) {
    return { segments, tolerance: 0 };
  }
  let previous: Point | undefined;
  for (const command of shape.path) {
    if ((command.type !== 'move' && command.type !== 'line') ||
      command.x === undefined || command.y === undefined) continue;
    const point = { x: shape.x + command.x * shape.w, y: shape.y + command.y * shape.h };
    if (command.type === 'line' && previous) segments.push([previous, point]);
    previous = point;
  }
  return { segments, tolerance: Math.max(3, shape.stroke.width / 2) };
}

export function containsBentConnector(hit: NonNullable<ReturnType<typeof bentConnectorHit>>, point: Point): boolean {
  return hit.segments.some(([a, b]) => {
    const dx = Math.max(Math.min(a.x, b.x) - point.x, point.x - Math.max(a.x, b.x), 0);
    const dy = Math.max(Math.min(a.y, b.y) - point.y, point.y - Math.max(a.y, b.y), 0);
    return dx * dx + dy * dy <= hit.tolerance * hit.tolerance;
  });
}
