type Point = { x: number; y: number };

const presets = new Set(['bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5']);

export function bentConnectorHit(shape: {
  geometry: string; x: number; y: number; w: number; h: number;
  stroke?: { width: number; join?: 'round' | 'bevel' | 'miter' } | null;
  path: readonly { type: string; x?: number; y?: number }[];
}) {
  if (!presets.has(shape.geometry)) return undefined;
  const segments: [Point, Point][] = [];
  const joins: [Point, Point, Point][] = [];
  if (!shape.stroke || !Number.isFinite(shape.stroke.width) || shape.stroke.width <= 0) {
    return { segments, joins, tolerance: 0, miterRadius: 0 };
  }
  const radius = shape.stroke.width / 2;
  const miterRadius = shape.stroke.join === 'round' || shape.stroke.join === 'bevel' ? 0 : radius;
  let previous: Point | undefined;
  let incoming: Point | undefined;
  for (const command of shape.path) {
    if ((command.type !== 'move' && command.type !== 'line') ||
      command.x === undefined || command.y === undefined) continue;
    const point = { x: shape.x + command.x * shape.w, y: shape.y + command.y * shape.h };
    if (command.type === 'move') {
      incoming = undefined;
    } else if (previous) {
      segments.push([previous, point]);
      if (point.x === previous.x && point.y === previous.y) continue;
      if (miterRadius > 0 && incoming) joins.push([incoming, previous, point]);
      incoming = previous;
    }
    previous = point;
  }
  return { segments, joins, tolerance: Math.max(3, radius), miterRadius };
}

export function containsBentConnector(hit: NonNullable<ReturnType<typeof bentConnectorHit>>, point: Point): boolean {
  return hit.segments.some(([a, b]) => {
    const dx = Math.max(Math.min(a.x, b.x) - point.x, point.x - Math.max(a.x, b.x), 0);
    const dy = Math.max(Math.min(a.y, b.y) - point.y, point.y - Math.max(a.y, b.y), 0);
    return dx * dx + dy * dy <= hit.tolerance * hit.tolerance;
  }) || hit.joins.some(([a, b, c]) => {
    const dx = point.x - b.x;
    const dy = point.y - b.y;
    if (Math.abs(dx) > hit.miterRadius || Math.abs(dy) > hit.miterRadius) return false;
    return (a.y === b.y && c.x === b.x && (a.x - b.x) * dx < 0 && (c.y - b.y) * dy < 0) ||
      (a.x === b.x && c.y === b.y && (c.x - b.x) * dx < 0 && (a.y - b.y) * dy < 0);
  });
}
