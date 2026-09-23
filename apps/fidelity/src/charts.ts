const NS = 'http://www.w3.org/2000/svg';

type Attrs = Record<string, string | number | undefined>;
type Child = Node | string | null | undefined | false;

export function s<K extends keyof SVGElementTagNameMap>(
  tag: K,
  attrs: Attrs = {},
  ...children: Child[]
): SVGElementTagNameMap[K] {
  const node = document.createElementNS(NS, tag);
  for (const [key, value] of Object.entries(attrs))
    if (value !== undefined) node.setAttribute(key, String(value));
  node.append(...children.filter((child): child is Node | string => !!child));
  return node;
}

export function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Attrs = {},
  ...children: Child[]
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs))
    if (value !== undefined) node.setAttribute(key, String(value));
  node.append(...children.filter((child): child is Node | string => !!child));
  return node;
}

export const linear =
  (d0: number, d1: number, r0: number, r1: number) =>
  (value: number): number =>
    r0 + ((value - d0) / (d1 - d0 || 1)) * (r1 - r0);

export const logarithmic = (d0: number, d1: number, r0: number, r1: number) => {
  const scale = linear(Math.log10(d0), Math.log10(d1), r0, r1);
  return (value: number): number => scale(Math.log10(Math.max(value, d0)));
};

/** 1-3-10 ticks: dense enough to read a log axis, sparse enough to label. */
export function logTicks(d0: number, d1: number): number[] {
  const ticks: number[] = [];
  const steps = Math.log10(d1 / d0) > 4 ? [1] : [1, 3];
  for (let exponent = Math.floor(Math.log10(d0)); exponent <= Math.ceil(Math.log10(d1)); exponent += 1)
    for (const step of steps) {
      const value = step * 10 ** exponent;
      if (value >= d0 * 0.999 && value <= d1 * 1.001) ticks.push(value);
    }
  return ticks;
}

export function duration(ms: number): string {
  if (ms >= 1000) return `${+(ms / 1000).toFixed(ms >= 10_000 ? 0 : 1)} s`;
  if (ms >= 10) return `${Math.round(ms)} ms`;
  if (ms >= 1) return `${+ms.toFixed(1)} ms`;
  if (ms >= 0.1) return `${+ms.toFixed(2)} ms`;
  return `${Math.round(ms * 1000)} µs`;
}

const round = (value: number): number => +value.toPrecision(3);

/** The nearest 1-3-10 step at or below, and at or above, a value. */
const floorStep = (value: number) => {
  const base = 10 ** Math.floor(Math.log10(value));
  return value >= 3 * base ? 3 * base : base;
};
const ceilStep = (value: number) => {
  const base = 10 ** Math.floor(Math.log10(value));
  return value <= base ? base : value <= 3 * base ? 3 * base : 10 * base;
};
const tick = (ms: number): string =>
  ms >= 1000 ? `${round(ms / 1000)} s` : ms >= 1 ? `${round(ms)} ms` : `${round(ms * 1000)} µs`;

let tipNode: HTMLElement | null = null;

export function tip(content: Node | null, x = 0, y = 0): void {
  tipNode ??= document.getElementById('tip');
  if (!tipNode) return;
  if (!content) {
    tipNode.hidden = true;
    return;
  }
  tipNode.replaceChildren(content);
  tipNode.hidden = false;
  const box = tipNode.getBoundingClientRect();
  let left = x + 16;
  let top = y + 16;
  if (left + box.width > window.innerWidth - 8) left = Math.max(8, x - box.width - 16);
  if (top + box.height > window.innerHeight - 8) top = Math.max(8, y - box.height - 16);
  tipNode.style.transform = `translate(${Math.round(left)}px, ${Math.round(top)}px)`;
}

export function tipCard(title: string, rows: [cls: string, label: string, value: string][], foot?: string) {
  return h(
    'div',
    {},
    h('div', { class: 'tip-title' }, title),
    ...rows.map(([cls, label, value]) =>
      h('div', { class: 'tip-row' }, h('i', { class: `swatch ${cls}` }), h('span', {}, label), h('b', {}, value))
    ),
    foot ? h('div', { class: 'tip-foot' }, foot) : null
  );
}

/** Re-renders a chart whenever its container changes width; the result stops watching. */
export function mount(container: HTMLElement, build: (width: number) => Element): () => void {
  let width = -1;
  const draw = () => {
    const next = Math.floor(container.clientWidth);
    if (next === width || next <= 0) return;
    width = next;
    container.replaceChildren(build(next));
  };
  const observer = new ResizeObserver(draw);
  observer.observe(container);
  draw();
  return () => observer.disconnect();
}

export interface SwarmSeries {
  key: string;
  cls: string;
  label: string;
  sub: string;
  hollow?: boolean;
}

export interface SwarmPoint {
  id: string;
  values: Record<string, number | null>;
}

interface Placed {
  id: string;
  value: number;
  x: number;
  y: number;
}

/** Places each dot at the smallest offset that touches no earlier dot. */
function dodge(points: { id: string; value: number; x: number }[], radius: number): Placed[] {
  const distance = 2 * radius + 0.9;
  const placed: Placed[] = [];
  let start = 0;
  for (const point of [...points].sort((left, right) => left.x - right.x)) {
    while (start < placed.length && point.x - placed[start]!.x > distance) start += 1;
    const near = placed.slice(start);
    const candidates = [0];
    for (const other of near) {
      const dy = Math.sqrt(Math.max(0, distance ** 2 - (point.x - other.x) ** 2));
      candidates.push(other.y + dy, other.y - dy);
    }
    const flip = placed.length % 2 ? 1 : -1;
    candidates.sort((left, right) => Math.abs(left) - Math.abs(right) || flip * (left - right));
    const y =
      candidates.find((candidate) =>
        near.every((other) => (point.x - other.x) ** 2 + (candidate - other.y) ** 2 >= distance ** 2 - 1e-6)
      ) ?? 0;
    placed.push({ ...point, y });
  }
  return placed;
}

export function beeswarm(
  points: SwarmPoint[],
  series: SwarmSeries[],
  width: number,
  options: { format: (value: number) => string; onPick: (id: string) => void; describe: (id: string) => Node }
): SVGSVGElement {
  const narrow = width < 560;
  const labelWidth = narrow ? 0 : 148;
  const left = labelWidth + 10;
  const right = width - 14;
  const all = points
    .flatMap((point) => series.map((entry) => point.values[entry.key]))
    .filter((value): value is number => value !== null && value !== undefined)
    .sort((a, b) => a - b);
  const floor = all[Math.floor(all.length * 0.01)] ?? 0.5;
  const low = Math.max(0, Math.floor((floor - 0.01) * 20) / 20);
  const clipped = all.length > 0 && all[0]! < low;
  const lane = clipped ? 24 : 0;
  const scale = linear(low, 1, left + lane, right);
  const x = (value: number) => (value < low ? left + 8 : scale(value));
  const radius = narrow ? 2.3 : points.length > 150 ? 2.9 : 3.4;
  const swarms = series.map((entry) =>
    dodge(
      points
        .filter((point) => typeof point.values[entry.key] === 'number')
        .map((point) => ({ id: point.id, value: point.values[entry.key]!, x: x(point.values[entry.key]!) })),
      radius
    )
  );
  const spread = Math.max(8, ...swarms.flat().map((dot) => Math.abs(dot.y)));
  const header = narrow ? 18 : 0;
  const row = Math.max(narrow ? 34 : 46, 2 * (spread + radius) + 14) + header;
  const top = 6;
  const height = top + series.length * row + 28;
  const svg = s('svg', { width, height, viewBox: `0 0 ${width} ${height}`, class: 'chart swarm', role: 'img' });
  const grid = s('g', { class: 'grid' });
  if (clipped)
    grid.append(
      s('line', { x1: left + lane - 6, x2: left + lane - 6, y1: top, y2: top + series.length * row, class: 'lane' }),
      s('text', { x: left + 8, y: top + series.length * row + 18, class: 'axis', 'text-anchor': 'middle' }, `<${low.toFixed(2)}`)
    );
  for (let value = low; value <= 1.0001; value += 0.05) {
    const major = Math.abs((value * 10) % 1) < 1e-6 || Math.abs((value * 10) % 1) > 1 - 1e-6;
    const px = x(value);
    grid.append(s('line', { x1: px, x2: px, y1: top, y2: top + series.length * row, class: major ? 'major' : '' }));
    if ((major || !narrow) && !(clipped && value === low))
      grid.append(
        s('text', { x: px, y: top + series.length * row + 18, class: 'axis', 'text-anchor': 'middle' }, value.toFixed(2))
      );
  }
  svg.append(grid);
  const layer = s('g');
  const positions = new Map<string, { cx: number; cy: number; cls: string; key: string }[]>();
  series.forEach((entry, index) => {
    const band = top + index * row;
    const cy = band + header + (row - header) / 2;
    if (index % 2 === 0)
      svg.insertBefore(s('rect', { x: 0, y: band, width, height: row, class: 'band' }), grid);
    const values = swarms[index]!.map((dot) => dot.value);
    const mean = values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
    const labelY = narrow ? band + 12 : cy;
    svg.append(
      s(
        'text',
        { x: narrow ? left : 0, y: labelY - (narrow ? 0 : 4), class: `series-label ${entry.cls}` },
        entry.label,
        s('tspan', { class: 'series-sub', dx: narrow ? 6 : 0, x: narrow ? undefined : 0, dy: narrow ? 0 : 15 }, entry.sub)
      )
    );
    if (mean !== null) {
      const mx = x(mean);
      layer.append(
        s('line', { x1: mx, x2: mx, y1: band + header + 3, y2: band + row - 3, class: `mean ${entry.cls}` }),
        s('text', { x: mx + 5, y: band + header + 11, class: `mean-label ${entry.cls}` }, `μ ${options.format(mean)}`)
      );
    }
    const dots = s('g', { class: `dots ${entry.cls}${entry.hollow ? ' hollow' : ''}` });
    for (const dot of swarms[index]!) {
      const cyDot = cy + dot.y;
      dots.append(s('circle', { cx: dot.x.toFixed(1), cy: cyDot.toFixed(1), r: radius, class: 'dot' }));
      const list = positions.get(dot.id) ?? [];
      if (!list.length) positions.set(dot.id, list);
      list.push({ cx: dot.x, cy: cyDot, cls: entry.cls, key: entry.key });
    }
    layer.append(dots);
  });
  svg.append(layer);
  const focus = s('g', { class: 'focus' });
  svg.append(focus);

  let active: string | null = null;
  const nearest = (event: PointerEvent): string | null => {
    const box = svg.getBoundingClientRect();
    const px = event.clientX - box.left;
    const py = event.clientY - box.top;
    let best: string | null = null;
    let bestDistance = 12 ** 2;
    for (const [id, list] of positions)
      for (const dot of list) {
        const distance = (dot.cx - px) ** 2 + (dot.cy - py) ** 2;
        if (distance < bestDistance) [best, bestDistance] = [id, distance];
      }
    return best;
  };
  svg.addEventListener('pointermove', (event) => {
    const id = nearest(event);
    if (id !== active) {
      active = id;
      focus.replaceChildren();
      svg.classList.toggle('focused', !!id);
      if (id) {
        const list = positions.get(id)!;
        if (list.length > 1)
          focus.append(
            s('path', { d: list.map((dot, index) => `${index ? 'L' : 'M'}${dot.cx},${dot.cy}`).join(''), class: 'trace' })
          );
        for (const dot of list)
          focus.append(s('circle', { cx: dot.cx, cy: dot.cy, r: radius + 3.5, class: `ring ${dot.cls}` }));
      }
    }
    if (id) tip(options.describe(id), event.clientX, event.clientY);
    else tip(null);
    svg.style.cursor = id ? 'pointer' : 'default';
  });
  svg.addEventListener('pointerleave', () => {
    active = null;
    focus.replaceChildren();
    svg.classList.remove('focused');
    tip(null);
  });
  svg.addEventListener('click', (event) => {
    const id = nearest(event);
    if (id) options.onPick(id);
  });
  return svg;
}

export interface CdfSeries {
  cls: string;
  label: string;
  values: number[];
}

/** Share of items at or below each value, on a log axis. */
export function cdf(
  series: CdfSeries[],
  width: number,
  options: { height?: number; frame?: number; noun: string }
): SVGSVGElement {
  const height = options.height ?? 250;
  const margin = { top: 14, right: 16, bottom: 30, left: 40 };
  const values = series.flatMap((entry) => entry.values).filter((value) => value > 0);
  if (!values.length)
    return s(
      'svg',
      { width, height: 60, viewBox: `0 0 ${width} 60`, class: 'chart cdf', role: 'img' },
      s('text', { x: 0, y: 30, class: 'axis' }, `No timed ${options.noun}`)
    );
  const d0 = floorStep(Math.min(...values));
  const d1 = ceilStep(Math.max(...values));
  const x = logarithmic(d0, d1, margin.left, width - margin.right);
  const y = linear(0, 1, height - margin.bottom, margin.top);
  const svg = s('svg', { width, height, viewBox: `0 0 ${width} ${height}`, class: 'chart cdf', role: 'img' });
  const grid = s('g', { class: 'grid' });
  const ticks = logTicks(d0, d1);
  const every = width < 420 && ticks.length > 6 ? 2 : 1;
  ticks.forEach((value, index) => {
    const major = Math.abs(Math.log10(value) % 1) < 1e-9;
    grid.append(s('line', { x1: x(value), x2: x(value), y1: margin.top, y2: height - margin.bottom, class: major ? 'major' : '' }));
    if (index % every === 0)
      grid.append(s('text', { x: x(value), y: height - 10, class: 'axis', 'text-anchor': 'middle' }, tick(value)));
  });
  for (const share of [0, 0.25, 0.5, 0.75, 1]) {
    grid.append(s('line', { x1: margin.left, x2: width - margin.right, y1: y(share), y2: y(share), class: share === 0 ? 'major' : '' }));
    grid.append(s('text', { x: margin.left - 8, y: y(share) + 3.5, class: 'axis', 'text-anchor': 'end' }, `${share * 100}%`));
  }
  svg.append(grid);
  if (options.frame) {
    const fx = x(options.frame);
    svg.append(
      s('line', { x1: fx, x2: fx, y1: margin.top - 4, y2: height - margin.bottom, class: 'frame' }),
      s(
        'text',
        { x: fx + 6, y: height - margin.bottom - 8, class: 'frame-label' },
        width < 480 ? '16.7 ms' : '16.7 ms · one 60 Hz frame'
      )
    );
  }
  const sorted = series.map((entry) => [...entry.values].sort((left, right) => left - right));
  series.forEach((entry, index) => {
    const list = sorted[index]!;
    if (!list.length) return;
    const n = list.length;
    let d = `M${x(list[0]!)},${y(0)}`;
    list.forEach((value, rank) => {
      d += `H${x(value).toFixed(1)}V${y((rank + 1) / n).toFixed(1)}`;
    });
    d += `H${width - margin.right}`;
    const area = `${d}V${y(0)}H${x(list[0]!)}Z`;
    const median = list[Math.ceil(n / 2) - 1]!;
    svg.append(
      s('path', { d: area, class: `area ${entry.cls}` }),
      s('path', { d, class: `line ${entry.cls}` }),
      s('circle', { cx: x(median), cy: y(0.5), r: 3.5, class: `median ${entry.cls}` })
    );
  });
  const guide = s('line', { y1: margin.top, y2: height - margin.bottom, class: 'guide', visibility: 'hidden' });
  svg.append(guide);
  const inverse = (px: number) =>
    10 ** (Math.log10(d0) + ((px - margin.left) / (width - margin.right - margin.left)) * (Math.log10(d1) - Math.log10(d0)));
  svg.addEventListener('pointermove', (event) => {
    const box = svg.getBoundingClientRect();
    const px = Math.min(Math.max(event.clientX - box.left, margin.left), width - margin.right);
    const at = inverse(px);
    guide.setAttribute('x1', String(px));
    guide.setAttribute('x2', String(px));
    guide.setAttribute('visibility', 'visible');
    tip(
      tipCard(
        `≤ ${duration(at)}`,
        series.map((entry, index) => {
          const list = sorted[index]!;
          let count = 0;
          while (count < list.length && list[count]! <= at) count += 1;
          return [entry.cls, entry.label, `${Math.round((100 * count) / Math.max(1, list.length))}%`];
        }),
        `share of ${options.noun}`
      ),
      event.clientX,
      event.clientY
    );
  });
  svg.addEventListener('pointerleave', () => {
    guide.setAttribute('visibility', 'hidden');
    tip(null);
  });
  return svg;
}

export interface SpectrumRow {
  name: string;
  count: number;
  p50: number;
  p95: number;
  max: number;
}

/** Median dot and median-to-p95 bar per operation, on a shared log axis. */
export function spectrum(rows: SpectrumRow[], width: number, options: { cls: string; frame: number }): SVGSVGElement {
  const narrow = width < 560;
  const labelWidth = narrow ? 0 : Math.min(210, width * 0.3);
  const rowHeight = narrow ? 38 : 27;
  const margin = { top: 24, right: narrow ? 12 : 74, bottom: 26, left: labelWidth + 12 };
  const height = margin.top + rows.length * rowHeight + margin.bottom;
  const d0 = 0.01;
  const d1 = 1000;
  const x = logarithmic(d0, d1, margin.left, width - margin.right);
  const svg = s('svg', { width, height, viewBox: `0 0 ${width} ${height}`, class: `chart spectrum ${options.cls}`, role: 'img' });
  const grid = s('g', { class: 'grid' });
  for (const value of [0.01, 0.1, 1, 10, 100, 1000]) {
    grid.append(s('line', { x1: x(value), x2: x(value), y1: margin.top - 6, y2: height - margin.bottom, class: 'major' }));
    grid.append(s('text', { x: x(value), y: height - 8, class: 'axis', 'text-anchor': 'middle' }, tick(value)));
  }
  svg.append(grid);
  const fx = x(options.frame);
  svg.append(
    s('rect', { x: fx, y: margin.top - 6, width: width - margin.right - fx, height: rows.length * rowHeight + 6, class: 'over-frame' }),
    s('line', { x1: fx, x2: fx, y1: margin.top - 14, y2: height - margin.bottom, class: 'frame' }),
    s('text', { x: fx + 6, y: margin.top - 8, class: 'frame-label' }, 'one frame')
  );
  rows.forEach((row, index) => {
    const top = margin.top + index * rowHeight;
    const cy = top + (narrow ? 26 : rowHeight / 2);
    const group = s('g', { class: 'op' });
    group.append(s('rect', { x: 0, y: top, width, height: rowHeight, class: 'hit' }));
    group.append(
      s(
        'text',
        { x: narrow ? margin.left : 0, y: narrow ? top + 11 : cy + 4, class: 'op-name' },
        row.name,
        s('tspan', { class: 'op-count', dx: 6 }, `×${row.count}`)
      )
    );
    const from = x(row.p50);
    const to = Math.max(from + 2, x(row.p95));
    group.append(
      s('rect', { x: from, y: cy - 3, width: to - from, height: 6, rx: 3, class: 'range' }),
      s('circle', { cx: from, cy, r: 4.2, class: 'p50' })
    );
    if (!narrow)
      group.append(s('text', { x: width - 4, y: cy + 4, class: 'op-value', 'text-anchor': 'end' }, duration(row.p50)));
    group.addEventListener('pointermove', (event) =>
      tip(
        tipCard(
          row.name,
          [
            [options.cls, 'median', duration(row.p50)],
            ['muted', 'p95', duration(row.p95)],
            ['muted', 'slowest', duration(row.max)],
          ],
          `${row.count.toLocaleString('en-US')} timed calls`
        ),
        event.clientX,
        event.clientY
      )
    );
    group.addEventListener('pointerleave', () => tip(null));
    svg.append(group);
  });
  return svg;
}

export interface WaffleCell {
  id: string;
  state: string;
  describe: () => Node;
}

/** One square per item; hovering an item lights the same item in sibling waffles. */
export function waffle(cells: WaffleCell[], columns: number, onPick: (id: string) => void): HTMLElement {
  const grid = h('div', { class: 'waffle', style: `--columns:${columns}` });
  cells.forEach((cell) => {
    grid.append(h('i', { class: `cell s-${cell.state}`, 'data-id': cell.id }));
  });
  const byId = new Map(cells.map((cell) => [cell.id, cell]));
  grid.addEventListener('pointermove', (event) => {
    const target = (event.target as HTMLElement).closest<HTMLElement>('.cell');
    const panel = grid.closest('.panel') ?? grid;
    for (const lit of panel.querySelectorAll('.cell.lit')) lit.classList.remove('lit');
    if (!target?.dataset.id) return tip(null);
    for (const same of panel.querySelectorAll(`.cell[data-id="${CSS.escape(target.dataset.id)}"]`))
      same.classList.add('lit');
    tip(byId.get(target.dataset.id)!.describe(), event.clientX, event.clientY);
  });
  grid.addEventListener('pointerleave', () => {
    for (const lit of (grid.closest('.panel') ?? grid).querySelectorAll('.cell.lit')) lit.classList.remove('lit');
    tip(null);
  });
  grid.addEventListener('click', (event) => {
    const id = (event.target as HTMLElement).closest<HTMLElement>('.cell')?.dataset.id;
    if (id) onPick(id);
  });
  return grid;
}

/** A release-to-main dumbbell for one document. */
export function dumbbell(from: number | null, to: number, low: number, high: number): SVGSVGElement {
  const width = 84;
  const x = linear(low, high, 5, width - 5);
  const svg = s('svg', { width, height: 12, viewBox: `0 0 ${width} 12`, class: 'dumbbell', 'aria-hidden': 'true' });
  svg.append(s('line', { x1: 5, x2: width - 5, y1: 6, y2: 6, class: 'rail' }));
  if (from !== null) {
    svg.append(s('line', { x1: x(from), x2: x(to), y1: 6, y2: 6, class: to >= from ? 'gain' : 'loss' }));
    svg.append(s('circle', { cx: x(from), cy: 6, r: 3, class: 'from' }));
  }
  svg.append(s('circle', { cx: x(to), cy: 6, r: 3.4, class: 'to' }));
  return svg;
}
