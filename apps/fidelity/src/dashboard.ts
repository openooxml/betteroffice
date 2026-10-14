import {
  SWATCH,
  type SwarmSeries,
  beeswarm,
  cdf,
  dumbbell,
  duration,
  h,
  mount,
  spectrum,
  tipCard,
  waffle,
} from './charts';
import { FRAME_MS, type Latency, e2eUrl, parseLatestE2e, percentile, summarizeLatency } from './latency';
import { sameOrigin } from './origin';
import { parseLatest, reportUrl } from './report';
import {
  ENGINES,
  type Engine,
  FORMATS,
  type Format,
  type FormatSummary,
  type Summary,
  compareVersions,
  headlineText,
  msText,
  percentText,
  percentage,
  speedup,
  ssimText,
  summarize,
} from './summary';

const REPO = 'https://github.com/openooxml/betteroffice';
const METHOD = `${REPO}/blob/main/scripts/office-quality/README.md`;
const OFFICE: Record<Format, string> = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' };
const TITLES: Record<Format, string> = { docx: 'Word documents', pptx: 'Presentations', xlsx: 'Workbooks' };
const NOUNS: Record<Format, [string, string]> = {
  docx: ['document', 'documents'],
  pptx: ['deck', 'decks'],
  xlsx: ['workbook', 'workbooks'],
};

/** The page compares the latest release with LibreOffice; per-commit scores stay in the README. */
const OURS: Engine = 'published';
const SHOWN: Engine[] = [OURS, 'libreoffice'];

const params = new URL(window.location.href).searchParams;
const find = (id: string) => document.getElementById(id) as HTMLElement;
const plural = (count: number, [one, many]: [string, string]) => `${count.toLocaleString('en-US')} ${count === 1 ? one : many}`;
const thousands = (value: number) => value.toLocaleString('en-US');
const signed = (value: number, digits = 4) => `${value >= 0 ? '+' : '−'}${Math.abs(value).toFixed(digits)}`;
const short = (sha: string) => sha.slice(0, 8);
const change = (to: number, from: number) => Number(to.toFixed(4)) - Number(from.toFixed(4));
const columns = (count: number, aspect = 1.6) => Math.max(6, Math.ceil(Math.sqrt(count * aspect)));
const median = (values: number[]) => percentile([...values].sort((left, right) => left - right), 50);
const ratio = (part: number, whole: number) => (whole > 0 ? part / whole : 0);

const SECTION = 'pt-18 md:pt-24';
const PANELS = 'grid grid-cols-1 gap-3.5 md:grid-cols-2';
const STRIP = 'grid grid-cols-1 gap-px overflow-hidden rounded-panel border border-line bg-line md:grid-cols-2 lg:grid-cols-4';
const BOX = 'min-w-0 rounded-panel border border-line bg-panel p-4 md:p-5.5';
const WAFFLES = 'grid grid-cols-2 gap-3.5 md:gap-7';
const MONO = 'font-mono leading-[normal]';
const LINK = 'underline decoration-line-2 underline-offset-3 hover:decoration-fg';

/** Overrides may only name this origin, so a shared link cannot pass off someone else's numbers. */
const local = (name: string) => sameOrigin(params.get(name), window.location.href);

function compareHref(id: string): string {
  const url = new URL('/compare', window.location.origin);
  const report = local('report');
  if (report) url.searchParams.set('report', report);
  url.searchParams.set('doc', id);
  return url.pathname + url.search;
}

const open = (id: string) => window.location.assign(compareHref(id));

function when(iso: string | null): string {
  if (!iso) return '';
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return '';
  const days = Math.round((Date.now() - date.getTime()) / 86_400_000);
  const relative = new Intl.RelativeTimeFormat('en', { numeric: 'auto' }).format(-days, 'day');
  return `${date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })} · ${relative}`;
}

interface EngineView {
  label: string;
  sub: string;
  cls: string;
}

function views(summary: Summary, format: Format): Record<Engine, EngineView> {
  return {
    commit: { label: 'BetterOffice', sub: short(summary.commit), cls: 'c-commit' },
    published: { label: 'BetterOffice', sub: `v${summary.versions[format]}`, cls: 'c-ours' },
    libreoffice: { label: 'LibreOffice', sub: summary.libreoffice ?? '', cls: 'c-libreoffice' },
  };
}

function legendItem(cls: string, label: string, sub: string): HTMLElement {
  return h(
    'li',
    { class: 'flex items-center gap-[7px]' },
    h('i', { class: `${SWATCH} ${cls}` }),
    label,
    h('span', { class: `${MONO} text-[11px] text-dim` }, sub)
  );
}

function legend(entries: EngineView[]): HTMLElement {
  return h(
    'ul',
    { class: 'mt-4 flex flex-wrap gap-x-[22px] gap-y-1.5 text-[13px]' },
    ...entries.map((entry) => legendItem(entry.cls, entry.label, entry.sub))
  );
}

function count(value: number, digits: number, suffix = ''): HTMLElement {
  return h('span', {}, `${value.toLocaleString('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits })}${suffix}`);
}

interface KpiBar {
  cls: string;
  label: string;
  text: string;
  share: number;
}

/** A rounded bar filled to a 0–1 share; `null` draws the empty track. */
function track(share: number | null, size: string, fill = 'bg-series'): HTMLElement {
  return h(
    'i',
    { class: `block overflow-hidden rounded-full bg-track ${size}` },
    share === null
      ? null
      : h('i', {
          class: `block h-full w-(--w) rounded-[inherit] ${fill}`,
          style: `--w:${(Math.max(0, Math.min(1, share)) * 100).toFixed(2)}%`,
        })
  );
}

function kpi(
  label: string,
  value: HTMLElement,
  bars: KpiBar[],
  { note, compact = false }: { note?: string; compact?: boolean } = {}
) {
  return h(
    'article',
    { class: 'min-w-0 bg-solid px-6 pt-5.5 pb-5' },
    h('div', { class: 'text-[13px] text-ink' }, label),
    h(
      'div',
      {
        class: `mt-2.5 mb-4.5 leading-none font-semibold tracking-[-0.045em] whitespace-nowrap tabular-nums ${
          compact ? 'text-[34px]' : 'text-[38px] md:text-[44px]'
        }`,
      },
      value
    ),
    h(
      'div',
      { class: 'grid gap-1.25' },
      ...bars.map((bar) =>
        h(
          'div',
          { class: `grid grid-cols-[78px_minmax(0,1fr)_66px] items-center gap-2.5 ${MONO} text-[11px] text-dim ${bar.cls}` },
          h('span', { class: 'truncate' }, bar.label),
          track(bar.share, 'h-1.5'),
          h('b', { class: `text-right font-medium ${bar.cls === 'c-ours' ? 'text-fg' : 'text-ink'}` }, bar.text)
        )
      )
    ),
    note ? h('p', { class: 'mt-3.5 text-[12px] text-dim' }, note) : null
  );
}

function unit(text: string): HTMLElement {
  return h('span', { class: 'ml-0.5 text-[0.5em] font-medium tracking-normal text-dim' }, text);
}

function withUnit(node: HTMLElement, suffix: string): HTMLElement {
  return h('span', {}, node, unit(suffix));
}

function heroKpis(summary: Summary, latency: Latency | null): HTMLElement[] {
  const tiles: HTMLElement[] = [];
  const docx = summary.formats.docx;
  const pages = docx.fidelity[OURS];
  if (pages?.paged) {
    const engine = views(summary, 'docx');
    tiles.push(
      kpi(
        'Page counts match Word',
        withUnit(count(percentage(pages.exact, pages.paged, pages.exact === pages.paged ? 0 : 1)!, pages.exact === pages.paged ? 0 : 1), '%'),
        SHOWN.filter((key) => docx.fidelity[key]?.paged).map((key) => ({
          cls: engine[key].cls,
          label: key === 'libreoffice' ? 'LibreOffice' : engine[key].sub,
          text: `${docx.fidelity[key]!.exact}/${docx.fidelity[key]!.paged}`,
          share: docx.fidelity[key]!.exact / docx.fidelity[key]!.paged,
        })),
      )
    );
  }
  const calc = summary.formats.xlsx.calculation;
  if (calc?.engines[OURS].total) {
    const engine = views(summary, 'xlsx');
    const ours = calc.engines[OURS];
    tiles.push(
      kpi(
        'Formulas match Excel',
        withUnit(count(percentage(ours.correct, ours.total, 2)!, 2), '%'),
        SHOWN.map((key) => ({
          cls: engine[key].cls,
          label: key === 'libreoffice' ? 'LibreOffice' : engine[key].sub,
          text: headlineText(calc.engines[key].correct, calc.engines[key].total),
          share: calc.engines[key].correct / Math.max(1, calc.engines[key].total),
        })),
      )
    );
  }
  const render = summary.formats.pptx.render;
  const ratio = speedup(render?.engines[OURS].meanMs ?? null, render?.engines.libreoffice.meanMs ?? null);
  if (render && ratio) {
    const engine = views(summary, 'pptx');
    const slowest = Math.max(...SHOWN.map((key) => render.engines[key].meanMs ?? 0));
    tiles.push(
      kpi(
        ratio >= 1 ? 'Faster slides than LibreOffice' : 'Slower slides than LibreOffice',
        withUnit(count(ratio >= 1 ? ratio : 1 / ratio, 1), '×'),
        SHOWN.map((key) => ({
          cls: engine[key].cls,
          label: key === 'libreoffice' ? 'LibreOffice' : engine[key].sub,
          text: msText(render.engines[key].meanMs),
          share: (render.engines[key].meanMs ?? 0) / slowest,
        })),
        { note: 'Mean first-slide render, native CLI' }
      )
    );
  }
  if (latency?.ops) {
    tiles.push(
      kpi(
        'Calls within one 60 Hz frame',
        withUnit(count(percentage(latency.withinFrame, latency.ops, 1)!, 1), '%'),
        FORMATS.filter((format) => latency.formats[format]?.ops).map((format) => {
          const entry = latency.formats[format]!;
          return {
            cls: `f-${format}`,
            label: format.toUpperCase(),
            text: headlineText(entry.withinFrame, entry.ops, 1),
            share: entry.withinFrame / entry.ops,
          };
        }),
        { note: `End-to-end suite on main at ${short(latency.commit)}` }
      )
    );
  } else {
    const parsed = FORMATS.reduce((sum, format) => sum + (summary.formats[format].parsing?.[OURS].parsed ?? 0), 0);
    const total = FORMATS.reduce((sum, format) => sum + (summary.formats[format].parsing?.[OURS].total ?? 0), 0);
    if (total)
      tiles.push(
        kpi(
          'Documents opened',
          withUnit(count(percentage(parsed, total, 2)!, 2), '%'),
          FORMATS.filter((format) => summary.formats[format].parsing).map((format) => {
            const entry = summary.formats[format].parsing![OURS];
            return {
              cls: `f-${format}`,
              label: format.toUpperCase(),
              text: `${entry.parsed}/${entry.total}`,
              share: entry.parsed / Math.max(1, entry.total),
            };
          }),
        )
      );
  }
  return tiles;
}

interface Cell {
  value: number;
  text: string;
}

interface Metric {
  label: string;
  hint: string;
  better: 'high' | 'low';
  /** The value of a full bar; rows without one fill to their largest value. */
  full?: number;
  cells: Record<Engine, Cell | null>;
}

function metrics(summary: Summary, format: Format): Metric[] {
  const entry = summary.formats[format];
  const fidelity = (make: (engine: Engine) => Cell | null) =>
    Object.fromEntries(ENGINES.map((engine) => [engine, entry.fidelity[engine] ? make(engine) : null])) as Record<
      Engine,
      Cell | null
    >;
  const byEngine = (make: (engine: Engine) => Cell | null) =>
    Object.fromEntries(ENGINES.map((engine) => [engine, make(engine)])) as Record<Engine, Cell | null>;
  const rows: Metric[] = [];
  if (format === 'docx')
    rows.push(
      {
        label: 'Exact page counts',
        hint: 'documents with Word’s page count',
        better: 'high',
        full: 1,
        cells: fidelity((engine) => {
          const value = entry.fidelity[engine]!;
          return value.paged ? { value: value.exact / value.paged, text: `${value.exact}/${value.paged}` } : null;
        }),
      },
      {
        label: 'Absolute page error',
        hint: 'pages off, summed over documents',
        better: 'low',
        cells: fidelity((engine) => {
          const value = entry.fidelity[engine]!;
          return value.paged ? { value: value.pageError, text: thousands(value.pageError) } : null;
        }),
      }
    );
  rows.push(
    {
      label: 'Visual similarity',
      hint: `mean SSIM against ${OFFICE[format]}`,
      better: 'high',
      full: 1,
      cells: fidelity((engine) => {
        const mean = entry.fidelity[engine]!.mean;
        return mean === null ? null : { value: mean, text: ssimText(mean) };
      }),
    },
    {
      label: 'Scored / total',
      hint: `${NOUNS[format][1]} with a score`,
      better: 'high',
      full: 1,
      cells: fidelity((engine) => {
        const value = entry.fidelity[engine]!;
        return { value: value.scored / Math.max(1, value.total), text: `${value.scored}/${value.total}` };
      }),
    }
  );
  if (entry.render)
    rows.push({
      label: format === 'pptx' ? 'Slide render time' : 'Page render time',
      hint: `native CLI mean, first ${format === 'pptx' ? 'slide' : 'page'}, ${entry.render.common} common`,
      better: 'low',
      cells: byEngine((engine) => {
        const ms = entry.render!.engines[engine].meanMs;
        return ms === null ? null : { value: ms, text: msText(ms) };
      }),
    });
  if (entry.calculation) {
    const calc = entry.calculation;
    rows.push(
      {
        label: 'Recalc accuracy',
        hint: 'formula results matching Excel',
        better: 'high',
        full: 1,
        cells: byEngine((engine) => {
          const { correct, total } = calc.engines[engine];
          return total ? { value: correct / total, text: percentText(correct, total) } : null;
        }),
      },
      {
        label: 'Recalc time',
        hint: `native CLI mean, ${calc.common} common workbooks`,
        better: 'low',
        cells: byEngine((engine) => {
          const ms = calc.engines[engine].meanMs;
          return ms === null ? null : { value: ms, text: msText(ms) };
        }),
      }
    );
  }
  if (entry.parsing)
    rows.push({
      label: 'Parse success',
      hint: `${NOUNS[format][1]} opened`,
      better: 'high',
      full: 1,
      cells: byEngine((engine) => {
        const { parsed, total } = entry.parsing![engine];
        return total ? { value: parsed / total, text: percentText(parsed, total) } : null;
      }),
    });
  return rows;
}

function formatLabel(format: Format, color = 'text-dim'): HTMLElement {
  return h('span', { class: `${MONO} text-[11px] font-medium tracking-[0.06em] ${color}` }, format.toUpperCase());
}

/** One row per metric, the engines' bars stacked on one scale so they compare at a glance. */
function scoreboard(summary: Summary): HTMLElement {
  const board = h('div', { class: 'overflow-clip rounded-panel border border-line bg-panel' });
  for (const format of FORMATS) {
    const entry = summary.formats[format];
    if (!entry.documents) continue;
    const engine = views(summary, format);
    board.append(
      h(
        'a',
        {
          class:
            'flex flex-wrap items-baseline gap-x-2.5 gap-y-1 border-line px-5 pt-5.5 pb-2.5 text-[14px] font-semibold not-first:border-t',
          href: `#${format}`,
        },
        formatLabel(format),
        h('span', {}, TITLES[format]),
        h('small', { class: 'text-[12.5px] font-normal text-dim' }, `${plural(entry.documents, NOUNS[format])} · release v${entry.version}`)
      )
    );
    for (const metric of metrics(summary, format)) {
      const present = SHOWN.map((key) => metric.cells[key]?.value).filter((value): value is number => value !== undefined);
      const best = metric.better === 'high' ? Math.max(...present) : Math.min(...present);
      const tie = present.every((value) => value === best);
      const full = metric.full ?? Math.max(...present);
      const share = (value: number) => (full > 0 ? value / full : 0);
      board.append(
        h(
          'div',
          {
            class:
              'grid grid-cols-1 items-center border-t border-line hover:bg-hover md:grid-cols-[minmax(220px,1fr)_minmax(0,2fr)]',
          },
          h(
            'div',
            { class: 'grid gap-px px-3.5 pt-3.5 text-[13.5px] md:px-5 md:py-3.5' },
            metric.label,
            h('small', { class: 'text-[12px] text-dim' }, metric.hint)
          ),
          h(
            'div',
            { class: 'grid gap-[7px] px-3.5 pt-2.5 pb-3.5 md:px-5 md:py-3.5' },
            ...SHOWN.map((key) => {
              const cell = metric.cells[key];
              const view = engine[key];
              const top = !!cell && present.length > 1 && !tie && cell.value === best;
              const tone = !cell ? 'text-faint' : top ? 'text-fg' : 'text-ink';
              return h(
                'div',
                {
                  class: `grid grid-cols-[84px_minmax(0,1fr)_84px] items-center gap-2.5 md:grid-cols-[96px_minmax(0,1fr)_104px] md:gap-3.5 ${view.cls}`,
                },
                h('span', { class: `text-[12.5px] ${top ? 'text-fg' : 'text-ink'}` }, view.label),
                track(cell ? share(cell.value) : null, 'h-2', top ? 'bg-series' : 'bg-series opacity-55'),
                h(
                  'b',
                  { class: `text-right ${MONO} text-[13.5px] font-medium tabular-nums ${tone}` },
                  top ? h('i', { class: 'mr-2 inline-block size-1.5 rounded-full bg-series align-[2px]' }) : null,
                  cell ? cell.text : '—'
                )
              );
            })
          )
        )
      );
    }
  }
  return board;
}

function labelled<T extends Element>(node: T, label: string): T {
  node.setAttribute('role', 'img');
  node.setAttribute('aria-label', label);
  return node;
}

function head(title: string, meta?: string, ...extra: HTMLElement[]): HTMLElement {
  return h(
    'header',
    { class: 'mb-6' },
    h('h2', { class: 'text-[24px] leading-[1.15] font-semibold tracking-[-0.03em] md:text-[28px]' }, title),
    meta ? h('p', { class: 'mt-2 max-w-[760px] text-[14px] text-ink' }, meta) : null,
    ...extra
  );
}

function caption(title: string, text: string, aside: HTMLElement | null): HTMLElement {
  return h(
    'figcaption',
    { class: 'mb-4.5 flex flex-col items-start justify-between gap-2.5 md:flex-row md:gap-6' },
    h(
      'div',
      {},
      h('h3', { class: 'text-[15px] font-semibold tracking-[-0.01em]' }, title),
      h('p', { class: 'mt-1 max-w-[640px] text-[13px] text-dim' }, text)
    ),
    aside
  );
}

function panel(title: string, text: string, stat: HTMLElement | null, body: HTMLElement, wide = false) {
  return h('figure', { class: `${BOX}${wide ? ' col-span-full' : ''}` }, caption(title, text, stat), body);
}

function empty(message: string, href: string, label: string): HTMLElement {
  return h(
    'div',
    { class: `grid justify-items-start gap-2 text-ink ${BOX}` },
    h('p', {}, message),
    h('a', { class: 'text-[13px] underline', href }, label)
  );
}

function stat(value: string, note: string, trend: '' | 'up' | 'down' = ''): HTMLElement {
  const tone = trend === 'up' ? 'text-bo' : trend === 'down' ? 'text-bad' : 'text-dim';
  return h(
    'div',
    { class: 'flex-none md:text-right' },
    h('b', { class: 'block text-[28px] leading-[1.1] font-semibold tracking-[-0.04em] tabular-nums' }, value),
    h('span', { class: `text-[12px] ${tone}` }, note)
  );
}

function waffleBlock(view: EngineView, grid: HTMLElement, value: string, label: string, sub: string): HTMLElement {
  return h(
    'div',
    { class: view.cls },
    h(
      'div',
      {
        class:
          'mb-2.5 flex flex-col items-start gap-0.5 overflow-hidden text-[12.5px] whitespace-nowrap md:flex-row md:items-center md:gap-[7px]',
      },
      h('i', { class: `${SWATCH} ${view.cls} max-md:hidden` }),
      view.label,
      h('span', { class: `${MONO} text-[10.5px] text-dim` }, view.sub)
    ),
    grid,
    h(
      'div',
      { class: 'mt-3 flex items-baseline gap-[7px]' },
      h('b', { class: 'text-[21px] font-semibold tracking-[-0.03em] tabular-nums' }, value),
      h('span', { class: `${MONO} text-[11px] text-dim` }, label)
    ),
    h('div', { class: `mt-0.5 truncate ${MONO} text-[11px] text-dim` }, sub)
  );
}

function stateKey(entries: [state: string, label: string][]): HTMLElement {
  return h(
    'ul',
    { class: `mt-4.5 flex flex-wrap gap-x-4 gap-y-2 ${MONO} text-[11px] text-dim` },
    ...entries.map(([state, label]) =>
      h('li', { class: 'flex items-center gap-1.5' }, h('i', { class: `cell s-${state} w-2.5` }), label)
    )
  );
}

function chart(build: (width: number) => Element, label: string): HTMLElement {
  const container = h('div', { class: 'min-w-0' });
  queueMicrotask(() =>
    mount(container, (width) => {
      const node = build(width);
      node.setAttribute('aria-label', label);
      return node;
    })
  );
  return container;
}

function similarity(summary: Summary, entry: FormatSummary): HTMLElement {
  const engine = views(summary, entry.format);
  const series: SwarmSeries[] = SHOWN.filter((key) => entry.fidelity[key]).map((key) => ({
    key,
    cls: engine[key].cls,
    label: engine[key].label,
    sub: engine[key].sub,
  }));
  const byId = new Map(entry.rows.map((row) => [row.id, row]));
  const describe = (id: string) => {
    const row = byId.get(id)!;
    const delta =
      row.ssim[OURS] !== null && row.ssim.libreoffice !== null
        ? `${signed(change(row.ssim[OURS], row.ssim.libreoffice))} vs LibreOffice · `
        : '';
    return tipCard(
      id,
      series.map((item) => [item.cls, `${item.label} ${item.sub}`, row.ssim[item.key as Engine] === null ? 'no score' : ssimText(row.ssim[item.key as Engine])]),
      `${delta}click to compare pages`
    );
  };
  const mean = entry.fidelity[OURS]?.mean ?? null;
  const other = entry.fidelity.libreoffice?.mean ?? null;
  const caption =
    entry.format === 'xlsx'
      ? 'SSIM over the recorded print ranges against Excel, one dot per workbook. Click a dot to open its pages.'
      : `Page-penalized SSIM against ${OFFICE[entry.format]}, one dot per ${NOUNS[entry.format][0]}. Click a dot to open its pages.`;
  return panel(
    'Visual similarity',
    caption,
    mean === null
      ? null
      : stat(
          ssimText(mean),
          other === null ? `mean of v${entry.version}` : `${signed(change(mean, other))} vs LibreOffice`,
          other === null || change(mean, other) >= 0 ? 'up' : 'down'
        ),
    chart(
      (width) =>
        beeswarm(
          entry.rows.map((row) => ({ id: row.id, values: row.ssim })),
          series,
          width,
          { format: ssimText, onPick: open, describe }
        ),
      `Similarity to ${OFFICE[entry.format]} per ${NOUNS[entry.format][0]}. ${series
        .map((item) => `${item.label} ${item.sub}: mean ${ssimText(entry.fidelity[item.key as Engine]?.mean)}`)
        .join('. ')}.`
    ),
    true
  );
}

function pagination(summary: Summary, entry: FormatSummary): HTMLElement {
  const engine = views(summary, entry.format);
  const rows = [...entry.rows].sort(
    (left, right) => (left.referencePages ?? 0) - (right.referencePages ?? 0) || left.id.localeCompare(right.id)
  );
  const blocks = h('div', { class: WAFFLES });
  for (const key of SHOWN) {
    const fidelity = entry.fidelity[key];
    if (!fidelity?.paged) continue;
    const cells = rows.map((row) => {
      const pages = row.pages[key];
      const offset = pages === null || row.referencePages === null ? null : pages - row.referencePages;
      const state = offset === null ? 'none' : offset === 0 ? 'exact' : Math.abs(offset) === 1 ? 'near' : 'far';
      return {
        id: row.id,
        state,
        describe: () =>
          tipCard(
            row.id,
            [
              ['muted', OFFICE[entry.format], row.referencePages === null ? '—' : plural(row.referencePages, ['page', 'pages'])],
              [
                engine[key].cls,
                `${engine[key].label} ${engine[key].sub}`,
                pages === null ? 'no score' : `${plural(pages, ['page', 'pages'])}${offset ? ` (${offset > 0 ? '+' : '−'}${Math.abs(offset)})` : ''}`,
              ],
            ],
            'click to compare pages'
          ),
      };
    });
    blocks.append(
      waffleBlock(
        engine[key],
        labelled(
          waffle(cells, columns(rows.length), open),
          `${engine[key].label} ${engine[key].sub}: ${fidelity.exact} of ${fidelity.paged} page counts exact`
        ),
        `${fidelity.exact}/${fidelity.paged}`,
        'exact',
        `${plural(fidelity.pageError, ['page', 'pages'])} off`
      )
    );
  }
  const ours = entry.fidelity[OURS]!;
  return panel(
    'Page count',
    'Rendered page count against Word, one square per document, shortest first.',
    stat(`${ours.exact}/${ours.paged}`, `exact in v${entry.version}`, ours.exact >= (entry.fidelity.libreoffice?.exact ?? 0) ? 'up' : 'down'),
    h(
      'div',
      {},
      blocks,
      stateKey([
        ['exact', 'exact'],
        ['near', 'one page off'],
        ['far', 'two or more off'],
        ['none', 'no score'],
      ])
    )
  );
}

interface Line {
  cls: string;
  label: string;
  values: number[];
}

function lines(series: Line[], options: { noun: string; frame?: number; height?: number }): HTMLElement {
  return h(
    'div',
    {},
    h(
      'ul',
      { class: 'mb-3.5 flex flex-wrap gap-x-[18px] gap-y-1.5 text-[13px]' },
      ...series.map((entry) => legendItem(entry.cls, entry.label, `median ${duration(median(entry.values))}`))
    ),
    chart(
      (width) => cdf(series, width, options),
      `Share of ${options.noun} finished within a given time. ${series
        .map((entry) => `${entry.label}: median ${duration(median(entry.values))}`)
        .join('. ')}.`
    )
  );
}

function renderTimes(summary: Summary, entry: FormatSummary): HTMLElement | null {
  const render = entry.render;
  if (!render?.common) return null;
  const engine = views(summary, entry.format);
  const common = entry.rows.filter((row) => ENGINES.every((key) => row.renderMs[key] !== null));
  const ratio = speedup(render.engines[OURS].meanMs, render.engines.libreoffice.meanMs);
  const unit = entry.format === 'pptx' ? 'slide' : 'page';
  return panel(
    `${unit === 'slide' ? 'Slide' : 'Page'} render time`,
    `Native CLI, first ${unit} at 96 DPI, over the ${render.common} ${NOUNS[entry.format][1]} every engine rendered.`,
    ratio ? stat(`${(ratio >= 1 ? ratio : 1 / ratio).toFixed(1)}×`, `${ratio >= 1 ? 'faster' : 'slower'} than LibreOffice`, ratio >= 1 ? 'up' : 'down') : null,
    lines(
      SHOWN.map((key) => ({
        cls: engine[key].cls,
        label: `${engine[key].label} ${engine[key].sub}`,
        values: common.map((row) => row.renderMs[key]!),
      })),
      { noun: NOUNS[entry.format][1] }
    )
  );
}

function formulas(summary: Summary, entry: FormatSummary): HTMLElement | null {
  const calc = entry.calculation;
  if (!calc?.workbooks) return null;
  const engine = views(summary, entry.format);
  const rows = entry.rows
    .filter((row) => ENGINES.every((key) => row.recalc[key] !== null))
    .sort((left, right) => right.recalc[OURS]!.total - left.recalc[OURS]!.total || left.id.localeCompare(right.id));
  const bars = h(
    'div',
    { class: 'mb-6.5 grid gap-3' },
    ...SHOWN.map((key) => {
      const { correct, total } = calc.engines[key];
      return h(
        'div',
        {
          class: `grid grid-cols-[minmax(0,1fr)_auto] items-center gap-x-3 gap-y-1.5 text-[13px] md:grid-cols-[190px_minmax(0,1fr)_76px_128px] md:gap-4 ${engine[key].cls}`,
        },
        h('span', {}, engine[key].label, h('small', { class: `ml-[7px] ${MONO} text-[10.5px] text-dim` }, engine[key].sub)),
        track(correct / Math.max(1, total), 'order-3 col-span-full h-2.5 md:order-none md:col-auto', 'bg-linear-to-r from-series/35 to-series'),
        h('b', { class: `text-right ${MONO} text-[14px] font-semibold` }, percentText(correct, total)),
        h(
          'small',
          { class: `ml-[7px] hidden text-right ${MONO} text-[10.5px] text-dim md:block` },
          `${thousands(correct)} / ${thousands(total)}`
        )
      );
    })
  );
  const blocks = h('div', { class: WAFFLES });
  for (const key of SHOWN) {
    const cells = rows.map((row) => {
      const result = row.recalc[key]!;
      const share = result.total ? result.correct / result.total : 0;
      const state = !result.ok ? 'none' : share === 1 ? 'exact' : share >= 0.99 ? 'high' : share >= 0.9 ? 'near' : 'far';
      return {
        id: row.id,
        state,
        describe: () =>
          tipCard(
            row.id,
            [
              [
                engine[key].cls,
                `${engine[key].label} ${engine[key].sub}`,
                result.ok ? `${thousands(result.correct)} / ${thousands(result.total)}` : 'failed',
              ],
            ],
            result.ok ? `${percentText(result.correct, result.total)} of formula results match Excel` : 'recalculation failed'
          ),
      };
    });
    blocks.append(
      waffleBlock(
        engine[key],
        labelled(
          waffle(cells, columns(rows.length, 3.2), open),
          `${engine[key].label} ${engine[key].sub}: ${calc.engines[key].perfect} of ${rows.length} workbooks recalculate perfectly`
        ),
        String(calc.engines[key].perfect),
        'perfect',
        `of ${plural(rows.length, NOUNS.xlsx)}`
      )
    );
  }
  const ours = calc.engines[OURS];
  return panel(
    'Formula accuracy',
    'Result cells matching Excel after a full recalculation, one square per workbook, largest first.',
    stat(
      headlineText(ours.correct, ours.total),
      `${thousands(ours.correct)} of ${thousands(ours.total)} match`,
      ours.correct >= calc.engines.libreoffice.correct ? 'up' : 'down'
    ),
    h(
      'div',
      {},
      bars,
      blocks,
      stateKey([
        ['exact', 'all match'],
        ['high', '≥ 99%'],
        ['near', '≥ 90%'],
        ['far', 'below 90%'],
        ['none', 'failed'],
      ])
    ),
    true
  );
}

function recalcTimes(summary: Summary, entry: FormatSummary): HTMLElement | null {
  const calc = entry.calculation;
  if (!calc?.common) return null;
  const engine = views(summary, entry.format);
  const common = entry.rows.filter((row) =>
    ENGINES.every((key) => row.recalc[key]?.ok && row.recalc[key]!.correct === row.recalc[key]!.total && row.recalc[key]!.ms !== null)
  );
  const ratio = speedup(calc.engines[OURS].meanMs, calc.engines.libreoffice.meanMs);
  return panel(
    'Recalculation time',
    `Native CLI, over the ${calc.common} workbooks every engine recalculates perfectly.`,
    ratio ? stat(`${(ratio >= 1 ? ratio : 1 / ratio).toFixed(1)}×`, `${ratio >= 1 ? 'faster' : 'slower'} than LibreOffice`, ratio >= 1 ? 'up' : 'down') : null,
    lines(
      SHOWN.map((key) => ({
        cls: engine[key].cls,
        label: `${engine[key].label} ${engine[key].sub}`,
        values: common.map((row) => row.recalc[key]!.ms!),
      })),
      { noun: 'workbooks' }
    )
  );
}

function movers(entry: FormatSummary, wide: boolean): HTMLElement {
  const version = `v${entry.version}`;
  const scored = entry.rows.filter((row) => row.ssim[OURS] !== null);
  const values = scored.flatMap((row) => [row.ssim[OURS]!, row.ssim.libreoffice ?? row.ssim[OURS]!]);
  const low = Math.min(...values);
  const high = Math.max(...values);
  const lead = (row: (typeof scored)[number]) => row.ssim[OURS]! - (row.ssim.libreoffice ?? row.ssim[OURS]!);
  const leads = scored
    .filter((row) => row.ssim.libreoffice !== null && lead(row) > 0)
    .sort((left, right) => lead(right) - lead(left))
    .slice(0, 5);
  const lowest = [...scored].sort((left, right) => left.ssim[OURS]! - right.ssim[OURS]!).slice(0, 5);
  const failed = entry.rows.filter((row) => row.ssim[OURS] === null);
  const item = (row: (typeof scored)[number]) => {
    const delta = row.ssim.libreoffice === null ? null : change(row.ssim[OURS]!, row.ssim.libreoffice);
    const bell = dumbbell(row.ssim.libreoffice, row.ssim[OURS]!, low, high);
    bell.classList.add('max-md:hidden');
    const tone = delta === null ? 'text-dim' : delta >= 0 ? 'text-bo' : 'text-bad';
    return h(
      'li',
      {},
      h(
        'a',
        {
          href: compareHref(row.id),
          class:
            '-mx-2.5 grid grid-cols-[minmax(0,1fr)_58px_76px] items-center gap-3 rounded-[9px] px-2.5 py-2 text-[13px] hover:bg-hover md:grid-cols-[minmax(0,1fr)_84px_58px_76px]',
        },
        h('span', { class: `truncate ${MONO} text-[12.5px]` }, row.id),
        bell,
        h('b', { class: `text-right ${MONO} text-[12.5px] font-medium` }, ssimText(row.ssim[OURS])),
        h('span', { class: `text-right ${MONO} text-[11.5px] ${tone}` }, delta === null ? 'no LO score' : signed(delta))
      )
    );
  };
  const column = (title: string, rows: typeof scored) =>
    h('div', {}, h('h4', { class: 'mb-2 text-[12.5px] font-medium text-dim' }, title), h('ol', {}, ...rows.map(item)));
  const note = failed.length ? `${plural(failed.length, NOUNS[entry.format])} did not score in ${version}.` : '';
  return panel(
    'Notable documents',
    'Largest leads over LibreOffice and lowest scores. Each row opens the page viewer.',
    null,
    h(
      'div',
      {},
      h(
        'div',
        { class: wide ? 'grid grid-cols-1 gap-5 md:grid-cols-2 md:gap-8' : 'grid grid-cols-1 gap-[22px]' },
        column('Largest leads over LibreOffice', leads),
        column(`Lowest in ${version}`, lowest)
      ),
      note
        ? h(
            'p',
            { class: 'mt-4 text-[12.5px] text-dim' },
            note,
            ...failed
              .slice(0, 6)
              .flatMap((row) => [' ', h('a', { class: 'font-mono text-ink underline', href: compareHref(row.id) }, row.id)])
          )
        : null
    ),
    wide
  );
}

function formatSection(summary: Summary, format: Format): HTMLElement {
  const entry = summary.formats[format];
  const engine = views(summary, format);
  const reference = entry.reference;
  const versions = reference?.versions ?? [];
  const office = reference
    ? `${reference.engine} ${versions.length > 1 ? `${versions[0]}–${versions.at(-1)}` : (versions[0] ?? '')}`.trim()
    : `Microsoft ${OFFICE[format]}`;
  const panels = h('div', { class: PANELS }, similarity(summary, entry));
  if (format === 'docx') panels.append(pagination(summary, entry));
  const render = renderTimes(summary, entry);
  if (render) panels.append(render);
  const accuracy = formulas(summary, entry);
  if (accuracy) panels.append(accuracy);
  const recalc = recalcTimes(summary, entry);
  if (recalc) panels.append(recalc);
  const halves = [...panels.children].filter((node) => !node.classList.contains('col-span-full')).length;
  panels.append(movers(entry, halves % 2 === 0));
  return h(
    'section',
    { class: SECTION, id: format },
    head(
      TITLES[format],
      `${format.toUpperCase()} · ${plural(entry.documents, NOUNS[format])} · reference ${office}`,
      legend(SHOWN.filter((key) => entry.fidelity[key] || entry.render || entry.calculation).map((key) => engine[key]))
    ),
    panels
  );
}

function latencySection(latency: Latency | null, published: boolean): HTMLElement {
  const header = head(
    'Editing latency',
    latency?.ops
      ? `SDK calls timed by the end-to-end suite on main at ${short(latency.commit)}, in WASM and Python.`
      : 'SDK calls timed by the end-to-end suite on main, in WASM and Python.'
  );
  if (!latency?.ops)
    return h(
      'section',
      { class: SECTION, id: 'latency' },
      header,
      empty(
        'Editing latency appears here once a green end-to-end run on main publishes its timings.',
        `${REPO}/actions/workflows/e2e.yml`,
        'End-to-end workflow ↗'
      )
    );
  const present = FORMATS.filter((format) => latency.formats[format]);
  const tiles = h(
    'div',
    { class: `mb-4 ${STRIP}` },
    kpi(
      'Scenarios passed',
      h('span', {}, count(latency.passed, 0), unit(` / ${latency.cases}`)),
      present.map((format) => ({
        cls: `f-${format}`,
        label: format.toUpperCase(),
        text: `${latency.formats[format]!.passed}/${latency.formats[format]!.cases.length}`,
        share: latency.formats[format]!.passed / Math.max(1, latency.formats[format]!.cases.length),
      })),
      {
        note: `${plural(latency.multiEditor, ['case', 'cases'])} with several editors, ${latency.crossSdk} across Python and WASM`,
        compact: true,
      }
    ),
    kpi(
      'Timed SDK calls',
      count(latency.ops, 0),
      present.map((format) => ({
        cls: `f-${format}`,
        label: format.toUpperCase(),
        text: thousands(latency.formats[format]!.ops),
        share: ratio(latency.formats[format]!.ops, Math.max(...present.map((key) => latency.formats[key]!.ops))),
      })),
      { compact: true }
    ),
    kpi(
      'Median call',
      h('span', {}, count(latency.p50, latency.p50 < 10 ? 2 : 0), unit(' ms')),
      present.map((format) => ({
        cls: `f-${format}`,
        label: format.toUpperCase(),
        text: duration(latency.formats[format]!.p50),
        share: ratio(latency.formats[format]!.p50, Math.max(...present.map((key) => latency.formats[key]!.p50))),
      })),
      { note: `Across all ${thousands(latency.ops)} calls`, compact: true }
    ),
    kpi(
      '95th percentile call',
      h('span', {}, count(latency.p95, latency.p95 < 10 ? 1 : 0), unit(' ms')),
      present.map((format) => ({
        cls: `f-${format}`,
        label: format.toUpperCase(),
        text: duration(latency.formats[format]!.p95),
        share: ratio(latency.formats[format]!.p95, Math.max(...present.map((key) => latency.formats[key]!.p95))),
      })),
      { note: `Across all ${thousands(latency.ops)} calls`, compact: true }
    )
  );

  let selected: Format = present[0]!;
  let release: (() => void) | null = null;
  const switcher = h('div', {
    class: 'inline-flex flex-none rounded-[9px] border border-line bg-hover p-0.75',
    role: 'tablist',
  });
  const plot = h('div', { class: 'min-w-0' });
  const draw = () => {
    for (const button of switcher.querySelectorAll('button'))
      button.setAttribute('aria-selected', String(button.dataset.format === selected));
    const groups = [...latency.formats[selected]!.groups]
      .filter((group) => group.count >= 5)
      .slice(0, 14)
      .sort((left, right) => left.p50 - right.p50);
    const container = h('div', {});
    plot.replaceChildren(container);
    release?.();
    release = mount(container, (width) => {
      const node = spectrum(groups, width, { cls: `f-${selected}`, frame: FRAME_MS });
      node.setAttribute(
        'aria-label',
        `${selected.toUpperCase()} latency by operation. ${groups.map((group) => `${group.name}: median ${duration(group.p50)}`).join(', ')}.`
      );
      return node;
    });
  };
  for (const format of present) {
    const button = h(
      'button',
      {
        type: 'button',
        role: 'tab',
        'data-format': format,
        class: `cursor-pointer rounded-md px-3 py-1.25 font-sans text-[12px] leading-[normal] font-medium text-dim hover:text-fg aria-selected:bg-solid aria-selected:text-fg aria-selected:shadow-[0_0_0_1px_var(--line-2)] f-${format}`,
      },
      format.toUpperCase()
    );
    button.addEventListener('click', () => {
      selected = format;
      draw();
    });
    switcher.append(button);
  }
  queueMicrotask(draw);

  const operations = h(
    'figure',
    { class: `${BOX} col-span-full` },
    caption('Latency by operation', 'Median call and 95th percentile per operation, grouped by verb.', switcher),
    plot
  );

  const all = panel(
    'Call latency',
    'Share of SDK calls finished within a given time.',
    stat(headlineText(latency.withinFrame, latency.ops, 1), 'inside one 60 Hz frame', 'up'),
    lines(
      present.map((format) => ({ cls: `f-${format}`, label: format.toUpperCase(), values: latency.formats[format]!.values })),
      { frame: FRAME_MS, noun: 'calls', height: 260 }
    )
  );

  const cases = h('div', { class: 'grid gap-3.5' });
  for (const format of present) {
    const entry = latency.formats[format]!;
    const byKey = new Map(entry.cases.map((item) => [`${item.scenario}/${item.sample}`, item]));
    cases.append(
      h(
        'div',
        {
          class:
            'grid grid-cols-[44px_minmax(0,1fr)_52px] items-center gap-3.5 [&_.waffle]:grid-cols-[repeat(auto-fill,16px)]',
        },
        formatLabel(format, `text-series f-${format}`),
        waffle(
          entry.cases.map((item) => ({
            id: `${format}:${item.scenario}/${item.sample}`,
            state: item.status === 'passed' ? 'exact' : 'far',
            describe: () =>
              tipCard(
                item.scenario,
                [
                  ['muted', 'document', item.sample],
                  ['muted', 'actors', item.participants.join(', ') || '—'],
                  [`f-${format}`, 'calls', `${item.opCount} in ${duration(byKey.get(`${item.scenario}/${item.sample}`)!.totalMs)}`],
                ],
                item.description
              ),
          })),
          Math.min(entry.cases.length, 13),
          () => undefined
        ),
        h('b', { class: `text-right ${MONO} text-[12px] font-medium text-ink` }, `${entry.passed}/${entry.cases.length}`)
      )
    );
  }
  const machine = latency.environment;
  const scenarios = panel(
    'Scenarios',
    'One square per scenario and document. Hover for details.',
    stat(`${latency.passed}/${latency.cases}`, 'passed', latency.passed === latency.cases ? 'up' : 'down'),
    h(
      'div',
      {},
      cases,
      h(
        'p',
        { class: 'mt-5 text-[12px] leading-[1.6] text-dim' },
        `Recorded ${when(latency.recordedAt)} at `,
        h('a', { class: 'font-mono text-ink underline', href: `${REPO}/commit/${latency.commit}` }, short(latency.commit)),
        machine ? ` on ${published ? 'a GitHub-hosted runner, ' : ''}${machine.cpu}, ${machine.cpus} vCPU. ` : '. ',
        'These are observations from one run, not controlled benchmark trials.'
      )
    )
  );

  return h('section', { class: SECTION, id: 'latency' }, header, tiles, h('div', { class: PANELS }, operations, all, scenarios));
}

function methodSection(
  summary: Summary | null,
  reportHref: string | null,
  latency: Latency | null,
  published: boolean
): HTMLElement {
  const lo = summary?.libreoffice ?? 'the pinned LibreOffice release';
  const versions = summary
    ? [...new Set(FORMATS.flatMap((format) => summary.formats[format].reference?.versions ?? []))].sort(compareVersions)
    : [];
  const office = versions.length ? ` ${versions[0]}${versions.length > 1 ? `–${versions.at(-1)}` : ''}` : '';
  const card = (title: string, body: string) =>
    h(
      'article',
      {},
      h('h3', { class: 'mb-1.5 text-[13.5px] font-semibold' }, title),
      h('p', { class: 'text-[13px] leading-[1.6] text-dim' }, body)
    );
  const links: [string, string][] = [
    ['Benchmark methodology', METHOD],
    ['Benchmarks workflow', `${REPO}/actions/workflows/visual-fidelity.yml`],
    ['End-to-end suite', `${REPO}/blob/main/e2e/README.md`],
  ];
  if (reportHref) links.push(['Raw report.json', reportHref]);
  if (latency) links.push(['Raw e2e results', e2eUrl(latency.commit, FORMATS.find((format) => latency.formats[format])!)]);
  return h(
    'section',
    { class: SECTION, id: 'method' },
    head('Method'),
    h(
      'div',
      { class: 'grid grid-cols-1 gap-7 border-t border-line pt-6 md:grid-cols-3 lg:grid-cols-5' },
      card('Ground truth', `Microsoft Word, PowerPoint and Excel${office} on macOS exported every reference page to PDF, rasterized at 150 DPI. The corpus and its references are public.`),
      card('Visual similarity', 'Mean page-penalized grayscale SSIM at 150 DPI, without resampling or alignment, of the BetterOffice browser renderer. Missing or extra pages lower the score. Failed renders get no score, so read coverage next to the mean.'),
      card('Speed', `Native command-line builds time the whole job in five fresh processes after a warmup: start, import, fonts, layout, rasterize and write. They measure the Rust rasterizer at 96 DPI, not the browser renderer scored for similarity. LibreOffice ${lo} runs its prebuilt CLI with an isolated profile. Means cover the documents every engine handled.`),
      card('Formulas', 'Caches are cleared and each workbook is recalculated from scratch. Numbers match within 1e-9, or 1e-12 relative; text, booleans and errors must match exactly. Missing results count as wrong.'),
      card('Latency', `The end-to-end suite times every SDK call from invocation to result ${published ? 'on a GitHub-hosted runner' : 'on the machine that recorded it'}. Repeated calls see different document state, so these are observations, not benchmark trials.`)
    ),
    h(
      'ul',
      { class: 'mt-7 flex flex-wrap gap-x-6 gap-y-2.5' },
      ...links.map(([label, href]) => h('li', {}, h('a', { class: 'text-[13px] text-ink hover:text-fg', href }, `${label} ↗`)))
    )
  );
}

async function json(url: string): Promise<unknown> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${response.status} ${url}`);
  return response.json();
}

async function benchmarks(): Promise<{ summary: Summary; publishedAt: string | null; href: string }> {
  const override = local('report');
  if (override) return { summary: summarize(await json(override)), publishedAt: null, href: override };
  const manifest = parseLatest(await json('/renders/latest.json'));
  const href = reportUrl(manifest.sha);
  return { summary: summarize(await json(href)), publishedAt: manifest.publishedAt, href };
}

async function endToEnd(): Promise<Latency | null> {
  try {
    const override = local('e2e');
    if (override)
      return summarizeLatency(await Promise.all(FORMATS.map((format) => json(`${override.replace(/\/$/, '')}/${format}.json`))));
    const latest = parseLatestE2e(await json('/e2e/latest.json'));
    return summarizeLatency(await Promise.all(latest.formats.map((format) => json(e2eUrl(latest.sha, format)))));
  } catch {
    return null;
  }
}

const THEME = 'betteroffice-benchmarks-theme';

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme;
  find('theme').setAttribute('aria-label', theme === 'light' ? 'Switch to dark mode' : 'Switch to light mode');
}

find('theme').addEventListener('click', () => {
  const next = document.documentElement.dataset.theme === 'light' ? 'dark' : 'light';
  applyTheme(next);
  try {
    localStorage.setItem(THEME, next);
  } catch {}
});
applyTheme(document.documentElement.dataset.theme === 'light' ? 'light' : 'dark');

const day = (iso: string) =>
  new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

function fail(message: string): void {
  const notice = empty(
    `The latest benchmark run could not be loaded: ${message}.`,
    `${REPO}#benchmarks`,
    'Read the generated tables on GitHub ↗'
  );
  notice.classList.add('col-span-full');
  find('kpis').replaceChildren(notice);
}

async function boot(): Promise<void> {
  const latencyPromise = endToEnd();
  let loaded: Awaited<ReturnType<typeof benchmarks>> | null = null;
  try {
    loaded = await benchmarks();
  } catch (error) {
    fail(error instanceof Error ? error.message : String(error));
  }
  const latency = await latencyPromise;
  if (loaded) {
    const { summary, publishedAt, href } = loaded;
    const present = FORMATS.filter((format) => summary.formats[format].documents > 0);
    const total = present.reduce((sum, format) => sum + summary.formats[format].documents, 0);
    const releases = present.map((format) => `${format.toUpperCase()} ${summary.versions[format]}`).join(' · ');
    find('lede').textContent = `${thousands(total)} real documents, scored page by page against Microsoft Office${
      summary.libreoffice ? ' and compared with LibreOffice' : ''
    }.`;
    const span = (versions: string[]) =>
      versions.length > 1 ? `${versions[0]}–${versions.at(-1)}` : (versions[0] ?? '');
    const references = present.map((format) => summary.formats[format].reference).filter((entry) => entry !== null);
    const officeVersions = [...new Set(references.flatMap((entry) => entry.versions))].sort(compareVersions);
    const macos = [...new Set(references.flatMap((entry) => entry.os))].sort(compareVersions);
    const office = h(
      'span',
      {
        title: [
          ...references.map((entry) => `${entry.engine.replace('Microsoft ', '')} ${span(entry.versions)}`),
          ...(macos.length ? [`macOS ${span(macos)}`] : []),
        ].join(' · '),
      },
      `Microsoft Office ${span(officeVersions)}`.trim()
    );
    const facts: [string, Node | string][] = [
      ['Release', h('a', { class: LINK, href: 'https://www.npmjs.com/org/betteroffice' }, releases)],
      ...(references.length ? ([['Reference', office]] as [string, Node][]) : []),
      ...(summary.libreoffice ? ([['Baseline', `LibreOffice ${summary.libreoffice}`]] as [string, string][]) : []),
      ['Updated', local('report') ? 'Loaded report' : publishedAt ? day(publishedAt) : '—'],
      ['Method', h('a', { class: LINK, href: METHOD }, 'How it’s measured')],
    ];
    find('facts').replaceChildren(
      ...facts.map(([term, value]) =>
        h('div', { class: 'grid gap-0.5' }, h('dt', { class: 'text-[12px] text-dim' }, term), h('dd', { class: 'text-[13.5px]' }, value))
      )
    );
    for (const format of FORMATS)
      if (!present.includes(format))
        document.querySelector(`nav a[href="#${format}"]`)?.setAttribute('hidden', '');
    find('kpis').replaceChildren(...heroKpis(summary, latency));
    const board = find('scoreboard');
    board.hidden = false;
    board.append(scoreboard(summary));
    const formats = find('formats');
    for (const format of present) {
      const section = formatSection(summary, format);
      formats.append(section);
    }
  }
  const published = !local('e2e');
  const latencyNode = latencySection(latency, published);
  find('latency').replaceWith(latencyNode);
  const method = methodSection(loaded?.summary ?? null, loaded?.href ?? null, latency, published);
  find('method').replaceWith(method);
  document.body.classList.add('ready');
}

void boot();
