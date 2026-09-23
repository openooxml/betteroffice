import {
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
const listed = (names: string[]) =>
  names.length > 1 ? `${names.slice(0, -1).join(', ')} and ${names.at(-1)}` : (names[0] ?? '');

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

function legend(entries: EngineView[]): HTMLElement {
  return h(
    'ul',
    { class: 'legend' },
    ...entries.map((entry) => h('li', {}, h('i', { class: `swatch ${entry.cls}` }), `${entry.label} `, h('span', {}, entry.sub)))
  );
}

function count(value: number, digits: number, suffix = ''): HTMLElement {
  const text = value.toFixed(digits);
  return h('span', { class: 'count', 'data-to': text, 'data-digits': digits, 'data-suffix': suffix }, `${text}${suffix}`);
}

interface KpiBar {
  cls: string;
  label: string;
  text: string;
  share: number;
}

function kpi(tag: string, tagCls: string, label: string, value: HTMLElement, bars: KpiBar[], foot: string) {
  return h(
    'article',
    { class: 'kpi' },
    h('div', { class: 'kpi-head' }, h('span', { class: `tag ${tagCls}` }, tag), h('span', {}, label)),
    h('div', { class: 'kpi-value' }, value),
    h(
      'div',
      { class: 'kpi-bars' },
      ...bars.map((bar) =>
        h(
          'div',
          { class: `kbar ${bar.cls}` },
          h('span', {}, bar.label),
          h('i', {}, h('i', { style: `--w:${(Math.max(0, Math.min(1, bar.share)) * 100).toFixed(2)}%` })),
          h('b', {}, bar.text)
        )
      )
    ),
    h('p', { class: 'kpi-foot' }, foot)
  );
}

function withUnit(node: HTMLElement, unit: string): HTMLElement {
  return h('span', {}, node, h('span', { class: 'unit' }, unit));
}

function heroKpis(summary: Summary, latency: Latency | null): HTMLElement[] {
  const tiles: HTMLElement[] = [];
  const docx = summary.formats.docx;
  const pages = docx.fidelity[OURS];
  if (pages?.paged) {
    const engine = views(summary, 'docx');
    tiles.push(
      kpi(
        'DOCX',
        'f-docx',
        'Page counts match Word',
        withUnit(count(percentage(pages.exact, pages.paged, pages.exact === pages.paged ? 0 : 1)!, pages.exact === pages.paged ? 0 : 1), '%'),
        SHOWN.filter((key) => docx.fidelity[key]?.paged).map((key) => ({
          cls: engine[key].cls,
          label: key === 'libreoffice' ? 'LibreOffice' : engine[key].sub,
          text: `${docx.fidelity[key]!.exact}/${docx.fidelity[key]!.paged}`,
          share: docx.fidelity[key]!.exact / docx.fidelity[key]!.paged,
        })),
        `${pages.exact} of ${pages.paged} documents render as many pages as Word does.`
      )
    );
  }
  const calc = summary.formats.xlsx.calculation;
  if (calc?.engines[OURS].total) {
    const engine = views(summary, 'xlsx');
    const ours = calc.engines[OURS];
    tiles.push(
      kpi(
        'XLSX',
        'f-xlsx',
        'Formulas match Excel',
        withUnit(count(percentage(ours.correct, ours.total, 2)!, 2), '%'),
        SHOWN.map((key) => ({
          cls: engine[key].cls,
          label: key === 'libreoffice' ? 'LibreOffice' : engine[key].sub,
          text: headlineText(calc.engines[key].correct, calc.engines[key].total),
          share: calc.engines[key].correct / Math.max(1, calc.engines[key].total),
        })),
        `${thousands(ours.correct)} of ${thousands(ours.total)} formula results match Excel after a full recalculation.`
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
        'PPTX',
        'f-pptx',
        ratio >= 1 ? 'Faster slides than LibreOffice' : 'Slower slides than LibreOffice',
        withUnit(count(ratio >= 1 ? ratio : 1 / ratio, 1), '×'),
        SHOWN.map((key) => ({
          cls: engine[key].cls,
          label: key === 'libreoffice' ? 'LibreOffice' : engine[key].sub,
          text: msText(render.engines[key].meanMs),
          share: (render.engines[key].meanMs ?? 0) / slowest,
        })),
        `Native CLI, first slide at 96 DPI, process start included. Mean over the ${render.common} decks every engine rendered.`
      )
    );
  }
  if (latency?.ops) {
    tiles.push(
      kpi(
        'E2E',
        'f-e2e',
        'Calls inside one 60 Hz frame',
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
        `${thousands(latency.withinFrame)} of ${thousands(latency.ops)} timed SDK calls across ${latency.cases} end-to-end cases on main (${short(latency.commit)}) took 16.7 ms or less.`
      )
    );
  } else {
    const parsed = FORMATS.reduce((sum, format) => sum + (summary.formats[format].parsing?.[OURS].parsed ?? 0), 0);
    const total = FORMATS.reduce((sum, format) => sum + (summary.formats[format].parsing?.[OURS].total ?? 0), 0);
    if (total)
      tiles.push(
        kpi(
          'ALL',
          'f-e2e',
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
          `${thousands(parsed)} of ${thousands(total)} corpus files open in the native API.`
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
      cells: fidelity((engine) => {
        const mean = entry.fidelity[engine]!.mean;
        return mean === null ? null : { value: mean, text: ssimText(mean) };
      }),
    },
    {
      label: 'Scored / total',
      hint: `${NOUNS[format][1]} with a score`,
      better: 'high',
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
      cells: byEngine((engine) => {
        const { parsed, total } = entry.parsing![engine];
        return total ? { value: parsed / total, text: percentText(parsed, total) } : null;
      }),
    });
  return rows;
}

function scoreboard(summary: Summary): HTMLElement {
  const lo = summary.libreoffice ?? '';
  const board = h(
    'div',
    { class: 'board', role: 'table' },
    h(
      'div',
      { class: 'board-head', role: 'row' },
      h('span', { role: 'columnheader' }),
      h('span', { role: 'columnheader', class: 'c-ours' }, h('i', { class: 'swatch c-ours' }), 'BetterOffice', h('small', {}, 'latest release')),
      h('span', { role: 'columnheader', class: 'c-libreoffice' }, h('i', { class: 'swatch c-libreoffice' }), 'LibreOffice', h('small', {}, lo))
    )
  );
  const engine = (format: Format) => views(summary, format);
  for (const format of FORMATS) {
    const entry = summary.formats[format];
    if (!entry.documents) continue;
    board.append(
      h(
        'a',
        { class: 'board-group', href: `#${format}` },
        h('span', { class: `tag f-${format}` }, format.toUpperCase()),
        h('span', {}, TITLES[format]),
        h('small', {}, `${plural(entry.documents, NOUNS[format])} · release v${entry.version}`)
      )
    );
    for (const metric of metrics(summary, format)) {
      const present = SHOWN.map((key) => metric.cells[key]?.value).filter((value): value is number => value !== undefined);
      const best = metric.better === 'high' ? Math.max(...present) : Math.min(...present);
      const tie = present.every((value) => value === best);
      const share = (value: number) =>
        metric.better === 'high' ? (best > 0 ? value / best : 0) : value === best ? 1 : best > 0 ? best / value : 0;
      board.append(
        h(
          'div',
          { class: 'board-row', role: 'row' },
          h('span', { class: 'metric', role: 'rowheader' }, metric.label, h('small', {}, metric.hint)),
          ...SHOWN.map((key) => {
            const cell = metric.cells[key];
            const view = engine(format)[key];
            const tag = key === 'libreoffice' ? view.label : view.sub;
            if (!cell) return h('span', { class: `score ${view.cls} empty`, role: 'cell', 'data-engine': tag }, '—');
            const top = present.length > 1 && !tie && cell.value === best;
            return h(
              'span',
              { class: `score ${view.cls}${top ? ' best' : ''}`, role: 'cell', 'data-engine': tag },
              h('b', {}, cell.text),
              h('i', { class: 'track' }, h('i', { style: `--w:${(share(cell.value) * 100).toFixed(2)}%` }))
            );
          })
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

function panel(title: string, caption: string, stat: HTMLElement | null, body: HTMLElement, wide = false) {
  return h(
    'figure',
    { class: `panel${wide ? ' wide' : ''}` },
    h('figcaption', {}, h('div', {}, h('h3', {}, title), h('p', {}, caption)), stat),
    body
  );
}

function stat(value: string, note: string, cls = ''): HTMLElement {
  return h('div', { class: `figure-stat ${cls}` }, h('b', {}, value), h('span', {}, note));
}

function chart(build: (width: number) => Element, label: string): HTMLElement {
  const container = h('div', { class: 'plot' });
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
      ? 'Each dot is one workbook, scored over its recorded print ranges against Excel; BetterOffice renders with its browser renderer at 150 DPI. Hover to trace a workbook across engines, click to open its pages.'
      : `Each dot is one ${NOUNS[entry.format][0]}’s page-penalized SSIM against ${OFFICE[entry.format]}; BetterOffice renders with its browser renderer at 150 DPI. Hover to trace it across engines, click to open its pages.`;
  return panel(
    `Visual similarity, ${NOUNS[entry.format][0]} by ${NOUNS[entry.format][0]}`,
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
  const blocks = h('div', { class: 'waffles' });
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
      h(
        'div',
        { class: `waffle-block ${engine[key].cls}` },
        h('div', { class: 'waffle-label' }, h('i', { class: `swatch ${engine[key].cls}` }), engine[key].label, h('span', {}, engine[key].sub)),
        labelled(waffle(cells, columns(rows.length), open), `${engine[key].label} ${engine[key].sub}: ${fidelity.exact} of ${fidelity.paged} page counts exact`),
        h('div', { class: 'waffle-stat' }, h('b', {}, `${fidelity.exact}/${fidelity.paged}`), h('span', {}, 'exact')),
        h('div', { class: 'waffle-sub' }, `${plural(fidelity.pageError, ['page', 'pages'])} off`)
      )
    );
  }
  const ours = entry.fidelity[OURS]!;
  return panel(
    'Page count against Word',
    'One square per document, smallest first, in the same position for every engine.',
    stat(`${ours.exact}/${ours.paged}`, `exact in v${entry.version}`, ours.exact >= (entry.fidelity.libreoffice?.exact ?? 0) ? 'up' : 'down'),
    h(
      'div',
      {},
      blocks,
      h(
        'ul',
        { class: 'key' },
        h('li', {}, h('i', { class: 'cell s-exact' }), 'exact'),
        h('li', {}, h('i', { class: 'cell s-near' }), 'one page off'),
        h('li', {}, h('i', { class: 'cell s-far' }), 'two or more off'),
        h('li', {}, h('i', { class: 'cell s-none' }), 'no score')
      )
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
      { class: 'legend chart-legend' },
      ...series.map((entry) =>
        h('li', {}, h('i', { class: `swatch ${entry.cls}` }), entry.label, h('span', {}, `median ${duration(median(entry.values))}`))
      )
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
    `First-${unit} render time`,
    `The native CLI’s Rust rasterizer, not the browser renderer scored above: first ${unit} at 96 DPI, process start included, over the ${render.common} ${NOUNS[entry.format][1]} every engine rendered.`,
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
    { class: 'accuracy' },
    ...SHOWN.map((key) => {
      const { correct, total } = calc.engines[key];
      return h(
        'div',
        { class: `abar ${engine[key].cls}` },
        h('span', {}, engine[key].label, h('small', {}, engine[key].sub)),
        h('i', { class: 'track' }, h('i', { style: `--w:${((100 * correct) / Math.max(1, total)).toFixed(3)}%` })),
        h('b', {}, percentText(correct, total)),
        h('small', { class: 'abar-count' }, `${thousands(correct)} / ${thousands(total)}`)
      );
    })
  );
  const blocks = h('div', { class: 'waffles' });
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
      h(
        'div',
        { class: `waffle-block ${engine[key].cls}` },
        h('div', { class: 'waffle-label' }, h('i', { class: `swatch ${engine[key].cls}` }), engine[key].label, h('span', {}, engine[key].sub)),
        labelled(waffle(cells, columns(rows.length, 3.2), open), `${engine[key].label} ${engine[key].sub}: ${calc.engines[key].perfect} of ${rows.length} workbooks recalculate perfectly`),
        h('div', { class: 'waffle-stat' }, h('b', {}, String(calc.engines[key].perfect)), h('span', {}, 'perfect')),
        h('div', { class: 'waffle-sub' }, `of ${plural(rows.length, NOUNS.xlsx)}`)
      )
    );
  }
  const ours = calc.engines[OURS];
  return panel(
    'Formula results against Excel',
    `Formula caches are cleared, each workbook is recalculated from scratch, and every result cell is compared with Excel. One square per workbook with an Excel reference, largest first.`,
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
      h(
        'ul',
        { class: 'key' },
        h('li', {}, h('i', { class: 'cell s-exact' }), 'all match'),
        h('li', {}, h('i', { class: 'cell s-high' }), '≥ 99%'),
        h('li', {}, h('i', { class: 'cell s-near' }), '≥ 90%'),
        h('li', {}, h('i', { class: 'cell s-far' }), 'below 90%'),
        h('li', {}, h('i', { class: 'cell s-none' }), 'failed')
      )
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
    `Import, full recalculation and export by the native CLI, process start included, over the ${calc.common} workbooks every engine recalculates perfectly.`,
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
    return h(
      'li',
      {},
      h(
        'a',
        { href: compareHref(row.id), class: 'mover' },
        h('span', { class: 'mover-id' }, row.id),
        dumbbell(row.ssim.libreoffice, row.ssim[OURS]!, low, high),
        h('b', {}, ssimText(row.ssim[OURS])),
        h('span', { class: `delta ${delta === null ? 'new' : delta >= 0 ? 'up' : 'down'}` }, delta === null ? 'no LO score' : signed(delta))
      )
    );
  };
  const column = (title: string, rows: typeof scored) =>
    h('div', { class: 'movers-column' }, h('h4', {}, title), h('ol', {}, ...rows.map(item)));
  const note = failed.length ? `${plural(failed.length, NOUNS[entry.format])} did not score in ${version}.` : '';
  return panel(
    'Where to look',
    `Each row opens the page viewer. The bar runs from LibreOffice’s score to ${version}’s.`,
    null,
    h(
      'div',
      {},
      h('div', { class: 'movers' }, column('Largest leads over LibreOffice', leads), column(`Lowest in ${version}`, lowest)),
      note
        ? h(
            'p',
            { class: 'movers-note' },
            note,
            ...failed.slice(0, 6).flatMap((row) => [' ', h('a', { href: compareHref(row.id) }, row.id)])
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
  const panels = h('div', { class: 'panels' }, similarity(summary, entry));
  if (format === 'docx') panels.append(pagination(summary, entry));
  const render = renderTimes(summary, entry);
  if (render) panels.append(render);
  const accuracy = formulas(summary, entry);
  if (accuracy) panels.append(accuracy);
  const recalc = recalcTimes(summary, entry);
  if (recalc) panels.append(recalc);
  const halves = [...panels.children].filter((node) => !node.classList.contains('wide')).length;
  panels.append(movers(entry, halves % 2 === 0));
  return h(
    'section',
    { class: 'section format', id: format },
    h(
      'header',
      { class: 'section-head' },
      h('span', { class: `tag big f-${format}` }, format.toUpperCase()),
      h('h2', {}, TITLES[format]),
      h('p', { class: 'meta' }, `${plural(entry.documents, NOUNS[format])} · reference ${office}`),
      legend(SHOWN.filter((key) => entry.fidelity[key] || entry.render || entry.calculation).map((key) => engine[key]))
    ),
    panels
  );
}

function latencySection(latency: Latency | null, published: boolean): HTMLElement {
  const head = h(
    'header',
    { class: 'section-head' },
    h('span', { class: 'tag big f-e2e' }, 'E2E'),
    h('h2', {}, 'Editing latency'),
    h(
      'p',
      { class: 'meta' },
      'The end-to-end suite drives the WASM SDK in Bun and the Python bindings through scripted editing sessions on pinned documents, including concurrent editors and Python ↔ web collaboration. Every SDK call is timed.'
    )
  );
  if (!latency?.ops)
    return h(
      'section',
      { class: 'section', id: 'latency' },
      head,
      h(
        'div',
        { class: 'panel empty' },
        h('p', {}, 'Editing latency appears here once a green end-to-end run on main publishes its timings.'),
        h('a', { href: `${REPO}/actions/workflows/e2e.yml` }, 'End-to-end workflow ↗')
      )
    );
  const present = FORMATS.filter((format) => latency.formats[format]);
  const tiles = h(
    'div',
    { class: 'kpis compact' },
    kpi(
      'CASES',
      'f-e2e',
      'Scenarios passed',
      h('span', {}, count(latency.passed, 0), h('span', { class: 'unit' }, ` / ${latency.cases}`)),
      present.map((format) => ({
        cls: `f-${format}`,
        label: format.toUpperCase(),
        text: `${latency.formats[format]!.passed}/${latency.formats[format]!.cases.length}`,
        share: latency.formats[format]!.passed / Math.max(1, latency.formats[format]!.cases.length),
      })),
      `${plural(latency.multiEditor, ['case', 'cases'])} with several editors, ${latency.crossSdk} across Python and WASM.`
    ),
    kpi(
      'CALLS',
      'f-e2e',
      'Timed SDK calls',
      count(latency.ops, 0),
      present.map((format) => ({
        cls: `f-${format}`,
        label: format.toUpperCase(),
        text: thousands(latency.formats[format]!.ops),
        share: ratio(latency.formats[format]!.ops, Math.max(...present.map((key) => latency.formats[key]!.ops))),
      })),
      'Every open, edit, layout, sync, undo and save the scenarios make.'
    ),
    kpi(
      'MEDIAN',
      'f-e2e',
      'Median call',
      h('span', {}, count(latency.p50, latency.p50 < 10 ? 2 : 0), h('span', { class: 'unit' }, ' ms')),
      present.map((format) => ({
        cls: `f-${format}`,
        label: format.toUpperCase(),
        text: duration(latency.formats[format]!.p50),
        share: ratio(latency.formats[format]!.p50, Math.max(...present.map((key) => latency.formats[key]!.p50))),
      })),
      `Across all ${thousands(latency.ops)} calls. Each bar is one format’s median.`
    ),
    kpi(
      'P95',
      'f-e2e',
      '95th percentile call',
      h('span', {}, count(latency.p95, latency.p95 < 10 ? 1 : 0), h('span', { class: 'unit' }, ' ms')),
      present.map((format) => ({
        cls: `f-${format}`,
        label: format.toUpperCase(),
        text: duration(latency.formats[format]!.p95),
        share: ratio(latency.formats[format]!.p95, Math.max(...present.map((key) => latency.formats[key]!.p95))),
      })),
      `Across all ${thousands(latency.ops)} calls. Each bar is one format’s 95th percentile.`
    )
  );

  let selected: Format = present[0]!;
  let release: (() => void) | null = null;
  const switcher = h('div', { class: 'segmented', role: 'tablist' });
  const plot = h('div', { class: 'plot' });
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
    const button = h('button', { type: 'button', role: 'tab', 'data-format': format, class: `f-${format}` }, format.toUpperCase());
    button.addEventListener('click', () => {
      selected = format;
      draw();
    });
    switcher.append(button);
  }
  queueMicrotask(draw);

  const operations = h(
    'figure',
    { class: 'panel wide' },
    h(
      'figcaption',
      {},
      h(
        'div',
        {},
        h('h3', {}, 'Latency by operation'),
        h('p', {}, 'The dot is the median call, the bar runs to the 95th percentile. Operations are grouped by verb, most frequent first, sorted by median.')
      ),
      switcher
    ),
    plot
  );

  const all = panel(
    'Every timed call',
    'Share of SDK calls finished within a given time, per format.',
    stat(headlineText(latency.withinFrame, latency.ops, 1), 'inside one 60 Hz frame', 'up'),
    lines(
      present.map((format) => ({ cls: `f-${format}`, label: format.toUpperCase(), values: latency.formats[format]!.values })),
      { frame: FRAME_MS, noun: 'calls', height: 260 }
    )
  );

  const cases = h('div', { class: 'cases' });
  for (const format of present) {
    const entry = latency.formats[format]!;
    const byKey = new Map(entry.cases.map((item) => [`${item.scenario}/${item.sample}`, item]));
    cases.append(
      h(
        'div',
        { class: 'case-row' },
        h('span', { class: `tag f-${format}` }, format.toUpperCase()),
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
        h('b', {}, `${entry.passed}/${entry.cases.length}`)
      )
    );
  }
  const machine = latency.environment;
  const scenarios = panel(
    'Every case',
    'One square per scenario and document. Hover for what it does.',
    stat(`${latency.passed}/${latency.cases}`, 'passed', latency.passed === latency.cases ? 'up' : 'down'),
    h(
      'div',
      {},
      cases,
      h(
        'p',
        { class: 'fine' },
        `Recorded ${when(latency.recordedAt)} at `,
        h('a', { href: `${REPO}/commit/${latency.commit}` }, short(latency.commit)),
        machine ? ` on ${published ? 'a GitHub-hosted runner, ' : ''}${machine.cpu}, ${machine.cpus} vCPU. ` : '. ',
        'These are observations from one run, not controlled benchmark trials.'
      )
    )
  );

  return h('section', { class: 'section', id: 'latency' }, head, tiles, h('div', { class: 'panels' }, operations, all, scenarios));
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
  const card = (title: string, body: string) => h('article', { class: 'card' }, h('h3', {}, title), h('p', {}, body));
  const links: [string, string][] = [
    ['Benchmark methodology', METHOD],
    ['Benchmarks workflow', `${REPO}/actions/workflows/visual-fidelity.yml`],
    ['End-to-end suite', `${REPO}/blob/main/e2e/README.md`],
  ];
  if (reportHref) links.push(['Raw report.json', reportHref]);
  if (latency) links.push(['Raw e2e results', e2eUrl(latency.commit, FORMATS.find((format) => latency.formats[format])!)]);
  return h(
    'section',
    { class: 'section', id: 'method' },
    h('header', { class: 'section-head' }, h('h2', {}, 'How it’s measured')),
    h(
      'div',
      { class: 'cards' },
      card('Ground truth', `Microsoft Word, PowerPoint and Excel${office} on macOS exported every reference page to PDF, rasterized at 150 DPI. The corpus and its references are public.`),
      card('Visual similarity', 'Mean page-penalized grayscale SSIM at 150 DPI, without resampling or alignment, of the BetterOffice browser renderer. Missing or extra pages lower the score. Failed renders get no score, so read coverage next to the mean.'),
      card('Speed', `Native command-line builds time the whole job in five fresh processes after a warmup: start, import, fonts, layout, rasterize and write. They measure the Rust rasterizer at 96 DPI, not the browser renderer scored for similarity. LibreOffice ${lo} runs its prebuilt CLI with an isolated profile. Means cover the documents every engine handled.`),
      card('Formulas', 'Caches are cleared and each workbook is recalculated from scratch. Numbers match within 1e-9, or 1e-12 relative; text, booleans and errors must match exactly. Missing results count as wrong.'),
      card('Latency', `The end-to-end suite times every SDK call from invocation to result ${published ? 'on a GitHub-hosted runner' : 'on the machine that recorded it'}. Repeated calls see different document state, so these are observations, not benchmark trials.`)
    ),
    h('ul', { class: 'links' }, ...links.map(([label, href]) => h('li', {}, h('a', { href }, `${label} ↗`))))
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

const motion = !window.matchMedia('(prefers-reduced-motion: reduce)').matches;

function animate(root: Element): void {
  for (const node of root.querySelectorAll<HTMLElement>('.count')) {
    const to = Number(node.dataset.to);
    const digits = Number(node.dataset.digits ?? 0);
    const suffix = node.dataset.suffix ?? '';
    const final = `${to.toLocaleString('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits })}${suffix}`;
    if (!motion || !Number.isFinite(to)) {
      node.textContent = final;
      continue;
    }
    const started = performance.now();
    const frame = (now: number) => {
      const progress = Math.min(1, (now - started) / 1100);
      const eased = progress === 1 ? 1 : 1 - 2 ** (-10 * progress);
      node.textContent =
        progress === 1
          ? final
          : `${(to * eased).toLocaleString('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits })}${suffix}`;
      if (progress < 1) requestAnimationFrame(frame);
    };
    requestAnimationFrame(frame);
  }
}

function show(root: Element): void {
  root.classList.add('in');
  animate(root);
  window.setTimeout(() => root.classList.add('settled'), 1900);
}

const observer = new IntersectionObserver(
  (entries) => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      show(entry.target);
      observer.unobserve(entry.target);
    }
  },
  { threshold: 0.08 }
);

function reveal(root: Element): void {
  if (motion) observer.observe(root);
  else show(root);
}

function fail(message: string): void {
  find('kpis').replaceChildren(
    h(
      'div',
      { class: 'panel empty' },
      h('p', {}, `The latest benchmark run could not be loaded: ${message}.`),
      h('a', { href: `${REPO}#benchmarks` }, 'Read the generated tables on GitHub ↗')
    )
  );
  find('run-label').textContent = 'Unavailable';
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
    find('eyebrow').replaceChildren(
      h('i', { class: 'pulse' }),
      local('report') ? 'Loaded report · ' : 'Latest release · ',
      h('a', { href: 'https://www.npmjs.com/org/betteroffice' }, releases),
      publishedAt ? ` · measured ${when(publishedAt)}` : ''
    );
    const timed = present.some(
      (format) => summary.formats[format].render?.common || summary.formats[format].calculation?.common
    );
    const checks = [
      'each page is scored against the Office original',
      ...(summary.formats.xlsx.calculation?.workbooks ? ['each formula result is checked against Excel'] : []),
      ...(timed ? ['the native engines are timed on the same runner'] : []),
    ];
    find('lede').replaceChildren(
      h('b', {}, `${thousands(total)} real documents`),
      ` with reference pages exported from Microsoft ${listed(present.map((format) => OFFICE[format]))}. Each run renders them in the latest BetterOffice release${summary.libreoffice ? ` and in LibreOffice ${summary.libreoffice}` : ''}: ${listed(checks)}.`
    );
    const run = find('run');
    run.setAttribute('href', `${REPO}/actions/workflows/visual-fidelity.yml`);
    find('run-label').textContent = publishedAt
      ? `measured ${new Date(publishedAt).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`
      : 'benchmarks';
    for (const format of FORMATS)
      if (!present.includes(format))
        document.querySelector(`nav a[href="#${format}"]`)?.setAttribute('hidden', '');
    find('kpis').replaceChildren(...heroKpis(summary, latency));
    reveal(find('hero'));
    const board = find('scoreboard');
    board.hidden = false;
    board.append(scoreboard(summary));
    reveal(board);
    const formats = find('formats');
    for (const format of present) {
      const section = formatSection(summary, format);
      formats.append(section);
      reveal(section);
    }
  }
  const published = !local('e2e');
  const latencyNode = latencySection(latency, published);
  find('latency').replaceWith(latencyNode);
  reveal(latencyNode);
  const method = methodSection(loaded?.summary ?? null, loaded?.href ?? null, latency, published);
  find('method').replaceWith(method);
  reveal(method);
  document.body.classList.add('ready');
}

void boot();
