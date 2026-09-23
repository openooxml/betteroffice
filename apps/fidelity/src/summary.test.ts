import { expect, test } from 'bun:test';
import { timingSummary } from '../../../scripts/office-quality/docx-benchmark.mjs';
import { renderSection } from '../../../scripts/office-quality/readme.mjs';
import { roundtripSummary } from '../../../scripts/office-quality/roundtrip.mjs';
import { calculationSummary } from '../../../scripts/office-quality/xlsx-benchmark.mjs';
import {
  ENGINES,
  FORMATS,
  type Engine,
  type Format,
  headlineText,
  percentText,
  speedup,
  summarize,
} from './summary';

const commit = 'c'.repeat(40);
const office = '26.2.3.2';
const versions: Record<Format, string> = { docx: '0.2.1', pptx: '0.1.1', xlsx: '0.2.1' };
const reference: Record<Format, string> = {
  docx: 'Microsoft Word',
  pptx: 'Microsoft PowerPoint',
  xlsx: 'Microsoft Excel',
};

type Triple<T> = [commit: T, published: T, libreoffice: T];

interface Spec {
  ssim: Triple<number | null>;
  pages?: [reference: number, ...Triple<number>];
  render?: Triple<number | null>;
  calc?: { total: number; correct: Triple<number | null> };
  parsed?: Triple<boolean>;
  stalePublished?: boolean;
}

const five = (ms: number) => [0.9, 0.95, 1, 1.05, 1.1].map((factor) => ms * factor);

function identity(format: Format, engine: Engine, stale = false) {
  if (engine === 'commit') return { renderer_source_commit: commit };
  return { version: engine === 'published' ? (stale ? '0.0.9' : versions[format]) : office };
}

function sample(format: Format, id: string, spec: Spec) {
  const at = (engine: Engine) => ENGINES.indexOf(engine);
  const comparisons = ENGINES.map((engine) => {
    const ssim = spec.ssim[at(engine)] ?? null;
    const base = { channel: engine, ...identity(format, engine, engine === 'published' && spec.stalePublished) };
    if (ssim === null) return { ...base, status: 'failed', stage: 'capture', error: 'Could not render' };
    return {
      ...base,
      status: 'ok',
      source_verified: true,
      reference: { status: 'ok', sha256: `${id}-source`, engine: reference[format], version: '16.112.3' },
      actual: { status: 'ok', sha256: `${id}-source` },
      penalized_ssim: ssim,
      ...(spec.pages ? { reference_pages: spec.pages[0], actual_pages: spec.pages[1 + at(engine)] } : {}),
    };
  });
  const row: Record<string, unknown> = { id, format, comparisons };
  const byEngine = (make: (engine: Engine) => unknown) =>
    Object.fromEntries(ENGINES.map((engine) => [engine, make(engine)]));
  if (spec.render)
    row.native_timings = byEngine((engine) => {
      const ms = spec.render![at(engine)] ?? null;
      return ms === null ? { status: 'failed', error: 'timed out' } : { status: 'ok', elapsed_ms: five(ms) };
    });
  if (spec.calc)
    row.calculations = byEngine((engine) => {
      const correct = spec.calc!.correct[at(engine)] ?? null;
      const total = spec.calc!.total;
      return correct === null
        ? { status: 'failed', correct: 0, total, error: 'crashed' }
        : { status: 'ok', correct, total, elapsed_ms: five(10 + correct) };
    });
  row.roundtrip = byEngine((engine) => ({
    parse: spec.parsed?.[at(engine)] === false ? { status: 'failed', error: 'unreadable' } : { status: 'ok' },
    roundtrip: { status: 'failed', stage: 'edit', error: 'No eligible deterministic edit in the source' },
  }));
  return row;
}

function report() {
  const build = (format: Format) => ({
    libreoffice_version: office,
    published_version: versions[format],
    builds: { commit: { source_sha: commit } },
  });
  return {
    source_sha: commit,
    commit,
    versions,
    docx_benchmark: build('docx'),
    pptx_benchmark: build('pptx'),
    xlsx_benchmark: build('xlsx'),
    xlsx_fidelity_benchmark: { libreoffice_version: office, source_sha: commit },
    roundtrip_benchmark: { docx: build('docx'), pptx: build('pptx'), xlsx: build('xlsx') },
    samples: [
      sample('docx', 'memo', { ssim: [0.91, 0.84, 0.8], pages: [3, 3, 4, 3], render: [120, 150, 700] }),
      sample('docx', 'thesis', {
        ssim: [0.77, 0.7, 0.65],
        pages: [40, 40, 38, 43],
        render: [900, 950, 1100],
      }),
      sample('docx', 'form', { ssim: [0.88, null, 0.9], pages: [1, 1, 1, 1], render: [60, 55, null] }),
      sample('docx', 'letter', {
        ssim: [0.93, 0.92, 0.81],
        pages: [2, 2, 2, 3],
        render: [80, 90, 640],
        parsed: [true, true, false],
        stalePublished: true,
      }),
      sample('pptx', 'pitch', { ssim: [0.93, 0.9, 0.95], render: [110, 180, 900] }),
      sample('pptx', 'lecture', { ssim: [0.86, null, 0.84], render: [300, 420, 2600] }),
      sample('pptx', 'poster', { ssim: [0.9, 0.87, 0.92], render: [95, null, 810], parsed: [true, false, true] }),
      sample('xlsx', 'budget', { ssim: [0.8, 0.7, 0.66], calc: { total: 200, correct: [200, 150, 190] } }),
      sample('xlsx', 'ledger', { ssim: [0.72, 0.69, null], calc: { total: 50, correct: [50, 50, 50] } }),
      sample('xlsx', 'model', { ssim: [0.75, 0.73, 0.7], calc: { total: 900, correct: [897, 400, null] } }),
      sample('xlsx', 'chart', { ssim: [null, 0.6, 0.62], parsed: [false, true, true] }),
    ],
  };
}

/** Maps each generated README row label to its published, commit and LibreOffice cells. */
function readmeRows(section: string, format: Format): Map<string, string[]> {
  const start = section.indexOf(`### ${format.toUpperCase()}`);
  const end = section.indexOf('### ', start + 4);
  const block = section.slice(start, end < 0 ? undefined : end);
  const rows = new Map<string, string[]>();
  for (const match of block.matchAll(/<tr><td>([^<]+)<\/td>(.*?)<\/tr>/g))
    rows.set(match[1]!, [...match[2]!.matchAll(/<td align="right">([^<]*)<\/td>/g)].map((cell) => cell[1]!));
  return rows;
}

const readmeOrder: Engine[] = ['published', 'commit', 'libreoffice'];

type Channels<T> = Record<Engine, T>;
interface TimingHelper {
  common: number;
  channels: Channels<{ successful: number; mean_ms: number }>;
}
interface CalculationHelper {
  workbooks: number;
  common: number;
  channels: Channels<{ correct: number; total: number; mean_ms: number }>;
}
type ParsingHelper = Channels<{ parsed: number; total: number }>;

test('every aggregate agrees with the generated README tables', () => {
  const input = report();
  const section = renderSection(input);
  const summary = summarize(input);
  for (const format of FORMATS) {
    const rows = readmeRows(section, format);
    const { fidelity, render, calculation, parsing } = summary.formats[format];
    const cells = (make: (engine: Engine) => string) => readmeOrder.map(make);
    expect(rows.get('SSIM')).toEqual(cells((engine) => fidelity[engine]?.mean?.toFixed(4) ?? '—'));
    expect(rows.get('Scored/total')).toEqual(
      cells((engine) => `${fidelity[engine]!.scored}/${fidelity[engine]!.total}`)
    );
    if (format === 'docx') {
      expect(rows.get('Exact page counts')).toEqual(
        cells((engine) => `${fidelity[engine]!.exact}/${fidelity[engine]!.paged}`)
      );
      expect(rows.get('Absolute page error')).toEqual(cells((engine) => String(fidelity[engine]!.pageError)));
    }
    if (render)
      expect(rows.get('Render time (avg)')).toEqual(
        cells((engine) => `${render.engines[engine].meanMs!.toFixed(0)} ms`)
      );
    if (calculation) {
      expect(rows.get('Recalc accuracy')).toEqual(
        cells((engine) => {
          const { correct, total } = calculation.engines[engine];
          return `${((100 * correct) / total).toFixed(2)}%`;
        })
      );
      expect(rows.get('Recalc time (avg)')).toEqual(
        cells((engine) => `${calculation.engines[engine].meanMs!.toFixed(0)} ms`)
      );
    }
    expect(rows.get('Parse success')).toEqual(
      cells((engine) => `${((100 * parsing![engine].parsed) / parsing![engine].total).toFixed(2)}%`)
    );
  }
});

test('timing, recalculation and parsing match the benchmark helpers exactly', () => {
  const input = report();
  const summary = summarize(input);
  for (const format of ['docx', 'pptx'] as const) {
    const expected = timingSummary(input.samples, format) as unknown as TimingHelper;
    const render = summary.formats[format].render!;
    expect(render.common).toBe(expected.common);
    for (const engine of ENGINES) {
      expect(render.engines[engine].successful).toBe(expected.channels[engine].successful);
      expect(render.engines[engine].meanMs).toBeCloseTo(expected.channels[engine].mean_ms, 9);
    }
  }
  const calc = calculationSummary(input.samples) as unknown as CalculationHelper;
  const calculation = summary.formats.xlsx.calculation!;
  expect([calculation.workbooks, calculation.common]).toEqual([calc.workbooks, calc.common]);
  for (const engine of ENGINES) {
    expect(calculation.engines[engine].correct).toBe(calc.channels[engine].correct);
    expect(calculation.engines[engine].total).toBe(calc.channels[engine].total);
    expect(calculation.engines[engine].meanMs).toBeCloseTo(calc.channels[engine].mean_ms, 9);
  }
  expect(calculation.engines.commit.perfect).toBe(2);
  expect(calculation.engines.libreoffice.perfect).toBe(1);
  for (const format of FORMATS) {
    const expected = roundtripSummary(input.samples, format) as unknown as ParsingHelper;
    for (const engine of ENGINES) {
      expect(summary.formats[format].parsing![engine].parsed).toBe(expected[engine].parsed);
      expect(summary.formats[format].parsing![engine].total).toBe(expected[engine].total);
    }
  }
});

test('a perfect recalculation without valid timings stays out of the timing mean', () => {
  const input = report();
  input.samples.push(sample('xlsx', 'census', { ssim: [0.7, 0.7, 0.7], calc: { total: 10, correct: [10, 10, 10] } }));
  const ledger = input.samples.find((row) => row.id === 'ledger') as { calculations: Record<Engine, Record<string, unknown>> };
  delete ledger.calculations.published.elapsed_ms;
  const calculation = summarize(input).formats.xlsx.calculation!;
  expect(calculation.common).toBe(1);
  for (const engine of ENGINES) expect(calculation.engines[engine].meanMs).toBeCloseTo(20, 9);
  expect(calculation.engines.published.perfect).toBe(2);
});

test('rows keep per-document values and drop scores from another revision', () => {
  const summary = summarize(report());
  const letter = summary.formats.docx.rows.find((row) => row.id === 'letter')!;
  expect(letter.ssim).toEqual({ commit: 0.93, published: null, libreoffice: 0.81 });
  expect(letter.pages).toEqual({ commit: 2, published: null, libreoffice: 3 });
  expect(letter.referencePages).toBe(2);
  expect(letter.renderMs.libreoffice).toBeCloseTo(640, 9);
  expect(letter.parsed).toEqual({ commit: true, published: true, libreoffice: false });
  const form = summary.formats.docx.rows.find((row) => row.id === 'form')!;
  expect(form.renderMs.libreoffice).toBeNull();
  expect(summary.formats.docx.reference).toEqual({ engine: 'Microsoft Word', versions: ['16.112.3'], os: [] });
  expect(summary.formats.xlsx.rows.find((row) => row.id === 'model')!.recalc.libreoffice).toEqual({
    ok: false,
    correct: 0,
    total: 900,
    ms: null,
  });
  expect(summary.libreoffice).toBe(office);
  expect(summary.versions).toEqual(versions);
});

test('an unmeasured engine or benchmark is absent rather than zero', () => {
  const input = report() as Record<string, unknown>;
  delete input.docx_benchmark;
  delete input.xlsx_benchmark;
  delete input.roundtrip_benchmark;
  const summary = summarize(input);
  expect(summary.formats.docx.render).toBeNull();
  expect(summary.formats.docx.fidelity.libreoffice).toBeUndefined();
  expect(summary.formats.xlsx.calculation).toBeNull();
  expect(summary.formats.pptx.parsing).toBeNull();
  expect(summary.formats.pptx.fidelity.libreoffice!.scored).toBe(3);
});

test('speedup reads as how many times faster ours is', () => {
  expect(speedup(100, 910)).toBeCloseTo(9.1, 9);
  expect(speedup(null, 910)).toBeNull();
  expect(speedup(100, null)).toBeNull();
});

test('a report without a commit or samples is refused', () => {
  expect(() => summarize({ commit: 'nope', samples: [] })).toThrow('Malformed');
  expect(() => summarize({ commit })).toThrow('Malformed');
  expect(() => summarize(null)).toThrow('Malformed');
});

test('table cells round like the README; headline figures never round a partial share to 100%', () => {
  expect(percentText(125_788, 125_794)).toBe('100.00%');
  expect(headlineText(125_788, 125_794)).toBe('99.99%');
  expect(headlineText(199, 200, 0)).toBe('99%');
  expect(headlineText(200, 200, 0)).toBe('100%');
  expect(percentText(125_290, 125_794)).toBe(headlineText(125_290, 125_794));
  expect(percentText(1, 0)).toBe('—');
});
