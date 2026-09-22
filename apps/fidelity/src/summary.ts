export const FORMATS = ['docx', 'pptx', 'xlsx'] as const;
export type Format = (typeof FORMATS)[number];

/** Display order: the measured commit, the latest release, then LibreOffice. */
export const ENGINES = ['commit', 'published', 'libreoffice'] as const;
export type Engine = (typeof ENGINES)[number];

export interface Fidelity {
  scored: number;
  total: number;
  mean: number | null;
  exact: number;
  paged: number;
  pageError: number;
}

export interface Timing {
  successful: number;
  meanMs: number | null;
}

export interface Calculation {
  correct: number;
  total: number;
  perfect: number;
  meanMs: number | null;
}

export interface Recalc {
  ok: boolean;
  correct: number;
  total: number;
  ms: number | null;
}

export interface DocumentRow {
  id: string;
  format: Format;
  referencePages: number | null;
  ssim: Record<Engine, number | null>;
  pages: Record<Engine, number | null>;
  renderMs: Record<Engine, number | null>;
  recalc: Record<Engine, Recalc | null>;
  parsed: Record<Engine, boolean | null>;
}

export interface FormatSummary {
  format: Format;
  version: string;
  documents: number;
  reference: { engine: string; versions: string[] } | null;
  fidelity: Partial<Record<Engine, Fidelity>>;
  render: { common: number; engines: Record<Engine, Timing> } | null;
  calculation: { workbooks: number; common: number; engines: Record<Engine, Calculation> } | null;
  parsing: Record<Engine, { parsed: number; total: number }> | null;
  rows: DocumentRow[];
}

export interface Summary {
  commit: string;
  versions: Record<Format, string>;
  libreoffice: string | null;
  formats: Record<Format, FormatSummary>;
}

type Json = Record<string, unknown>;

const object = (value: unknown): Json | null =>
  value && typeof value === 'object' && !Array.isArray(value) ? (value as Json) : null;
const list = (value: unknown): Json[] =>
  Array.isArray(value) ? value.map(object).filter((entry): entry is Json => entry !== null) : [];
const integer = (value: unknown): number | null => (Number.isInteger(value) ? (value as number) : null);
const finite = (value: unknown): number | null =>
  typeof value === 'number' && Number.isFinite(value) ? value : null;
const text = (value: unknown): string | null => (typeof value === 'string' && value ? value : null);

function mean(values: number[]): number | null {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
}

function trials(value: unknown): number | null {
  const timing = object(value);
  if (timing?.status !== 'ok' || !Array.isArray(timing.elapsed_ms) || !timing.elapsed_ms.length) return null;
  const elapsed = timing.elapsed_ms.map(finite);
  return elapsed.every((ms): ms is number => ms !== null && ms > 0) ? mean(elapsed as number[]) : null;
}

function comparison(sample: Json, engine: Engine, revision: string | null): Json | null {
  if (!revision) return null;
  const key = engine === 'commit' ? 'renderer_source_commit' : 'version';
  const found = list(sample.comparisons).find(
    (entry) => entry.channel === engine && entry[key] === revision
  );
  return found && found.status !== 'failed' && finite(found.penalized_ssim) !== null ? found : null;
}

function fidelity(rows: Json[], engine: Engine, revision: string | null): Fidelity {
  const scored = rows.map((row) => comparison(row, engine, revision)).filter((entry) => entry !== null);
  const paged = scored.filter(
    (entry) => integer(entry.reference_pages) !== null && integer(entry.actual_pages) !== null
  );
  return {
    scored: scored.length,
    total: rows.length,
    mean: mean(scored.map((entry) => entry.penalized_ssim as number)),
    exact: paged.filter((entry) => entry.reference_pages === entry.actual_pages).length,
    paged: paged.length,
    pageError: paged.reduce(
      (sum, entry) => sum + Math.abs((entry.reference_pages as number) - (entry.actual_pages as number)),
      0
    ),
  };
}

function recalc(value: unknown): Recalc | null {
  const result = object(value);
  const correct = integer(result?.correct);
  const total = integer(result?.total);
  if (!result || correct === null || total === null) return null;
  return { ok: result.status === 'ok', correct, total, ms: trials(result) };
}

function referenceEngine(rows: Json[], revision: string): FormatSummary['reference'] {
  const engines = new Map<string, number>();
  const versions = new Set<string>();
  for (const row of rows) {
    const reference = object(comparison(row, 'commit', revision)?.reference);
    const engine = text(reference?.engine);
    if (!engine) continue;
    engines.set(engine, (engines.get(engine) ?? 0) + 1);
    const version = text(reference?.version);
    if (version) versions.add(version);
  }
  const engine = [...engines].sort((left, right) => right[1] - left[1])[0]?.[0];
  if (!engine) return null;
  return { engine, versions: [...versions].sort(compareVersions) };
}

export function compareVersions(left: string, right: string): number {
  const [a, b] = [left.split('.').map(Number), right.split('.').map(Number)];
  for (let index = 0; index < Math.max(a.length, b.length); index += 1)
    if ((a[index] ?? 0) !== (b[index] ?? 0)) return (a[index] ?? 0) - (b[index] ?? 0);
  return 0;
}

function byEngine<T>(make: (engine: Engine) => T): Record<Engine, T> {
  return Object.fromEntries(ENGINES.map((engine) => [engine, make(engine)])) as Record<Engine, T>;
}

function summarizeFormat(root: Json, format: Format): FormatSummary {
  const commit = String(root.commit);
  const version = String(object(root.versions)?.[format] ?? '');
  const rows = list(root.samples).filter((sample) => sample.format === format);
  const measured = object(root[format === 'xlsx' ? 'xlsx_fidelity_benchmark' : `${format}_benchmark`]);
  const office = text(measured?.libreoffice_version);
  const revisions: Record<Engine, string | null> = { commit, published: version, libreoffice: office };

  const documents: DocumentRow[] = rows.map((row) => {
    const scored = byEngine((engine) => comparison(row, engine, revisions[engine]));
    const timings = object(row.native_timings);
    const calculations = object(row.calculations);
    const roundtrip = object(row.roundtrip);
    return {
      id: String(row.id),
      format,
      referencePages:
        ENGINES.map((engine) => integer(scored[engine]?.reference_pages)).find((pages) => pages !== null) ??
        null,
      ssim: byEngine((engine) => finite(scored[engine]?.penalized_ssim)),
      pages: byEngine((engine) => integer(scored[engine]?.actual_pages)),
      renderMs: byEngine((engine) => trials(timings?.[engine])),
      recalc: byEngine((engine) => recalc(calculations?.[engine])),
      parsed: byEngine((engine) => {
        const status = object(object(roundtrip?.[engine])?.parse)?.status;
        return status === 'ok' ? true : status === 'failed' ? false : null;
      }),
    };
  });

  let render: FormatSummary['render'] = null;
  if (format !== 'xlsx' && object(root[`${format}_benchmark`])) {
    const common = documents.filter((row) => ENGINES.every((engine) => row.renderMs[engine] !== null));
    render = {
      common: common.length,
      engines: byEngine((engine) => ({
        successful: documents.filter((row) => row.renderMs[engine] !== null).length,
        meanMs: mean(common.map((row) => row.renderMs[engine] as number)),
      })),
    };
  }

  let calculation: FormatSummary['calculation'] = null;
  if (format === 'xlsx' && object(root.xlsx_benchmark)) {
    const workbooks = documents.filter((row) => ENGINES.every((engine) => row.recalc[engine] !== null));
    const perfect = (result: Recalc | null) => !!result?.ok && result.correct === result.total;
    const common = workbooks.filter((row) => ENGINES.every((engine) => perfect(row.recalc[engine])));
    calculation = {
      workbooks: workbooks.length,
      common: common.length,
      engines: byEngine((engine) => ({
        correct: workbooks.reduce((sum, row) => sum + row.recalc[engine]!.correct, 0),
        total: workbooks.reduce((sum, row) => sum + row.recalc[engine]!.total, 0),
        perfect: workbooks.filter((row) => perfect(row.recalc[engine])).length,
        meanMs: mean(common.map((row) => row.recalc[engine]!.ms ?? 0)),
      })),
    };
  }

  const parsing = object(object(root.roundtrip_benchmark)?.[format])
    ? byEngine((engine) => ({
        parsed: documents.filter((row) => row.parsed[engine] === true).length,
        total: documents.length,
      }))
    : null;

  return {
    format,
    version,
    documents: documents.length,
    reference: referenceEngine(rows, commit),
    fidelity: {
      commit: fidelity(rows, 'commit', commit),
      published: fidelity(rows, 'published', version),
      ...(office ? { libreoffice: fidelity(rows, 'libreoffice', office) } : {}),
    },
    render,
    calculation,
    parsing,
    rows: documents,
  };
}

/** Aggregates a published report.json exactly as the README's generated tables do. */
export function summarize(value: unknown): Summary {
  const root = object(value);
  if (!root || !/^[a-f0-9]{40}$/.test(String(root.commit)) || !Array.isArray(root.samples))
    throw new Error('Malformed benchmark report');
  const formats = Object.fromEntries(
    FORMATS.map((format) => [format, summarizeFormat(root, format)])
  ) as Record<Format, FormatSummary>;
  const office = ['docx_benchmark', 'pptx_benchmark', 'xlsx_fidelity_benchmark', 'xlsx_benchmark']
    .map((key) => text(object(root[key])?.libreoffice_version))
    .find(Boolean);
  return {
    commit: String(root.commit),
    versions: byFormat((format) => formats[format].version),
    libreoffice: office ?? null,
    formats,
  };
}

function byFormat<T>(make: (format: Format) => T): Record<Format, T> {
  return Object.fromEntries(FORMATS.map((format) => [format, make(format)])) as Record<Format, T>;
}

/** Ratio of `other` over `ours`: above one means ours is faster. */
export function speedup(ours: number | null, other: number | null): number | null {
  return ours && other ? other / ours : null;
}

export const ssimText = (value: number | null | undefined): string =>
  value === null || value === undefined ? '—' : value.toFixed(4);

/** A percentage that never rounds a partial share up to 100. */
export function percentage(part: number, whole: number, digits = 2): number | null {
  if (!whole) return null;
  const value = Number(((100 * part) / whole).toFixed(digits));
  return part < whole && value >= 100 ? Number((100 - 10 ** -digits).toFixed(digits)) : value;
}

/** Rounds like the README tables, so scoreboard cells match them exactly. */
export const percentText = (part: number, whole: number, digits = 2): string =>
  whole ? `${((100 * part) / whole).toFixed(digits)}%` : '—';

/** For headline figures, where a near-total share must not read as 100%. */
export const headlineText = (part: number, whole: number, digits = 2): string => {
  const value = percentage(part, whole, digits);
  return value === null ? '—' : `${value.toFixed(digits)}%`;
};

export const msText = (value: number | null | undefined): string =>
  value === null || value === undefined ? '—' : `${Math.round(value).toLocaleString('en-US')} ms`;
