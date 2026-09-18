export const CORPUS_ORIGIN = 'https://corpus.betteroffice.dev';
export const RENDER_BASE = '/renders';

export interface PageScore {
  page: number;
  ssim: number | null;
  width: number | null;
  height: number | null;
}

export type Channel =
  | {
      status: 'ok';
      revision: string;
      penalizedSsim: number;
      commonPageSsim: number;
      referencePages: number;
      actualPages: number;
      pages: PageScore[];
    }
  | { status: 'failed'; revision: string; stage: string; error: string };

export interface Sample {
  id: string;
  format: string;
  published: Channel | null;
  commit: Channel | null;
  pages: PageScore[];
}

export interface Report {
  commit: string;
  versions: Record<string, string>;
  documents: Sample[];
}

export interface Latest {
  sha: string;
  publishedAt: string | null;
  samples: Record<string, number>;
}

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('Malformed fidelity report');
  return value as Record<string, unknown>;
}

function size(value: unknown, index: number): number | null {
  return Array.isArray(value) && typeof value[index] === 'number' ? (value[index] as number) : null;
}

function pageScores(value: unknown): PageScore[] {
  if (!Array.isArray(value)) throw new Error('Malformed fidelity report');
  return value
    .map((entry) => record(entry))
    .filter((entry) => Number.isInteger(entry.page))
    .map((entry) => ({
      page: entry.page as number,
      ssim: typeof entry.ssim === 'number' && Number.isFinite(entry.ssim) ? entry.ssim : null,
      width: size(entry.reference_size, 0) ?? size(entry.actual_size, 0),
      height: size(entry.reference_size, 1) ?? size(entry.actual_size, 1),
    }))
    .sort((left, right) => left.page - right.page);
}

function channel(value: Record<string, unknown>): Channel {
  const revision = String(value.renderer_source_commit ?? value.version ?? '');
  if (value.status === 'failed')
    return {
      status: 'failed',
      revision,
      stage: String(value.stage ?? 'unknown'),
      error: String(value.error ?? 'Unknown error'),
    };
  if (typeof value.penalized_ssim !== 'number' || !Number.isFinite(value.penalized_ssim))
    throw new Error('Malformed fidelity report');
  return {
    status: 'ok',
    revision,
    penalizedSsim: value.penalized_ssim,
    commonPageSsim: typeof value.common_page_ssim === 'number' ? value.common_page_ssim : 0,
    referencePages: Number(value.reference_pages ?? 0),
    actualPages: Number(value.actual_pages ?? 0),
    pages: pageScores(value.pages),
  };
}

export function parseReport(value: unknown): Report {
  const root = record(value);
  if (!/^[a-f0-9]{40}$/.test(String(root.commit))) throw new Error('Malformed fidelity report');
  if (!Array.isArray(root.samples)) throw new Error('Malformed fidelity report');
  const documents = root.samples.map((entry) => {
    const sample = record(entry);
    const comparisons = Array.isArray(sample.comparisons) ? sample.comparisons.map(record) : [];
    const pick = (name: string) =>
      comparisons.filter((item) => item.channel === name).map(channel)[0] ?? null;
    const commit = pick('commit');
    const published = pick('published');
    return {
      id: String(sample.id),
      format: String(sample.format),
      published,
      commit,
      pages: commit?.status === 'ok' ? commit.pages : published?.status === 'ok' ? published.pages : [],
    };
  });
  return {
    commit: String(root.commit),
    versions: record(root.versions ?? {}) as Record<string, string>,
    documents,
  };
}

export function parseLatest(value: unknown): Latest {
  const root = record(value);
  if (!/^[a-f0-9]{40}$/.test(String(root.sha))) throw new Error('Malformed render manifest');
  const samples: Record<string, number> = {};
  for (const [id, pages] of Object.entries(record(root.samples ?? {})))
    if (Number.isInteger(pages) && (pages as number) > 0) samples[id] = pages as number;
  return {
    sha: String(root.sha),
    publishedAt: typeof root.published_at_utc === 'string' ? root.published_at_utc : null,
    samples,
  };
}

export function pageName(page: number): string {
  if (!Number.isInteger(page) || page < 1) throw new Error(`Invalid page: ${page}`);
  return `page_${String(page).padStart(4, '0')}.png`;
}

export function referenceUrl(id: string, page: number, origin = CORPUS_ORIGIN): string {
  return `${origin}/${encodeURIComponent(id)}/reference/${pageName(page)}`;
}

export function renderUrl(sha: string, id: string, page: number, base = RENDER_BASE): string {
  if (!/^[a-f0-9]{40}$/.test(sha)) throw new Error('Expected a full commit SHA');
  return `${base}/${sha}/${encodeURIComponent(id)}/${pageName(page)}`;
}

export function reportUrl(sha: string, base = RENDER_BASE): string {
  if (!/^[a-f0-9]{40}$/.test(sha)) throw new Error('Expected a full commit SHA');
  return `${base}/${sha}/report.json`;
}

export function hasRender(latest: Latest | null, id: string, page: number): boolean {
  return page >= 1 && page <= (latest?.samples[id] ?? 0);
}

/** Failures first, then the lowest measured score: the pages worth opening. */
export function rank(documents: Sample[]): Sample[] {
  const key = (sample: Sample) =>
    sample.commit?.status === 'ok' ? sample.commit.penalizedSsim : -1;
  return [...documents].sort((left, right) => key(left) - key(right) || left.id.localeCompare(right.id));
}

export function score(channel: Channel | null): string {
  return channel?.status === 'ok' ? channel.penalizedSsim.toFixed(4) : '—';
}
