export const RENDER_PREFIX = 'renders';
export const SCHEMA_VERSION = 1;

function pageName(page) {
  return `page_${String(page).padStart(4, '0')}.png`;
}

export function renderKey(sha, id, page) {
  if (!/^[a-f0-9]{40}$/.test(sha)) throw new Error('Expected a full commit SHA');
  if (!/^[a-z0-9-]+$/.test(id)) throw new Error(`Invalid sample folder: ${id}`);
  if (!Number.isInteger(page) || page < 1) throw new Error(`Invalid page: ${page}`);
  return `${RENDER_PREFIX}/${sha}/${id}/${pageName(page)}`;
}

export function reportKey(sha) {
  if (!/^[a-f0-9]{40}$/.test(sha)) throw new Error('Expected a full commit SHA');
  return `${RENDER_PREFIX}/${sha}/report.json`;
}

export function commitComparison(sample, sha) {
  const matches = sample.comparisons.filter(
    (result) => result.channel === 'commit' && result.renderer_source_commit === sha
  );
  if (matches.length > 1) throw new Error(`Duplicate commit comparison: ${sample.id}`);
  const result = matches[0];
  return result && result.status !== 'failed' ? result : null;
}

export function planRenders(report) {
  if (!/^[a-f0-9]{40}$/.test(report.commit ?? '')) throw new Error('Expected a full commit SHA');
  const samples = {};
  const uploads = [];
  for (const sample of report.samples) {
    const comparison = commitComparison(sample, report.commit);
    const pages = Number.isInteger(comparison?.actual_pages) ? comparison.actual_pages : 0;
    if (pages < 1) continue;
    samples[sample.id] = pages;
    for (let page = 1; page <= pages; page += 1)
      uploads.push({
        sample: sample.id,
        page,
        name: pageName(page),
        key: renderKey(report.commit, sample.id, page),
      });
  }
  return { sha: report.commit, samples, uploads };
}

export function latestManifest(plan, publishedAt) {
  return {
    schema_version: SCHEMA_VERSION,
    sha: plan.sha,
    published_at_utc: publishedAt,
    report: reportKey(plan.sha),
    samples: plan.samples,
  };
}
