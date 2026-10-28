import { readFileSync, writeFileSync } from 'node:fs';

const ENDPOINT = 'https://betteroffice.dev/api';
const REGISTRIES = [
  { key: 'npm', label: 'npm downloads', path: 'npm-downloads' },
  { key: 'crates', label: 'crates.io downloads', path: 'crates-downloads' },
  { key: 'pypi', label: 'PyPI downloads', path: 'pypi-downloads' }
];

export function parseMonthly(message) {
  const match = /^([\d,]+)\/month$/.exec(message ?? '');
  const count = match ? Number(match[1].replaceAll(',', '')) : NaN;
  return Number.isSafeInteger(count) ? count : null;
}

export async function fetchCounts(fetchImpl = fetch) {
  const counts = {};
  for (const { key, path } of REGISTRIES) {
    try {
      const response = await fetchImpl(`${ENDPOINT}/${path}`, { signal: AbortSignal.timeout(15000) });
      const count = response.ok ? parseMonthly((await response.json()).message) : null;
      if (count !== null) counts[key] = count;
    } catch {
      // a registry that cannot be read keeps its previous snapshot
    }
  }
  return counts;
}

const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function rewriteSnapshots(readme, counts, date) {
  let updated = readme;
  for (const { key, label } of REGISTRIES) {
    if (counts[key] === undefined) continue;
    const alt = `${label}: ${counts[key].toLocaleString('en-US')}/month (snapshot ${date})`;
    const pattern = new RegExp(`alt="${escapeRegExp(label)}(?:: [^"]*)?"`);
    if (!pattern.test(updated)) throw new Error(`README has no ${label} badge`);
    updated = updated.replace(pattern, `alt="${alt}"`);
  }
  return updated;
}

/** Best effort: a failed fetch leaves the README as it was and never fails the release. */
export async function updateDownloadSnapshots(
  path = 'README.md',
  { fetchImpl = fetch, today = new Date().toISOString().slice(0, 10) } = {}
) {
  const counts = await fetchCounts(fetchImpl);
  const before = readFileSync(path, 'utf8');
  const after = rewriteSnapshots(before, counts, today);
  if (after !== before) writeFileSync(path, after);
  return counts;
}
