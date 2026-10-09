import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fetchAsset } from './asset-cache.mjs';
import { digest } from './docx-benchmark.mjs';
import { download } from './download.mjs';
import { mapPool, validatePlan } from './plan.mjs';
import { validateComparison } from './results.mjs';
import { CORPUS_ORIGIN } from './samples.mjs';

export const METHOD = 'xlsx-libreoffice-print-ranges-v1';

export function xlsxFidelityShards(plan) {
  const samples = plan.samples.filter(sample => sample.format === 'xlsx');
  const shards = Array.from({ length: Math.min(4, samples.length) }, () => []);
  const weights = shards.map(() => 0);
  for (const sample of [...samples].sort((a, b) =>
    b.metadata.reference.pages - a.metadata.reference.pages || a.id.localeCompare(b.id))) {
    const index = weights.indexOf(Math.min(...weights));
    shards[index].push(sample.id);
    weights[index] += sample.metadata.reference.pages;
  }
  return shards;
}

export function mergeXlsxFidelity(plan, report, parts, planHash) {
  const shards = xlsxFidelityShards(plan);
  if (!shards.length || !Array.isArray(parts) || parts.length !== shards.length)
    throw new Error('Missing XLSX fidelity shards');
  const results = new Map();
  const seen = new Set();
  let identity;
  for (const part of parts) {
    const config = part?.benchmark;
    if (part?.schema_version !== 1 || part.plan_sha256 !== planHash ||
        !Number.isInteger(part.shard) || !shards[part.shard] || seen.has(part.shard) ||
        config?.method !== METHOD || config.dpi !== 150 || config.source_sha !== plan.source_sha ||
        !/^\d+(?:\.\d+){2,3}$/.test(config.libreoffice_version ?? '') ||
        !/^[a-f0-9]{64}$/.test(config.fonts_sha256 ?? '') ||
        !Array.isArray(part.samples) || part.samples.length !== shards[part.shard].length)
      throw new Error('Invalid XLSX fidelity identity');
    const serialized = JSON.stringify(config);
    if (identity && serialized !== identity) throw new Error('XLSX fidelity shard environments differ');
    identity = serialized;
    seen.add(part.shard);
    for (const row of part.samples) {
      const sample = plan.samples.find(sample => sample.id === row?.id);
      const result = row?.libreoffice;
      if (!shards[part.shard].includes(row?.id) || results.has(row.id) ||
          row.source_sha256 !== sample?.metadata.source.sha256 ||
          result?.channel !== 'libreoffice' || result.version !== config.libreoffice_version)
        throw new Error('Invalid XLSX fidelity sample');
      if (result.status === 'failed') {
        if (!['capture', 'compare'].includes(result.stage) || typeof result.error !== 'string' ||
            !result.error.trim() || 'penalized_ssim' in result)
          throw new Error('Invalid XLSX fidelity failure');
      } else {
        if (result.status !== 'ok') throw new Error('Invalid XLSX fidelity status');
        validateComparison(result);
        if (result.reference.sha256 !== row.source_sha256 ||
            result.reference_pages !== sample.metadata.reference.pages ||
            result.actual_pages !== result.reference_pages)
          throw new Error('XLSX fidelity ranges do not match the reference');
      }
      results.set(row.id, result);
    }
  }
  const config = parts[0].benchmark;
  if (report.xlsx_benchmark && (report.xlsx_benchmark.libreoffice_version !== config.libreoffice_version ||
      report.xlsx_benchmark.libreoffice_build !== config.libreoffice_build))
    throw new Error('XLSX fidelity and recalculation use different LibreOffice builds');
  return { ...report, xlsx_fidelity_benchmark: config, samples: report.samples.map(sample => {
    if (sample.format !== 'xlsx') return sample;
    if (!results.has(sample.id) || sample.comparisons.some(result => result.channel === 'libreoffice'))
      throw new Error('Missing or duplicate XLSX LibreOffice comparison');
    return { ...sample, comparisons: [...sample.comparisons, results.get(sample.id)] };
  }) };
}

export async function stageXlsxFidelity(planPath, output, shard, cacheDir) {
  const bytes = await readFile(planPath);
  const plan = validatePlan(JSON.parse(bytes));
  const ids = xlsxFidelityShards(plan)[shard];
  if (!Number.isInteger(shard) || !ids) throw new Error('Invalid XLSX fidelity shard');
  const samples = plan.samples.filter(sample => ids.includes(sample.id));
  const tasks = [];
  for (const sample of samples) {
    const root = resolve(output, sample.id);
    await mkdir(resolve(root, 'reference'), { recursive: true });
    await writeFile(resolve(root, 'reference/result.json'), JSON.stringify(sample.metadata.reference));
    tasks.push({ sample, asset: sample.metadata.source, path: resolve(root, 'source.xlsx'), maximum: 128 * 1024 * 1024 });
    sample.metadata.reference_pages.forEach((asset, index) => tasks.push({ sample, asset,
      path: resolve(root, 'reference', `page_${String(index + 1).padStart(4, '0')}.png`), maximum: 32 * 1024 * 1024 }));
  }
  let completed = 0;
  await mapPool(tasks, 8, async ({ sample, asset, path, maximum }) => {
    await writeFile(path, await fetchAsset(asset, sample.id, download, { cacheDir, origin: CORPUS_ORIGIN, maximum }));
    if (++completed % 50 === 0 || completed === tasks.length)
      console.log(`XLSX LibreOffice: staged ${completed}/${tasks.length} verified assets`);
  });
  await writeFile(resolve(output, 'job.json'), JSON.stringify({ plan_sha256: digest(bytes), samples,
    shard, method: METHOD, source_sha: plan.source_sha }, null, 2) + '\n');
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { QUALITY_PLAN, QUALITY_OUTPUT, QUALITY_SHARD, QUALITY_ASSET_CACHE } = process.env;
  if (!QUALITY_PLAN || !QUALITY_OUTPUT || !/^\d+$/.test(QUALITY_SHARD ?? ''))
    throw new Error('QUALITY_PLAN, QUALITY_OUTPUT and QUALITY_SHARD are required');
  await stageXlsxFidelity(resolve(QUALITY_PLAN), resolve(QUALITY_OUTPUT), Number(QUALITY_SHARD), QUALITY_ASSET_CACHE);
}
