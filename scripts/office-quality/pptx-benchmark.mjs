import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fetchAsset } from './asset-cache.mjs';
import { digest, timingSummary, TRIALS } from './docx-benchmark.mjs';
import { download } from './download.mjs';
import { mapPool, validatePlan } from './plan.mjs';
import { validateComparison } from './results.mjs';
import { CORPUS_ORIGIN } from './samples.mjs';

export const METHOD = 'pptx-cli-slide1-v1';

export function mergePptxBenchmark(plan, report, part, planHash) {
  const expected = plan.samples.filter(sample => sample.format === 'pptx');
  const benchmark = part?.benchmark;
  if (part?.schema_version !== 1 || part.plan_sha256 !== planHash ||
      benchmark?.method !== METHOD || benchmark.dpi !== 96 || benchmark.trials !== TRIALS ||
      benchmark.published_version !== plan.versions.pptx ||
      !/^[a-f0-9]{40}$/.test(plan.pptx_published_source_sha ?? '') ||
      benchmark.builds?.published?.source_sha !== plan.pptx_published_source_sha ||
      benchmark.builds?.commit?.source_sha !== plan.source_sha ||
      !/^\d+(?:\.\d+){2,3}$/.test(benchmark.libreoffice_version ?? '') ||
      !/^[a-f0-9]{64}$/.test(benchmark.fonts_sha256 ?? '') ||
      !Array.isArray(part.samples) || part.samples.length !== expected.length)
    throw new Error('Invalid PPTX LibreOffice benchmark identity');
  if (Object.keys(benchmark.builds).length !== 2 ||
      benchmark.builds.published.harness_sha256 !== benchmark.builds.commit.harness_sha256 ||
      benchmark.builds.published.rustc !== benchmark.builds.commit.rustc ||
      benchmark.builds.published.profile !== benchmark.builds.commit.profile)
    throw new Error('PPTX native builds use different harnesses or compilers');
  for (const build of Object.values(benchmark.builds)) {
    if (!/^[a-f0-9]{64}$/.test(build.binary_sha256 ?? '') || !/^[a-f0-9]{64}$/.test(build.harness_sha256 ?? ''))
      throw new Error('Invalid PPTX native build hashes');
  }
  const results = new Map();
  for (const row of part.samples) {
    const sample = expected.find(sample => sample.id === row?.id);
    const result = row?.libreoffice;
    if (!sample || results.has(row.id) || row.source_sha256 !== sample.metadata.source.sha256 ||
        result?.channel !== 'libreoffice' || result.version !== benchmark.libreoffice_version)
      throw new Error('Invalid PPTX LibreOffice sample identity');
    if (result.status === 'failed') {
      if (!['capture', 'compare'].includes(result.stage) || !result.error?.trim() || 'penalized_ssim' in result)
        throw new Error('Invalid PPTX LibreOffice failure');
    } else {
      if (result.status !== 'ok') throw new Error('Invalid PPTX LibreOffice status');
      validateComparison(result);
      if (result.reference.sha256 !== row.source_sha256 || result.reference_pages !== sample.metadata.reference.pages)
        throw new Error('Mismatched PPTX LibreOffice reference');
    }
    results.set(row.id, row);
  }
  const samples = report.samples.map(sample => sample.format === 'pptx' ? {
    ...sample, comparisons: [...sample.comparisons, results.get(sample.id).libreoffice],
    native_timings: results.get(sample.id).timings,
  } : sample);
  timingSummary(samples, 'pptx');
  return { ...report, pptx_benchmark: benchmark, samples };
}

export async function stagePptx(planPath, output, cacheDir) {
  const bytes = await readFile(planPath);
  const plan = validatePlan(JSON.parse(bytes));
  const samples = plan.samples.filter(sample => sample.format === 'pptx');
  if (!samples.length || !plan.pptx_published_source_sha) throw new Error('Missing PPTX samples or published source');
  const tasks = [];
  for (const sample of samples) {
    const root = resolve(output, sample.id);
    await mkdir(resolve(root, 'reference'), { recursive: true });
    await writeFile(resolve(root, 'reference/result.json'), JSON.stringify(sample.metadata.reference));
    tasks.push({ sample, asset: sample.metadata.source, path: resolve(root, 'source.pptx'), maximum: 128 * 1024 * 1024 });
    sample.metadata.reference_pages.forEach((asset, index) => tasks.push({ sample, asset,
      path: resolve(root, 'reference', `page_${String(index + 1).padStart(4, '0')}.png`), maximum: 32 * 1024 * 1024 }));
  }
  let completed = 0;
  await mapPool(tasks, 8, async ({ sample, asset, path, maximum }) => {
    await writeFile(path, await fetchAsset(asset, sample.id, download, { cacheDir, origin: CORPUS_ORIGIN, maximum }));
    if (++completed % 50 === 0 || completed === tasks.length)
      console.log(`PPTX LibreOffice: staged ${completed}/${tasks.length} verified assets`);
  });
  await writeFile(resolve(output, 'job.json'), JSON.stringify({ plan_sha256: digest(bytes), samples, method: METHOD, trials: TRIALS,
    source_sha: plan.source_sha, published_source_sha: plan.pptx_published_source_sha, published_version: plan.versions.pptx }, null, 2) + '\n');
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { QUALITY_PLAN, QUALITY_OUTPUT, QUALITY_ASSET_CACHE } = process.env;
  if (!QUALITY_PLAN || !QUALITY_OUTPUT) throw new Error('QUALITY_PLAN and QUALITY_OUTPUT are required');
  await stagePptx(resolve(QUALITY_PLAN), resolve(QUALITY_OUTPUT), QUALITY_ASSET_CACHE);
}
