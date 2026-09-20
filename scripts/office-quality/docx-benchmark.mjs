import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fetchAsset } from './asset-cache.mjs';
import { download } from './download.mjs';
import { mapPool, validatePlan } from './plan.mjs';
import { validateComparison } from './results.mjs';
import { CORPUS_ORIGIN } from './samples.mjs';

export const CHANNELS = ['published', 'commit', 'libreoffice'];
export const TRIALS = 5;
export const METHOD = 'docx-cli-page1-v1';
export const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');

export function docxShards(plan) {
  const samples = plan.samples.filter((sample) => sample.format === 'docx');
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

function fail(message) {
  throw new Error(`Invalid DOCX benchmark: ${message}`);
}

function validateTiming(timing) {
  if (timing?.status === 'failed') {
    if (typeof timing.error !== 'string' || !timing.error.trim() || 'elapsed_ms' in timing)
      fail('invalid timing failure');
    return;
  }
  if (timing?.status !== 'ok' || !Array.isArray(timing.elapsed_ms) ||
      timing.elapsed_ms.length !== TRIALS ||
      !timing.elapsed_ms.every((value) => Number.isFinite(value) && value > 0))
    fail('invalid timing trials');
}

export function timingSummary(samples) {
  const rows = samples.filter((sample) => sample.format === 'docx');
  for (const row of rows) {
    if (!row.native_timings || Object.keys(row.native_timings).length !== CHANNELS.length)
      fail('missing timing channels');
    for (const channel of CHANNELS) validateTiming(row.native_timings[channel]);
  }
  const common = rows.filter((row) => CHANNELS.every((channel) => row.native_timings[channel].status === 'ok'));
  return {
    total: rows.length,
    common: common.length,
    channels: Object.fromEntries(CHANNELS.map((channel) => [channel, {
      successful: rows.filter((row) => row.native_timings[channel].status === 'ok').length,
      mean_ms: common.length ? common.reduce((sum, row) =>
        sum + row.native_timings[channel].elapsed_ms.reduce((a, b) => a + b, 0) / TRIALS, 0) / common.length : null,
    }])),
  };
}

export function mergeDocxBenchmarks(plan, report, inputs, planHash) {
  const shards = docxShards(plan);
  if (!/^[a-f0-9]{40}$/.test(plan.docx_published_source_sha ?? '')) fail('missing published source');
  if (!Array.isArray(inputs) || inputs.length !== shards.length || !shards.length)
    fail('missing or unexpected shard reports');
  const seen = new Set();
  const results = new Map();
  let identity;
  for (const part of inputs) {
    if (part?.schema_version !== 1 || part.plan_sha256 !== planHash ||
        !Number.isInteger(part.shard) || !shards[part.shard] || seen.has(part.shard))
      fail('duplicate shard or mismatched plan');
    seen.add(part.shard);
    const config = part.benchmark;
    if (config?.method !== METHOD || config.trials !== TRIALS || config.dpi !== 96 ||
        !/^\d+(?:\.\d+){2,3}$/.test(config.libreoffice_version ?? '') ||
        !/^[a-f0-9]{64}$/.test(config.fonts_sha256 ?? '') ||
        config.published_version !== plan.versions.docx ||
        !config.builds || Object.keys(config.builds).length !== 2 ||
        config.builds?.published?.source_sha !== plan.docx_published_source_sha ||
        config.builds?.commit?.source_sha !== plan.source_sha)
      fail('invalid executable or measurement identity');
    for (const build of Object.values(config.builds)) {
      if (!/^[a-f0-9]{64}$/.test(build.binary_sha256 ?? '') ||
          !/^[a-f0-9]{64}$/.test(build.harness_sha256 ?? '')) fail('invalid build hashes');
    }
    if (config.builds.published.harness_sha256 !== config.builds.commit.harness_sha256 ||
        config.builds.published.rustc !== config.builds.commit.rustc ||
        config.builds.published.profile !== config.builds.commit.profile)
      fail('native builds use different harnesses or compilers');
    const serialized = JSON.stringify(config);
    if (identity && identity !== serialized) fail('shards use different benchmark environments');
    identity = serialized;
    if (!Array.isArray(part.samples) || part.samples.length !== shards[part.shard].length)
      fail('missing samples');
    for (const row of part.samples) {
      if (!shards[part.shard].includes(row?.id) || results.has(row.id))
        fail('duplicate, unexpected, or misplaced sample');
      const sample = plan.samples.find((sample) => sample.id === row.id);
      if (row.source_sha256 !== sample.metadata.source.sha256) fail('source hash mismatch');
      const result = row.libreoffice;
      if (result?.channel !== 'libreoffice' || result.version !== config.libreoffice_version)
        fail('LibreOffice comparison identity mismatch');
      if (result.status === 'failed') {
        if (!['capture', 'compare'].includes(result.stage) || !result.error?.trim() || 'penalized_ssim' in result)
          fail('invalid LibreOffice failure');
      } else {
        if (result.status !== 'ok') fail('invalid LibreOffice status');
        validateComparison(result);
        if (result.reference.sha256 !== row.source_sha256 ||
            result.reference_pages !== sample.metadata.reference.pages ||
            !Number.isInteger(result.actual_pages) || result.actual_pages < 1)
          fail('LibreOffice reference mismatch');
      }
      results.set(row.id, row);
    }
  }
  const samples = report.samples.map((sample) => {
    if (sample.format !== 'docx') return sample;
    const row = results.get(sample.id);
    if (!row) fail('missing sample');
    return { ...sample, comparisons: [...sample.comparisons, row.libreoffice], native_timings: row.timings };
  });
  timingSummary(samples);
  return { ...report, samples, docx_benchmark: JSON.parse(identity),
    docx_runners: inputs.map((part) => ({ shard: part.shard, ...part.runner })) };
}

export async function stageDocx(planPath, shard, output, cacheDir) {
  const bytes = await readFile(planPath);
  const plan = validatePlan(JSON.parse(bytes));
  const ids = docxShards(plan)[shard];
  if (!ids || !plan.docx_published_source_sha) fail('invalid shard or missing published source');
  const samples = ids.map((id) => plan.samples.find((sample) => sample.id === id));
  const tasks = [];
  for (const sample of samples) {
    const root = resolve(output, sample.id);
    await mkdir(resolve(root, 'reference'), { recursive: true });
    await writeFile(resolve(root, 'reference/result.json'), JSON.stringify(sample.metadata.reference));
    tasks.push({ sample, asset: sample.metadata.source, path: resolve(root, 'source.docx'), maximum: 128 * 1024 * 1024 });
    sample.metadata.reference_pages.forEach((asset, index) => tasks.push({
      sample, asset, path: resolve(root, 'reference', `page_${String(index + 1).padStart(4, '0')}.png`), maximum: 32 * 1024 * 1024,
    }));
  }
  let completed = 0;
  await mapPool(tasks, 8, async ({ sample, asset, path, maximum }) => {
    await writeFile(path, await fetchAsset(asset, sample.id, download, { cacheDir, origin: CORPUS_ORIGIN, maximum }));
    if (++completed % 50 === 0 || completed === tasks.length)
      console.log(`DOCX shard ${shard}: staged ${completed}/${tasks.length} verified assets`);
  });
  await writeFile(resolve(output, 'job.json'), JSON.stringify({
    plan_sha256: digest(bytes), shard, samples, source_sha: plan.source_sha,
    published_source_sha: plan.docx_published_source_sha, method: METHOD, trials: TRIALS,
    published_version: plan.versions.docx,
  }, null, 2) + '\n');
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { QUALITY_PLAN, QUALITY_OUTPUT, QUALITY_SHARD, QUALITY_ASSET_CACHE } = process.env;
  if (!QUALITY_PLAN || !QUALITY_OUTPUT || !/^\d+$/.test(QUALITY_SHARD ?? ''))
    throw new Error('QUALITY_PLAN, QUALITY_OUTPUT, and QUALITY_SHARD are required');
  await stageDocx(resolve(QUALITY_PLAN), Number(QUALITY_SHARD), resolve(QUALITY_OUTPUT), QUALITY_ASSET_CACHE);
}
