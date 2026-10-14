import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fetchAsset } from './asset-cache.mjs';
import { CHANNELS, TRIALS, digest } from './docx-benchmark.mjs';
import { download } from './download.mjs';
import { mapPool, validatePlan } from './plan.mjs';
import { CORPUS_ORIGIN } from './samples.mjs';

export const METHOD = 'xlsx-uncached-recalculation-v1';
export const ABS_TOLERANCE = 1e-9;
export const REL_TOLERANCE = 1e-12;

export function calculationSamples(plan) {
  return plan.samples.filter(sample => sample.format === 'xlsx' && sample.metadata.calculation?.status === 'ok');
}

export function xlsxShards(plan) {
  const samples = calculationSamples(plan);
  const shards = Array.from({ length: Math.min(4, samples.length) }, () => []);
  const weights = shards.map(() => 0);
  for (const sample of [...samples].sort((a, b) =>
    b.metadata.calculation.result_cells - a.metadata.calculation.result_cells || a.id.localeCompare(b.id))) {
    const size = sample.metadata.calculation.result_cells;
    if (!Number.isSafeInteger(size) || size < 1 || size > 250000)
      throw new Error('Invalid calculation reference size');
    const index = weights.indexOf(Math.min(...weights));
    shards[index].push(sample.id);
    weights[index] += size;
  }
  return shards;
}

function validateResult(result, total) {
  if (!result || result.total !== total || !Number.isSafeInteger(result.correct) || result.correct < 0 || result.correct > total)
    throw new Error('Invalid calculation result counts');
  if (result.status === 'failed') {
    if (result.correct !== 0 || !result.error?.trim() || 'elapsed_ms' in result)
      throw new Error('Invalid calculation failure');
  } else if (result.status !== 'ok' || !Array.isArray(result.elapsed_ms) || result.elapsed_ms.length !== TRIALS ||
             !result.elapsed_ms.every(ms => Number.isFinite(ms) && ms > 0)) {
    throw new Error('Invalid recalculation trials');
  }
}

export function calculationSummary(samples) {
  const rows = samples.filter(sample => sample.format === 'xlsx' && sample.calculations);
  for (const row of rows) {
    if (Object.keys(row.calculations).length !== CHANNELS.length)
      throw new Error('Missing calculation channels');
    const total = row.calculations.commit?.total;
    if (!Number.isSafeInteger(total) || total < 1) throw new Error('Invalid formula result total');
    for (const channel of CHANNELS) validateResult(row.calculations[channel], total);
  }
  const common = rows.filter(row => CHANNELS.every(channel => {
    const result = row.calculations[channel];
    return result.status === 'ok' && result.correct === result.total;
  }));
  return {
    workbooks: rows.length,
    common: common.length,
    channels: Object.fromEntries(CHANNELS.map(channel => [channel, {
      correct: rows.reduce((sum, row) => sum + row.calculations[channel].correct, 0),
      total: rows.reduce((sum, row) => sum + row.calculations[channel].total, 0),
      mean_ms: common.length ? common.reduce((sum, row) =>
        sum + row.calculations[channel].elapsed_ms.reduce((a, b) => a + b, 0) / TRIALS, 0) / common.length : null,
    }])),
  };
}

export function mergeXlsxBenchmarks(plan, report, inputs, planHash) {
  const shards = xlsxShards(plan);
  if (!shards.length || inputs.length !== shards.length) throw new Error('Missing XLSX calculation shards');
  const seen = new Set();
  const rows = new Map();
  let identity;
  for (const part of inputs) {
    const config = part?.benchmark;
    if (part?.schema_version !== 1 || part.plan_sha256 !== planHash ||
        !Number.isInteger(part.shard) || !shards[part.shard] || seen.has(part.shard) ||
        config?.method !== METHOD || config.trials !== TRIALS ||
        config.absolute_tolerance !== ABS_TOLERANCE || config.relative_tolerance !== REL_TOLERANCE ||
        !/^\d+(?:\.\d+){2,3}$/.test(config.libreoffice_version ?? '') ||
        config.published_version !== plan.versions.xlsx ||
        config.builds?.published?.source_sha !== plan.xlsx_published_source_sha ||
        config.builds?.commit?.source_sha !== plan.source_sha ||
        Object.keys(config.builds ?? {}).length !== 2)
      throw new Error('Invalid XLSX calculation benchmark identity');
    seen.add(part.shard);
    for (const build of Object.values(config.builds)) {
      if (!/^[a-f0-9]{64}$/.test(build.binary_sha256 ?? '') || !/^[a-f0-9]{64}$/.test(build.harness_sha256 ?? ''))
        throw new Error('Invalid XLSX native executable hashes');
    }
    if (config.builds.published.harness_sha256 !== config.builds.commit.harness_sha256 ||
        config.builds.published.rustc !== config.builds.commit.rustc || config.builds.published.profile !== config.builds.commit.profile)
      throw new Error('XLSX native builds use different harnesses or compilers');
    const serialized = JSON.stringify(config);
    if (identity && identity !== serialized) throw new Error('XLSX shards use different benchmark environments');
    identity = serialized;
    if (!Array.isArray(part.samples) || part.samples.length !== shards[part.shard].length)
      throw new Error('Missing XLSX calculation samples');
    for (const row of part.samples) {
      const sample = plan.samples.find(sample => sample.id === row?.id);
      if (!shards[part.shard].includes(row?.id) || rows.has(row.id) ||
          row.source_sha256 !== sample.metadata.source.sha256 ||
          row.reference_sha256 !== sample.metadata.calculation.reference.sha256 ||
          Object.keys(row.calculations ?? {}).length !== CHANNELS.length)
        throw new Error('Invalid XLSX calculation sample identity');
      for (const channel of CHANNELS) validateResult(row.calculations[channel], sample.metadata.calculation.result_cells);
      rows.set(row.id, row);
    }
  }
  const samples = report.samples.map(sample => rows.has(sample.id) ?
    { ...sample, calculations: rows.get(sample.id).calculations, calculation_reference_sha256: rows.get(sample.id).reference_sha256 } : sample);
  const summary = calculationSummary(samples);
  const xlsx = plan.samples.filter(sample => sample.format === 'xlsx');
  return { ...report, samples, xlsx_benchmark: { ...JSON.parse(identity), summary,
    reference_coverage: { total: xlsx.length, measured: rows.size,
      no_formulas: xlsx.filter(sample => sample.metadata.calculation?.status === 'not_applicable').map(sample => sample.id),
      missing_reference: xlsx.filter(sample => !['ok', 'not_applicable'].includes(sample.metadata.calculation?.status)).map(sample => sample.id) } } };
}

export async function stageXlsx(planPath, shard, output, cacheDir) {
  const bytes = await readFile(planPath);
  const plan = validatePlan(JSON.parse(bytes));
  const ids = xlsxShards(plan)[shard];
  if (!ids || !plan.xlsx_published_source_sha) throw new Error('Invalid XLSX shard or missing published source');
  const samples = ids.map(id => plan.samples.find(sample => sample.id === id));
  await mapPool(samples, 8, async sample => {
    const root = resolve(output, sample.id);
    await mkdir(root, { recursive: true });
    for (const [asset, name] of [[sample.metadata.source, 'source.xlsx'], [sample.metadata.calculation.reference, 'expected.json'], [sample.metadata.calculation.input, 'input.xlsx']]) {
      await writeFile(resolve(root, name), await fetchAsset(asset, sample.id, download, { cacheDir, origin: CORPUS_ORIGIN, maximum: 64 * 1024 * 1024 }));
    }
    console.log(`XLSX shard ${shard}: staged ${sample.id}`);
  });
  await writeFile(resolve(output, 'job.json'), JSON.stringify({ plan_sha256: digest(bytes), shard, samples,
    source_sha: plan.source_sha, published_source_sha: plan.xlsx_published_source_sha,
    published_version: plan.versions.xlsx, method: METHOD, trials: TRIALS,
    absolute_tolerance: ABS_TOLERANCE, relative_tolerance: REL_TOLERANCE }, null, 2) + '\n');
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { QUALITY_PLAN, QUALITY_OUTPUT, QUALITY_SHARD, QUALITY_ASSET_CACHE } = process.env;
  if (!QUALITY_PLAN || !QUALITY_OUTPUT || !/^\d+$/.test(QUALITY_SHARD ?? ''))
    throw new Error('QUALITY_PLAN, QUALITY_OUTPUT, and QUALITY_SHARD are required');
  await stageXlsx(resolve(QUALITY_PLAN), Number(QUALITY_SHARD), resolve(QUALITY_OUTPUT), QUALITY_ASSET_CACHE);
}
