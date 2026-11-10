import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fetchAsset } from './asset-cache.mjs';
import { CHANNELS, digest } from './docx-benchmark.mjs';
import { download } from './download.mjs';
import { mapPool, validatePlan } from './plan.mjs';
import { corpusOrigin } from './samples.mjs';
import { calculationSamples } from './xlsx-benchmark.mjs';

export const METHOD = 'sequential-paired-sample-v1';
export const SAMPLE_SIZE = 200;
export const ROUNDS = 3;
export const SEED = 'betteroffice-timing-v1';
export const BOOTSTRAP = 2000;
export const TIMING_FORMATS = ['docx', 'pptx', 'xlsx'];

function fail(message) {
  throw new Error(`Invalid timing benchmark: ${message}`);
}

function eligible(plan, format) {
  if (format === 'xlsx') return calculationSamples(plan);
  return plan.samples.filter((sample) => sample.format === format);
}

export function timingSamples(plan, format) {
  const rank = (id) => createHash('sha256').update(`${SEED}\0${format}\0${id}`).digest('hex');
  return eligible(plan, format)
    .map((sample) => ({ sample, rank: rank(sample.id) }))
    .sort((a, b) => a.rank.localeCompare(b.rank) || a.sample.id.localeCompare(b.sample.id))
    .slice(0, SAMPLE_SIZE)
    .map(({ sample }) => sample);
}

export function timingFormats(plan) {
  return TIMING_FORMATS.filter((format) => plan.formats.includes(format) && timingSamples(plan, format).length);
}

function publishedSource(plan, format) {
  return plan[`${format}_published_source_sha`];
}

export async function stageTiming(planPath, format, output, cacheDir) {
  const bytes = await readFile(planPath);
  const plan = validatePlan(JSON.parse(bytes));
  if (!timingFormats(plan).includes(format) || !publishedSource(plan, format))
    fail(`no ${format} sample or published source`);
  const samples = timingSamples(plan, format);
  const tasks = [];
  for (const sample of samples) {
    const root = resolve(output, sample.id);
    if (format === 'xlsx') {
      await mkdir(root, { recursive: true });
      for (const [asset, name] of [[sample.metadata.source, 'source.xlsx'],
        [sample.metadata.calculation.reference, 'expected.json'], [sample.metadata.calculation.input, 'input.xlsx']])
        tasks.push({ sample, asset, path: resolve(root, name), maximum: 64 * 1024 * 1024 });
    } else {
      await mkdir(resolve(root, 'reference'), { recursive: true });
      tasks.push({ sample, asset: sample.metadata.source, path: resolve(root, `source.${format}`), maximum: 128 * 1024 * 1024 });
      tasks.push({ sample, asset: sample.metadata.reference_pages[0], path: resolve(root, 'reference', 'page_0001.png'), maximum: 32 * 1024 * 1024 });
    }
  }
  let completed = 0;
  await mapPool(tasks, 8, async ({ sample, asset, path, maximum }) => {
    await writeFile(path, await fetchAsset(asset, sample.id, download, { cacheDir, origin: corpusOrigin(), maximum }));
    if (++completed % 50 === 0 || completed === tasks.length)
      console.log(`${format.toUpperCase()} timing: staged ${completed}/${tasks.length} verified assets`);
  });
  await writeFile(resolve(output, 'job.json'), JSON.stringify({
    plan_sha256: digest(bytes), format, samples, method: METHOD, rounds: ROUNDS, seed: SEED, sample_size: SAMPLE_SIZE,
    source_sha: plan.source_sha, published_source_sha: publishedSource(plan, format),
    published_version: plan.versions[format],
  }, null, 2) + '\n');
}

function validateChannel(result) {
  if (result?.status === 'failed') {
    if (typeof result.error !== 'string' || !result.error.trim() || 'elapsed_ms' in result) fail('invalid failure');
    return;
  }
  if (result?.status !== 'ok' || !Array.isArray(result.elapsed_ms) || result.elapsed_ms.length !== ROUNDS ||
      !result.elapsed_ms.every((value) => Number.isFinite(value) && value > 0))
    fail('invalid rounds');
}

const mean = (values) => values.reduce((sum, value) => sum + value, 0) / values.length;
const median = (values) => {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = sorted.length >> 1;
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
};

function random(seed) {
  let state = Number.parseInt(createHash('sha256').update(seed).digest('hex').slice(0, 8), 16);
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

export function pairedRatio(ours, theirs, seed) {
  if (ours.length !== theirs.length || !ours.length) return null;
  const logs = ours.map((value, index) => Math.log(value / theirs[index]));
  const next = random(seed);
  const resampled = Array.from({ length: BOOTSTRAP }, () =>
    mean(Array.from(logs, () => logs[Math.floor(next() * logs.length)])));
  resampled.sort((a, b) => a - b);
  const round = (value) => Number(Math.exp(value).toFixed(4));
  return {
    geometric_mean: round(mean(logs)),
    ci95: [round(resampled[Math.floor(0.025 * (BOOTSTRAP - 1))]), round(resampled[Math.ceil(0.975 * (BOOTSTRAP - 1))])],
  };
}

export function summarizeTiming(format, samples) {
  const common = samples.filter((sample) => CHANNELS.every((channel) => sample.timings[channel].status === 'ok'));
  const perDocument = (channel, statistic) => common.map((sample) => statistic(sample.timings[channel].elapsed_ms));
  const channels = Object.fromEntries(CHANNELS.map((channel) => {
    const rounds = Array.from({ length: ROUNDS }, (_, round) =>
      common.length ? mean(common.map((sample) => sample.timings[channel].elapsed_ms[round])) : null);
    const finite = rounds.filter((value) => value !== null);
    return [channel, {
      successful: samples.filter((sample) => sample.timings[channel].status === 'ok').length,
      mean_ms: common.length ? mean(perDocument(channel, mean)) : null,
      median_ms: common.length ? median(perDocument(channel, median)) : null,
      round_means_ms: rounds,
      round_spread: finite.length ? (Math.max(...finite) - Math.min(...finite)) / mean(finite) : null,
    }];
  }));
  const ratio = (channel) => pairedRatio(perDocument(channel, median), perDocument('libreoffice', median), `${SEED}\0${format}\0${channel}`);
  return { format, total: samples.length, common: common.length, channels,
    ratio_to_libreoffice: { published: ratio('published'), commit: ratio('commit') } };
}

export function mergeTiming(plan, report, parts, planHash) {
  const formats = timingFormats(plan);
  if (!Array.isArray(parts) || parts.length !== formats.length) fail('missing or unexpected format reports');
  const result = {};
  for (const part of parts) {
    const format = part?.format;
    const config = part?.benchmark;
    if (!formats.includes(format) || result[format] || part.schema_version !== 1 || part.plan_sha256 !== planHash ||
        config?.method !== METHOD || config.rounds !== ROUNDS || config.seed !== SEED ||
        config.sample_size !== SAMPLE_SIZE || config.published_version !== plan.versions[format] ||
        !/^\d+(?:\.\d+){2,3}$/.test(config.libreoffice_version ?? '') ||
        Object.keys(config.builds ?? {}).length !== 2 ||
        config.builds.published?.source_sha !== publishedSource(plan, format) ||
        config.builds.commit?.source_sha !== plan.source_sha ||
        !Number.isInteger(part.runner?.processors) || part.runner.processors < 1)
      fail(`invalid ${format} identity`);
    const office = format === 'xlsx' ? report.xlsx_benchmark : report[`${format}_benchmark`];
    if (office && (office.libreoffice_version !== config.libreoffice_version ||
        office.libreoffice_build !== config.libreoffice_build))
      fail(`${format} LibreOffice version differs from the report`);
    if (office && ['published', 'commit'].some((channel) =>
      JSON.stringify(office.builds?.[channel]) !== JSON.stringify(config.builds[channel])))
      fail(`${format} native builds differ from the report`);
    const expected = timingSamples(plan, format);
    if (!Array.isArray(part.samples) || part.samples.length !== expected.length) fail(`missing ${format} samples`);
    part.samples.forEach((row, index) => {
      const sample = expected[index];
      if (row?.id !== sample.id || row.source_sha256 !== sample.metadata.source.sha256 ||
          Object.keys(row.timings ?? {}).length !== CHANNELS.length)
        fail(`unexpected ${format} sample`);
      for (const channel of CHANNELS) validateChannel(row.timings[channel]);
    });
    result[format] = { ...config, runner: part.runner, summary: summarizeTiming(format, part.samples), samples: part.samples };
  }
  return { ...report, timing_benchmark: result };
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { QUALITY_PLAN, QUALITY_OUTPUT, QUALITY_FORMAT, QUALITY_ASSET_CACHE } = process.env;
  if (!QUALITY_PLAN || !QUALITY_OUTPUT || !TIMING_FORMATS.includes(QUALITY_FORMAT))
    throw new Error('QUALITY_PLAN, QUALITY_OUTPUT, and QUALITY_FORMAT are required');
  await stageTiming(resolve(QUALITY_PLAN), QUALITY_FORMAT, resolve(QUALITY_OUTPUT), QUALITY_ASSET_CACHE);
}
