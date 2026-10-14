import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fetchAsset } from './asset-cache.mjs';
import { digest } from './docx-benchmark.mjs';
import { download } from './download.mjs';
import { mapPool, validatePlan } from './plan.mjs';
import { CORPUS_ORIGIN } from './samples.mjs';

export const METHOD = 'office-single-edit-preservation-v2';
export const CHANNELS = ['published', 'commit', 'libreoffice'];
export const PARALLELISM = 4;
export const SHARD_SIZE = 16;
export const ENGINE_TIMEOUT = 180;
export const HELPER_TIMEOUT = 30;

export function roundtripShards(plan) {
  return plan.formats.flatMap(format => {
    const ids = plan.samples.filter(sample => sample.format === format).map(sample => sample.id).sort();
    return Array.from({ length: Math.ceil(ids.length / SHARD_SIZE) }, (_, shard) =>
      ({ format, shard, ids: ids.slice(shard * SHARD_SIZE, (shard + 1) * SHARD_SIZE) }));
  });
}

function validateOutcome(result) {
  const parsed = result?.parse;
  const roundtrip = result?.roundtrip;
  if (!['ok', 'failed'].includes(parsed?.status) || !['ok', 'failed'].includes(roundtrip?.status))
    throw new Error('Invalid parse or roundtrip status');
  for (const outcome of [parsed, roundtrip]) {
    if (outcome.status === 'failed' && (typeof outcome.error !== 'string' || !outcome.error.trim()))
      throw new Error('Missing parse or roundtrip failure reason');
  }
  if (roundtrip.status === 'ok' && (parsed.status !== 'ok' || result.native?.stage !== 'complete' ||
      result.native?.edit_verified !== true || roundtrip.edit_matches !== true || roundtrip.stage !== 'preserve' ||
      !Number.isInteger(roundtrip.original_parts) || roundtrip.original_parts < 1 ||
      roundtrip.identical_parts !== roundtrip.original_parts - 1 ||
      !Array.isArray(roundtrip.changed_parts) || roundtrip.changed_parts.length !== 1 ||
      !['added_parts', 'removed_parts', 'unrelated_changed_parts'].every(key =>
        Array.isArray(roundtrip[key]) && roundtrip[key].length === 0) ||
      !/^[a-f0-9]{64}$/.test(result.native.output_sha256 ?? '')))
    throw new Error('Roundtrip success lacks preservation or reopen evidence');
}

export function roundtripSummary(samples, format) {
  const rows = samples.filter(sample => sample.format === format);
  const channels = Object.fromEntries(CHANNELS.map(channel => [channel, { parsed: 0, preserved: 0, total: rows.length }]));
  for (const sample of rows) {
    if (!sample.roundtrip || Object.keys(sample.roundtrip).length !== CHANNELS.length)
      throw new Error('Missing roundtrip channels');
    for (const channel of CHANNELS) {
      const result = sample.roundtrip[channel];
      validateOutcome(result);
      if (result.parse.status === 'ok') channels[channel].parsed++;
      if (result.roundtrip.status === 'ok') channels[channel].preserved++;
    }
  }
  return channels;
}

export function mergeRoundtrips(plan, report, parts, planHash) {
  const shards = roundtripShards(plan);
  if (!Array.isArray(parts) || parts.length !== shards.length)
    throw new Error('Missing roundtrip format reports');
  const benchmarks = {}, results = new Map(), seen = new Set();
  let checker, office;
  for (const part of parts) {
    const format = part?.format;
    const config = part?.benchmark;
    const shard = shards.find(shard => shard.format === format && shard.shard === part?.shard);
    const key = `${format}/${part?.shard}`;
    const expected = plan.samples.filter(sample => shard?.ids.includes(sample.id));
    if (!shard || seen.has(key) || part.schema_version !== 1 ||
        part.plan_sha256 !== planHash || config?.method !== METHOD || config.timeout_seconds !== ENGINE_TIMEOUT ||
        config.helper_timeout_seconds !== HELPER_TIMEOUT || config.parallelism !== PARALLELISM ||
        config.shard_size !== SHARD_SIZE || !/^\d+(?:\.\d+){2,3}$/.test(config.libreoffice_version ?? '') ||
        typeof config.libreoffice_build !== 'string' || !config.libreoffice_build.trim() ||
        !/^[a-f0-9]{64}$/.test(config.libreoffice_host_sha256 ?? '') ||
        config.published_version !== plan.versions[format] ||
        !/^[a-f0-9]{40}$/.test(plan[`${format}_published_source_sha`] ?? '') ||
        config.builds?.published?.source_sha !== plan[`${format}_published_source_sha`] ||
        config.builds?.commit?.source_sha !== plan.source_sha ||
        !/^[a-f0-9]{64}$/.test(config.checker_sha256 ?? '') ||
        !Array.isArray(part.samples) || part.samples.length !== expected.length)
      throw new Error('Invalid roundtrip report identity');
    if (checker && checker !== config.checker_sha256) throw new Error('Roundtrip checkers differ across formats');
    checker = config.checker_sha256;
    const officeIdentity = JSON.stringify([config.libreoffice_version, config.libreoffice_build, config.libreoffice_host_sha256]);
    if (office && office !== officeIdentity) throw new Error('Roundtrip LibreOffice identities differ');
    office = officeIdentity;
    if (Object.keys(config.builds).length !== 2 ||
        !['harness_sha256', 'rustc', 'profile'].every(key =>
          typeof config.builds.published[key] === 'string' && config.builds.published[key].length &&
          config.builds.published[key] === config.builds.commit[key]))
      throw new Error('Roundtrip builds use different hosts or compilers');
    for (const build of Object.values(config.builds)) {
      if (!/^[a-f0-9]{64}$/.test(build.harness_sha256 ?? '') || !/^[a-f0-9]{64}$/.test(build.binary_sha256 ?? ''))
        throw new Error('Invalid roundtrip build hashes');
    }
    if (benchmarks[format] && JSON.stringify(benchmarks[format]) !== JSON.stringify(config))
      throw new Error('Roundtrip shard environments differ');
    for (const measured of [report[`${format}_benchmark`], format === 'xlsx' && report.xlsx_fidelity_benchmark]) {
      if (measured && (measured.libreoffice_version !== config.libreoffice_version || measured.libreoffice_build !== config.libreoffice_build))
        throw new Error('Roundtrip and fidelity use different LibreOffice builds');
    }
    seen.add(key);
    for (const row of part.samples) {
      const sample = expected.find(sample => sample.id === row?.id);
      if (!sample || results.has(row.id) || row.source_sha256 !== sample.metadata.source.sha256 ||
          !row.channels || Object.keys(row.channels).length !== CHANNELS.length)
        throw new Error('Missing, duplicated or mismatched roundtrip sample');
      for (const channel of CHANNELS) {
        const result = row.channels[channel];
        validateOutcome(result);
        if (result.parse.status === 'ok' && (result.native?.parse !== 'ok' || result.native.source_sha256 !== row.source_sha256))
          throw new Error('Parse success lacks source identity');
        if (result.roundtrip.status === 'ok' && (row.probe?.status !== 'ok' ||
            typeof row.probe.part !== 'string' || result.roundtrip.changed_parts[0] !== row.probe.part ||
            typeof row.probe.old !== 'string' || typeof row.probe.new !== 'string' || row.probe.old === row.probe.new))
          throw new Error('Roundtrip success lacks a content edit');
      }
      results.set(row.id, row);
    }
    benchmarks[format] = { ...config };
  }
  const samples = report.samples.map(sample => {
    const row = results.get(sample.id);
    if (!row || sample.roundtrip) throw new Error('Missing or duplicate roundtrip sample');
    return { ...sample, roundtrip: row.channels, roundtrip_probe: row.probe };
  });
  for (const format of plan.formats) benchmarks[format].summary = roundtripSummary(samples, format);
  return { ...report, roundtrip_benchmark: benchmarks, samples };
}

export async function stageRoundtrip(planPath, output, format, shardIndex, cacheDir) {
  const bytes = await readFile(planPath);
  const plan = validatePlan(JSON.parse(bytes));
  if (!plan.formats.includes(format) || !plan[`${format}_published_source_sha`])
    throw new Error('Missing roundtrip format or published source');
  const shard = roundtripShards(plan).find(shard => shard.format === format && shard.shard === shardIndex);
  if (!shard) throw new Error('Invalid roundtrip shard');
  const samples = plan.samples.filter(sample => shard.ids.includes(sample.id));
  let completed = 0;
  await mapPool(samples, 8, async sample => {
    const directory = resolve(output, sample.id);
    await mkdir(directory, { recursive: true });
    await writeFile(resolve(directory, `source.${format}`), await fetchAsset(sample.metadata.source, sample.id,
      download, { cacheDir, origin: CORPUS_ORIGIN, maximum: 128 * 1024 * 1024 }));
    if (++completed % 25 === 0 || completed === samples.length)
      console.log(`${format.toUpperCase()} roundtrip: staged ${completed}/${samples.length} verified files`);
  });
  await writeFile(resolve(output, 'job.json'), JSON.stringify({ plan_sha256: digest(bytes), samples, format, shard: shardIndex, method: METHOD,
    parallelism: PARALLELISM, shard_size: SHARD_SIZE, timeout_seconds: ENGINE_TIMEOUT, helper_timeout_seconds: HELPER_TIMEOUT,
    source_sha: plan.source_sha, published_source_sha: plan[`${format}_published_source_sha`],
    published_version: plan.versions[format] }, null, 2) + '\n');
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { QUALITY_PLAN, QUALITY_OUTPUT, QUALITY_FORMAT, QUALITY_SHARD, QUALITY_ASSET_CACHE } = process.env;
  if (!QUALITY_PLAN || !QUALITY_OUTPUT || !QUALITY_FORMAT || !/^\d+$/.test(QUALITY_SHARD ?? ''))
    throw new Error('QUALITY_PLAN, QUALITY_OUTPUT, QUALITY_FORMAT and QUALITY_SHARD are required');
  await stageRoundtrip(resolve(QUALITY_PLAN), resolve(QUALITY_OUTPUT), QUALITY_FORMAT, Number(QUALITY_SHARD), QUALITY_ASSET_CACHE);
}
