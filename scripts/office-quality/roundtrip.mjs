import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fetchAsset } from './asset-cache.mjs';
import { digest } from './docx-benchmark.mjs';
import { download } from './download.mjs';
import { mapPool, validatePlan } from './plan.mjs';
import { CORPUS_ORIGIN } from './samples.mjs';

export const METHOD = 'office-single-edit-preservation-v1';
const CHANNELS = ['published', 'commit'];

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
  if (!Array.isArray(parts) || parts.length !== plan.formats.length)
    throw new Error('Missing roundtrip format reports');
  const benchmarks = {}, results = new Map();
  let checker;
  for (const part of parts) {
    const format = part?.format;
    const config = part?.benchmark;
    const expected = plan.samples.filter(sample => sample.format === format);
    if (!plan.formats.includes(format) || benchmarks[format] || part.schema_version !== 1 ||
        part.plan_sha256 !== planHash || config?.method !== METHOD || config.timeout_seconds !== 180 ||
        config.published_version !== plan.versions[format] ||
        !/^[a-f0-9]{40}$/.test(plan[`${format}_published_source_sha`] ?? '') ||
        config.builds?.published?.source_sha !== plan[`${format}_published_source_sha`] ||
        config.builds?.commit?.source_sha !== plan.source_sha ||
        !/^[a-f0-9]{64}$/.test(config.checker_sha256 ?? '') ||
        !Array.isArray(part.samples) || part.samples.length !== expected.length)
      throw new Error('Invalid roundtrip report identity');
    if (checker && checker !== config.checker_sha256) throw new Error('Roundtrip checkers differ across formats');
    checker = config.checker_sha256;
    if (Object.keys(config.builds).length !== CHANNELS.length ||
        !['harness_sha256', 'rustc', 'profile'].every(key =>
          typeof config.builds.published[key] === 'string' && config.builds.published[key].length &&
          config.builds.published[key] === config.builds.commit[key]))
      throw new Error('Roundtrip builds use different hosts or compilers');
    for (const build of Object.values(config.builds)) {
      if (!/^[a-f0-9]{64}$/.test(build.harness_sha256 ?? '') || !/^[a-f0-9]{64}$/.test(build.binary_sha256 ?? ''))
        throw new Error('Invalid roundtrip build hashes');
    }
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

export async function stageRoundtrip(planPath, output, format, cacheDir) {
  const bytes = await readFile(planPath);
  const plan = validatePlan(JSON.parse(bytes));
  if (!plan.formats.includes(format) || !plan[`${format}_published_source_sha`])
    throw new Error('Missing roundtrip format or published source');
  const samples = plan.samples.filter(sample => sample.format === format);
  let completed = 0;
  await mapPool(samples, 8, async sample => {
    const directory = resolve(output, sample.id);
    await mkdir(directory, { recursive: true });
    await writeFile(resolve(directory, `source.${format}`), await fetchAsset(sample.metadata.source, sample.id,
      download, { cacheDir, origin: CORPUS_ORIGIN, maximum: 128 * 1024 * 1024 }));
    if (++completed % 25 === 0 || completed === samples.length)
      console.log(`${format.toUpperCase()} roundtrip: staged ${completed}/${samples.length} verified files`);
  });
  await writeFile(resolve(output, 'job.json'), JSON.stringify({ plan_sha256: digest(bytes), samples, format, method: METHOD,
    source_sha: plan.source_sha, published_source_sha: plan[`${format}_published_source_sha`],
    published_version: plan.versions[format] }, null, 2) + '\n');
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { QUALITY_PLAN, QUALITY_OUTPUT, QUALITY_FORMAT, QUALITY_ASSET_CACHE } = process.env;
  if (!QUALITY_PLAN || !QUALITY_OUTPUT || !QUALITY_FORMAT)
    throw new Error('QUALITY_PLAN, QUALITY_OUTPUT and QUALITY_FORMAT are required');
  await stageRoundtrip(resolve(QUALITY_PLAN), resolve(QUALITY_OUTPUT), QUALITY_FORMAT, QUALITY_ASSET_CACHE);
}
