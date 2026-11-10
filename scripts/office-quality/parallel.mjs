import { measureShards } from './measure-shards.mjs';
import { roundtripMatrices } from './roundtrip.mjs';
import { jobShards } from './shards.mjs';
import { timingFormats } from './timing.mjs';

export const UNBOUNDED = 256;
export const DEFAULT_PRIVATE_PARALLEL = 12;
const SETUP_MINUTES = { measure: 8, benchmark: 5, roundtrip: 5, timing: 10 };
const TIMING_MINUTES = 60;

export function matrixMinutes(plan) {
  const measure = measureShards(plan);
  const benchmark = (job) => jobShards(job, plan).reduce((sum, shard) => sum + shard.seconds / 60 + SETUP_MINUTES.benchmark, 0);
  const [roundtripA, roundtripB] = roundtripMatrices(plan);
  return {
    measure: measure.reduce((sum, shard) => sum + shard.seconds / 60 + SETUP_MINUTES.measure, 0),
    'docx-benchmark': benchmark('docx-benchmark'),
    'xlsx-benchmark': benchmark('xlsx-benchmark'),
    'pptx-benchmark': benchmark('pptx-benchmark'),
    'xlsx-fidelity': benchmark('xlsx-fidelity'),
    'roundtrip-a': roundtripA.length * SETUP_MINUTES.roundtrip,
    'roundtrip-b': roundtripB.length * SETUP_MINUTES.roundtrip,
    timing: timingFormats(plan).length * (TIMING_MINUTES + SETUP_MINUTES.timing),
  };
}

export function parallelPlan(plan, { privateRun, total = DEFAULT_PRIVATE_PARALLEL }) {
  const minutes = matrixMinutes(plan);
  const names = Object.keys(minutes);
  if (!privateRun) return Object.fromEntries(names.map((name) => [name, UNBOUNDED]));
  if (!Number.isSafeInteger(total) || total < 1) throw new Error('max_parallel must be a positive whole number');
  const running = names.filter((name) => minutes[name] > 0);
  if (total < running.length)
    throw new Error(`max_parallel ${total} is below the ${running.length} benchmark matrices this plan runs`);
  const share = Object.fromEntries(names.map((name) => [name, 1]));
  const sum = running.reduce((acc, name) => acc + minutes[name], 0);
  const spare = total - running.length;
  const exact = running.map((name) => ({ name, value: (spare * minutes[name]) / sum }));
  for (const { name, value } of exact) share[name] += Math.floor(value);
  let left = total - running.reduce((acc, name) => acc + share[name], 0);
  for (const { name } of [...exact].sort((a, b) => (b.value % 1) - (a.value % 1) || a.name.localeCompare(b.name))) {
    if (left <= 0) break;
    share[name] += 1;
    left -= 1;
  }
  return share;
}

export function parallelSummary(plan, share) {
  const minutes = matrixMinutes(plan);
  return Object.keys(share)
    .filter((name) => minutes[name] > 0)
    .map((name) => `${name}: at most ${share[name]} at once, ~${Math.round(minutes[name])} runner-minutes planned`);
}
