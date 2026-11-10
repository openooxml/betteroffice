import { measureShards } from './measure-shards.mjs';
import { roundtripMatrices } from './roundtrip.mjs';
import { jobShards } from './shards.mjs';
import { timingFormats, timingSamples } from './timing.mjs';

export const UNBOUNDED = 256;
export const DEFAULT_PRIVATE_PARALLEL = 12;
const SETUP_MINUTES = { measure: 8, benchmark: 5, roundtrip: 5, timing: 10 };
const TIMING_MINUTES_PER_DOCUMENT = { docx: 0.2, pptx: 0.25, xlsx: 0.35 };
const BENCHMARKS = ['docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity'];

export function matrixLoad(plan) {
  const measure = measureShards(plan);
  const [roundtripA, roundtripB] = roundtripMatrices(plan);
  const formats = timingFormats(plan);
  const load = {
    measure: { jobs: measure.length, minutes: measure.reduce((sum, shard) => sum + shard.seconds / 60 + SETUP_MINUTES.measure, 0) },
  };
  for (const job of BENCHMARKS) {
    const shards = jobShards(job, plan);
    load[job] = { jobs: shards.length, minutes: shards.reduce((sum, shard) => sum + shard.seconds / 60 + SETUP_MINUTES.benchmark, 0) };
  }
  load['roundtrip-a'] = { jobs: roundtripA.length, minutes: roundtripA.length * SETUP_MINUTES.roundtrip };
  load['roundtrip-b'] = { jobs: roundtripB.length, minutes: roundtripB.length * SETUP_MINUTES.roundtrip };
  load.timing = {
    jobs: formats.length,
    minutes: formats.reduce((sum, format) =>
      sum + SETUP_MINUTES.timing + timingSamples(plan, format).length * TIMING_MINUTES_PER_DOCUMENT[format], 0),
  };
  return load;
}

const runningMatrices = (load) => Object.keys(load).filter((name) => load[name].jobs > 0 && load[name].minutes > 0);

export function allocate(load, total) {
  if (!Number.isSafeInteger(total) || total < 1) throw new Error('max_parallel must be a positive whole number');
  const running = runningMatrices(load);
  if (total < running.length)
    throw new Error(`max_parallel ${total} is below the ${running.length} benchmark matrices this plan runs`);
  const share = Object.fromEntries(Object.keys(load).map((name) => [name, 1]));
  for (let left = total - running.length; left > 0; left -= 1) {
    const open = running.filter((name) => share[name] < load[name].jobs);
    if (!open.length) break;
    const next = open.reduce((best, name) => {
      const wall = load[name].minutes / share[name];
      const bestWall = load[best].minutes / share[best];
      return wall > bestWall || (wall === bestWall && name < best) ? name : best;
    });
    share[next] += 1;
  }
  return share;
}

export function parallelPlan(plan, { privateRun, total = DEFAULT_PRIVATE_PARALLEL }) {
  const load = matrixLoad(plan);
  if (!privateRun) return Object.fromEntries(Object.keys(load).map((name) => [name, UNBOUNDED]));
  return allocate(load, total);
}

const hours = (minutes) => (minutes / 60).toFixed(1);

export function parallelSummary(plan, share) {
  const load = matrixLoad(plan);
  const running = runningMatrices(load);
  const lines = running.map((name) =>
    `${name}: at most ${share[name]} at once for ${load[name].jobs} job${load[name].jobs === 1 ? '' : 's'}, ~${Math.round(load[name].minutes)} runner-minutes planned, about ${hours(load[name].minutes / share[name])} h wall`);
  const wall = Math.max(0, ...running.map((name) => load[name].minutes / share[name]));
  const jobs = running.reduce((sum, name) => sum + load[name].jobs, 0);
  return [...lines, `projected wall of the benchmark matrices: about ${hours(wall)} h for ${jobs} jobs`];
}
