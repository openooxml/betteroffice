export const JOB_TIMEOUT_MINUTES = 90;
export const SHARD_BUDGET_SECONDS = 60 * 60;

export const SHARD_COST = {
  'docx-benchmark': { sample: 10, page: 0.6 },
  'pptx-benchmark': { sample: 12, page: 0.8 },
  'xlsx-fidelity': { sample: 1, page: 1.5 },
  'xlsx-benchmark': { sample: 4, cell: 0.002 },
};

const BASE_SHARDS = { 'docx-benchmark': 4, 'pptx-benchmark': 1, 'xlsx-fidelity': 4, 'xlsx-benchmark': 4 };

function positive(value, label) {
  if (!Number.isSafeInteger(value) || value < 1) throw new Error(`Invalid ${label}`);
  return value;
}

export function shardSeconds(job, sample) {
  const cost = SHARD_COST[job];
  if (!cost) throw new Error(`Unknown benchmark job ${job}`);
  if (job === 'xlsx-benchmark')
    return cost.sample + cost.cell * positive(sample.metadata?.calculation?.result_cells, 'calculation reference size');
  return cost.sample + cost.page * positive(sample.metadata?.reference?.pages, 'reference page count');
}

function weight(job, sample) {
  return job === 'xlsx-benchmark' ? sample.metadata.calculation.result_cells : sample.metadata.reference.pages;
}

function assign(job, samples, count) {
  const shards = Array.from({ length: count }, () => ({ samples: [], weight: 0, seconds: 0 }));
  for (const sample of [...samples].sort((a, b) => weight(job, b) - weight(job, a) || a.id.localeCompare(b.id))) {
    const shard = shards.reduce((lightest, candidate) => (candidate.weight < lightest.weight ? candidate : lightest));
    shard.samples.push(sample.id);
    shard.weight += weight(job, sample);
    shard.seconds += shardSeconds(job, sample);
  }
  return shards;
}

export function budgetShards(job, samples, budget = SHARD_BUDGET_SECONDS) {
  if (!samples.length) return [];
  for (const sample of samples)
    if (shardSeconds(job, sample) > budget)
      throw new Error(`${job}: one document is planned at ${Math.round(shardSeconds(job, sample))} s, over the ${budget} s shard budget`);
  const total = samples.reduce((sum, sample) => sum + shardSeconds(job, sample), 0);
  let count = Math.min(samples.length, Math.max(BASE_SHARDS[job], Math.ceil(total / budget)));
  for (;;) {
    const shards = assign(job, samples, count);
    if (shards.every((shard) => shard.seconds <= budget)) return shards;
    if (count === samples.length) throw new Error(`${job}: cannot keep every shard inside the ${budget} s budget`);
    count += 1;
  }
}

const JOB_SAMPLES = {
  'docx-benchmark': (sample) => sample.format === 'docx',
  'pptx-benchmark': (sample) => sample.format === 'pptx',
  'xlsx-fidelity': (sample) => sample.format === 'xlsx',
  'xlsx-benchmark': (sample) => sample.format === 'xlsx' && sample.metadata.calculation?.status === 'ok',
};

export function benchmarkShardPlan(plan, budget = SHARD_BUDGET_SECONDS) {
  return Object.entries(JOB_SAMPLES).flatMap(([job, selected]) => {
    const shards = budgetShards(job, plan.samples.filter(selected), budget);
    if (!shards.length) return [];
    const longest = Math.max(...shards.map((shard) => Math.round(shard.seconds)));
    return [`${job}: ${shards.length} shard${shards.length === 1 ? '' : 's'}, longest planned at ${longest} s (budget ${budget} s, job timeout ${JOB_TIMEOUT_MINUTES} min)`];
  });
}
