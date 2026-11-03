export const MEASURE_BUDGET_SECONDS = 45 * 60;

export const MEASURE_COST = {
  docx: { sample: 10, page: 2.2 },
  pptx: { sample: 5, page: 1.1 },
  xlsx: { sample: 4, page: 1 },
};

export function measureSeconds(sample) {
  const cost = MEASURE_COST[sample.format];
  const pages = sample.metadata?.reference?.pages;
  if (!cost) throw new Error(`Unknown measure format ${sample.format}`);
  if (!Number.isSafeInteger(pages) || pages < 1) throw new Error('Invalid reference page count');
  return cost.sample + cost.page * pages;
}

export function measureShards(plan, budget = MEASURE_BUDGET_SECONDS) {
  const shards = [];
  for (const format of plan.formats) {
    const samples = plan.samples.filter((sample) => sample.format === format);
    if (!samples.length) continue;
    const costs = new Map(samples.map((sample) => [sample.id, measureSeconds(sample)]));
    const total = [...costs.values()].reduce((sum, seconds) => sum + seconds, 0);
    const count = Math.min(samples.length, Math.max(1, Math.ceil(total / budget)));
    const groups = Array.from({ length: count }, () => ({ samples: [], seconds: 0 }));
    for (const sample of [...samples].sort((a, b) => costs.get(b.id) - costs.get(a.id) || a.id.localeCompare(b.id))) {
      const group = groups.reduce((lightest, candidate) => (candidate.seconds < lightest.seconds ? candidate : lightest));
      group.samples.push(sample.id);
      group.seconds += costs.get(sample.id);
    }
    groups.forEach((group, shard) => shards.push({ format, shard, samples: group.samples, seconds: Math.round(group.seconds) }));
  }
  return shards;
}

export const MAX_MATRIX_JOBS = 256;

export function measureMatrix(plan) {
  const shards = measureShards(plan);
  if (shards.length > MAX_MATRIX_JOBS)
    throw new Error(`Measurement needs ${shards.length} shards; one workflow matrix holds ${MAX_MATRIX_JOBS}`);
  return shards.map(({ format, shard }) => ({ format, shard }));
}

export function measureShard(plan, format, shard) {
  const found = measureShards(plan).find((entry) => entry.format === format && entry.shard === shard);
  if (!found) throw new Error(`No measure shard ${shard} for ${format}`);
  return found;
}

export function combineShardReports(plan, format, reports) {
  const shards = measureShards(plan).filter((entry) => entry.format === format);
  if (!Array.isArray(reports) || reports.length !== shards.length)
    throw new Error(`Missing ${format} measure shards`);
  const [first] = reports;
  const samples = [];
  shards.forEach((shard, index) => {
    const report = reports[index];
    if (report?.format !== format || report.shard !== shard.shard)
      throw new Error(`${format} measure shard ${shard.shard} is missing or mislabelled`);
    for (const key of ['source_sha', 'commit', 'react_version'])
      if (report[key] !== first[key]) throw new Error(`${format} measure shards disagree on ${key}`);
    if (JSON.stringify(report.versions) !== JSON.stringify(first.versions))
      throw new Error(`${format} measure shards disagree on versions`);
    const ids = (report.samples ?? []).map((sample) => sample?.id);
    if (ids.length !== shard.samples.length || ids.some((id) => !shard.samples.includes(id)))
      throw new Error(`${format} measure shard ${shard.shard} does not hold its planned samples`);
    samples.push(...report.samples);
  });
  const combined = { ...first, samples };
  delete combined.shard;
  return combined;
}
