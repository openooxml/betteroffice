const MAX_REASON = 160;

export function failureRecords(sample) {
  const records = [];
  const add = (area, channel, stage, error) => records.push({ area, channel, stage: stage ?? area, error: String(error ?? '') });
  for (const result of sample.comparisons ?? [])
    if (result?.status === 'failed') add('render', result.channel, result.stage, result.error);
  for (const [channel, timing] of Object.entries(sample.native_timings ?? {}))
    if (timing?.status === 'failed') add('timing', channel, 'timing', timing.error);
  for (const [channel, calculation] of Object.entries(sample.calculations ?? {}))
    if (calculation?.status === 'failed') add('recalc', channel, 'calculation', calculation.error);
  for (const [channel, outcome] of Object.entries(sample.roundtrip ?? {})) {
    if (outcome?.parse?.status === 'failed') add('parse', channel, 'parse', outcome.parse.error);
    else if (outcome?.roundtrip?.status === 'failed') add('roundtrip', channel, outcome.roundtrip.stage, outcome.roundtrip.error);
  }
  if (sample.roundtrip_probe?.status === 'failed') add('probe', null, 'probe', sample.roundtrip_probe.error);
  return records;
}

const escape = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function normalizeReason(message, secrets = []) {
  let text = String(message ?? '');
  const inner = text.match(/"error"\s*:\s*"((?:[^"\\]|\\.)*)"/);
  if (inner) text = inner[1];
  for (const secret of [...secrets].filter(Boolean).sort((a, b) => b.length - a.length))
    text = text.replace(new RegExp(escape(secret), 'g'), '<sample>');
  text = text
    .replace(/[a-z][a-z0-9+.-]*:\/\/\S+/gi, '<url>')
    .replace(/(^|[\s"'(=])(?:[A-Za-z]:)?(?:\/|\\)[^\s"')]+/g, '$1<path>')
    .replace(/\b[a-f0-9]{7,}\b/gi, '<hash>')
    .replace(/\d+(?:\.\d+)?/g, 'N')
    .replace(/\s+/g, ' ')
    .trim();
  if (!text) return 'no message';
  return text.length > MAX_REASON ? `${text.slice(0, MAX_REASON - 1)}…` : text;
}

const order = (a, b) => b.count - a.count || a.key.localeCompare(b.key);

export function failureClusters(report) {
  const secrets = report.samples.flatMap((sample) => [sample.id, sample.source_sha256, sample.metadata_url]);
  const clusters = new Map();
  for (const sample of report.samples) {
    for (const record of failureRecords(sample)) {
      const reason = normalizeReason(record.error, secrets);
      const key = JSON.stringify([sample.format, record.area, record.channel, record.stage, reason]);
      const cluster = clusters.get(key) ?? {
        key, format: sample.format, area: record.area, channel: record.channel, stage: record.stage, reason, count: 0, samples: [],
      };
      cluster.count += 1;
      if (!cluster.samples.includes(sample.id)) cluster.samples.push(sample.id);
      clusters.set(key, cluster);
    }
  }
  return [...clusters.values()].sort(order).map(({ key, ...cluster }) => cluster);
}

export function failureCounts(report, format) {
  const counts = new Map();
  for (const sample of report.samples.filter((entry) => entry.format === format))
    for (const { area, channel, stage } of failureRecords(sample)) {
      const key = JSON.stringify([area, channel, stage]);
      const entry = counts.get(key) ?? { key, area, channel, stage, count: 0 };
      entry.count += 1;
      counts.set(key, entry);
    }
  return [...counts.values()].sort(order).map(({ key, ...entry }) => entry);
}
