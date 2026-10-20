import { appendFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const GITHUB_HOSTED = /^(ubuntu|macos|windows)-[\w.-]+$/;
export const LABEL = 'ubuntu-latest';
export const BATCH_RUNS = 10;

/** `packages python package_tests:2 rust`: job groups in the order they claim GitHub-hosted slots. */
export function parseGroups(spec) {
  return spec
    .split(/\s+/)
    .filter(Boolean)
    .map((entry) => {
      const [name, slots = '1', extra] = entry.split(':');
      const count = Number(slots);
      if (extra !== undefined || !/^[a-z][a-z0-9_]*$/.test(name) || !Number.isInteger(count) || count < 1) {
        throw new Error(`invalid route group ${entry}`);
      }
      return { name, slots: count };
    });
}

/** First fit: a group that does not fit stays on kyora, and a later, smaller one may still fit. */
export function allocate(groups, free) {
  const routes = {};
  for (const { name, slots } of groups) {
    routes[name] = slots <= free ? LABEL : '';
    if (routes[name]) free -= slots;
  }
  return routes;
}

export function share(free, own, peer) {
  return own + peer > 0 ? Math.floor((free * own) / (own + peer)) : 0;
}

export function hostedJobs(jobs) {
  return jobs.filter(
    (job) =>
      (job.status === 'queued' || job.status === 'in_progress') &&
      job.labels?.length === 1 &&
      GITHUB_HOSTED.test(job.labels[0])
  ).length;
}

/**
 * GitHub-hosted jobs queued or running in the repository, routed or not. Fetches run job pages in
 * batches of BATCH_RUNS and stops after the batch in which the count reaches `cap`.
 */
export async function busyHosted(api, repository, cap) {
  const runs = [];
  for (const status of ['in_progress', 'queued']) {
    const page = await api(`/repos/${repository}/actions/runs?status=${status}&per_page=100`);
    const listed = page.workflow_runs ?? [];
    if (page.total_count > listed.length) return cap;
    for (const { id } of listed) if (!runs.includes(id)) runs.push(id);
  }
  let busy = 0;
  for (let i = 0; i < runs.length; i += BATCH_RUNS) {
    const pages = await Promise.all(
      runs.slice(i, i + BATCH_RUNS).map((run) => api(`/repos/${repository}/actions/runs/${run}/jobs?filter=latest&per_page=100`))
    );
    for (const page of pages) {
      const jobs = page.jobs ?? [];
      if (page.total_count > jobs.length) return cap;
      busy += hostedJobs(jobs);
    }
    if (busy >= cap) return busy;
  }
  return busy;
}

function count(value, fallback) {
  return /^\d+$/.test(value ?? '') ? Number(value) : fallback;
}

export async function route(env, fetchImpl = fetch, signal = AbortSignal.timeout(45_000)) {
  const groups = parseGroups(env.ROUTE_GROUPS ?? '');
  if (env.BO_CI_SPILL === 'off' || env.BO_CI_RUNNER) return { routes: allocate(groups, 0), busy: null };
  const capacity = Math.max(0, count(env.BO_CI_GH_LIMIT, 20) - count(env.BO_CI_GH_RESERVE, 2));
  const api = async (path) => {
    const response = await fetchImpl(`${env.GITHUB_API_URL ?? 'https://api.github.com'}${path}`, {
      headers: {
        Accept: 'application/vnd.github+json',
        Authorization: `Bearer ${env.GH_TOKEN}`,
        'X-GitHub-Api-Version': '2022-11-28'
      },
      signal: AbortSignal.any([signal, AbortSignal.timeout(10_000)])
    });
    if (!response.ok) throw new Error(`GitHub answered ${response.status} for ${path}`);
    return response.json();
  };
  const busy = await busyHosted(api, env.GITHUB_REPOSITORY, capacity);
  const own = groups.reduce((sum, { slots }) => sum + slots, 0);
  const free = share(Math.max(0, capacity - busy), own, count(env.ROUTE_PEER_SLOTS, 0));
  return { routes: allocate(groups, free), busy };
}

// Never fails the job: any error or the deadline leaves every output empty, which keeps the jobs on their kyora class.
if (process.argv[1] === fileURLToPath(import.meta.url)) {
  let groups = [];
  let result;
  try {
    groups = parseGroups(process.env.ROUTE_GROUPS ?? '');
    const deadline = AbortSignal.timeout(count(process.env.BO_CI_ROUTE_DEADLINE_MS, 45_000));
    result = await route(process.env, fetch, deadline);
  } catch (error) {
    console.log(`::warning::runner routing fell back to kyora: ${error.message}`);
    result = { routes: allocate(groups, 0), busy: null };
  }
  const lines = Object.entries(result.routes).map(([name, label]) => `${name}=${label}`);
  if (process.env.GITHUB_OUTPUT) appendFileSync(process.env.GITHUB_OUTPUT, `${lines.join('\n')}\n`);
  console.log(`GitHub-hosted jobs busy: ${result.busy ?? 'not counted'}`);
  for (const [name, label] of Object.entries(result.routes)) console.log(`${name}: ${label || 'kyora'}`);
  process.exit(0);
}
