import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { publishedPackageVersions } from './published-packages.mjs';
import { RUST_CRATES } from './rust-crates.mjs';

const script = fileURLToPath(new URL('./check-publish-targets.mjs', import.meta.url));
const releaseWorkflow = fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url));

const release = Bun.YAML.parse(readFileSync(releaseWorkflow, 'utf8')) as any;
const packages = publishedPackageVersions();
const crates = RUST_CRATES.map((crate) => crate.name);

const NOT_FOUND = new Response('{"error":"Not found"}', { status: 404 });

function versions(...published: string[]) {
  return Response.json({ versions: Object.fromEntries(published.map((v) => [v, {}])) });
}

// Bun.spawnSync would block the loop this fake registry answers on.
async function guard(
  mode: string,
  respond: (name: string) => Response | Promise<Response>,
  env: object = {}
) {
  const server = Bun.serve({
    port: 0,
    hostname: '127.0.0.1',
    fetch: (request) => respond(decodeURIComponent(new URL(request.url).pathname.slice(1)))
  });
  const origin = `http://127.0.0.1:${server.port}`;
  try {
    const started = performance.now();
    const child = Bun.spawn(['node', script, mode], {
      stdout: 'pipe',
      stderr: 'pipe',
      env: {
        ...process.env,
        CRATES_IO_BOOTSTRAP_TOKEN: '',
        NPM_REGISTRY_URL: origin,
        CRATES_REGISTRY_URL: origin,
        ...env
      }
    });
    const [stdout, stderr, status] = await Promise.all([
      new Response(child.stdout).text(),
      new Response(child.stderr).text(),
      child.exited
    ]);
    return { stdout, stderr, status, elapsed: performance.now() - started };
  } finally {
    server.stop(true);
  }
}

/** A registry that answers differently on each call for a name. */
function flaky(respond: (name: string, call: number) => Response | Promise<Response>) {
  const calls = new Map<string, number>();
  return {
    calls,
    respond: (name: string) => {
      const call = (calls.get(name) ?? 0) + 1;
      calls.set(name, call);
      return respond(name, call);
    }
  };
}

function published(name: string) {
  return versions(packages.find((entry) => entry.name === name)!.version);
}

describe('npm publish targets', () => {
  test('a package that exists at its release version passes', async () => {
    const result = await guard('--npm', (name) => {
      const found = packages.find((entry) => entry.name === name);
      return found ? versions(found.version) : NOT_FOUND;
    });
    expect(result.status).toBe(0);
    expect(result.stdout).toContain(`${packages[0]!.name}@${packages[0]!.version} is on npm.`);
  });

  test('a package that exists at an older version passes: that is a normal publish', async () => {
    const result = await guard('--npm', () => versions('0.0.0'));
    expect(result.status).toBe(0);
    expect(result.stdout).toContain('is a new version of a package on npm.');
    expect(result.stderr).toBe('');
  });

  test('a package that does not exist fails, by name', async () => {
    const missing = packages[0]!.name;
    const result = await guard('--npm', (name) => {
      if (name === missing) return NOT_FOUND;
      const found = packages.find((entry) => entry.name === name);
      return versions(found!.version);
    });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${missing} is not on npm.`);
    expect(result.stderr).toContain('Initial npm release of a new package');
  });
});

describe('crates.io publish targets', () => {
  test('a crate that exists passes', async () => {
    const result = await guard('--crates', () => Response.json({}));
    expect(result.status).toBe(0);
    expect(result.stdout).toContain(`${crates[0]} is on crates.io.`);
  });

  test('a crate that does not exist fails, by name', async () => {
    const result = await guard('--crates', (name) =>
      name === crates[0] ? NOT_FOUND : Response.json({})
    );
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${crates[0]} is not on crates.io.`);
    expect(result.stderr).toContain('Initial crates.io release');
  });

  test('a bootstrap token creates crates, so a missing one only prints', async () => {
    const result = await guard(
      '--crates',
      (name) => (name === crates[0] ? NOT_FOUND : Response.json({})),
      { CRATES_IO_BOOTSTRAP_TOKEN: 'cio_bootstrap' }
    );
    expect(result.status).toBe(0);
    expect(result.stdout).toContain(`${crates[0]} is not on crates.io; CRATES_IO_BOOTSTRAP_TOKEN`);
    expect(result.stderr).toBe('');
  });
});

describe('release wiring', () => {
  const steps = release.jobs.release.steps.map((step: any) => step.name ?? step.uses);
  const guards = release.jobs.release.steps.filter((step: any) =>
    step.run?.includes('check-publish-targets.mjs')
  );

  test('both guards run only on the publish path', () => {
    expect(guards.map((step: any) => step.run.trim())).toEqual([
      'node scripts/check-publish-targets.mjs --crates',
      'node scripts/check-publish-targets.mjs --npm'
    ]);
    for (const step of guards) {
      expect(step.if).toBe("steps.pending.outputs.publishing == 'true'");
    }
  });

  test('both registries are checked before the first upload of either', () => {
    const publishes = ['Publish Rust crates', 'Release PR or publish'].map((step) =>
      steps.indexOf(step)
    );
    for (const guard of ['Check crates.io publish targets', 'Check npm publish targets']) {
      for (const publish of publishes) expect(steps.indexOf(guard)).toBeLessThan(publish);
    }
    expect(guards[0].env.CRATES_IO_BOOTSTRAP_TOKEN).toBe('${{ secrets.CRATES_IO_BOOTSTRAP_TOKEN }}');
  });

  test('the npm guard reads what the pin that follows it cannot change', () => {
    expect(steps.indexOf('Check npm publish targets')).toBeLessThan(
      steps.indexOf('Pin workspace deps for npm publish')
    );
    for (const entry of packages) expect(Object.keys(entry).sort()).toEqual(['name', 'version']);
  });
});

describe('the npm publish set', () => {
  test('is every non-private workspace, so no new package escapes the guard', () => {
    const names = packages.map((entry) => entry.name);
    expect(names).toContain('@betteroffice/fonts');
    expect(names).not.toContain('@betteroffice/rust-crates');
    expect(names).not.toContain('@betteroffice/collaboration-relay');
  });
});

describe('a registry that misbehaves', () => {
  const first = packages[0]!.name;

  test('waits out a 429 for as long as its Retry-After asks', async () => {
    const registry = flaky((name, call) =>
      name === first && call === 1
        ? new Response('{"error":"too many requests"}', {
            status: 429,
            headers: { 'Retry-After': '1' }
          })
        : published(name)
    );
    const result = await guard('--npm', registry.respond);

    expect(result.status).toBe(0);
    expect(registry.calls.get(first)).toBe(2);
    expect(result.elapsed).toBeGreaterThan(900);
  });

  test('gives up on a request that hangs, and retries it', async () => {
    const registry = flaky((name, call) =>
      name === first && call === 1 ? new Promise<Response>(() => {}) : published(name)
    );
    const result = await guard('--npm', registry.respond, { REGISTRY_TIMEOUT_MS: '400' });

    expect(result.status).toBe(0);
    expect(registry.calls.get(first)).toBe(2);
    expect(result.elapsed).toBeGreaterThan(400);
  });

  test('retries a body that stops mid-JSON', async () => {
    const registry = flaky((name, call) =>
      name === first && call === 1
        ? new Response('{"versions":{"0.0.1"', { headers: { 'Content-Type': 'application/json' } })
        : published(name)
    );
    const result = await guard('--npm', registry.respond);

    expect(result.status).toBe(0);
    expect(registry.calls.get(first)).toBe(2);
  });

  test('stops after five attempts instead of publishing on a guess', async () => {
    const registry = flaky(
      () =>
        new Response('{"error":"unavailable"}', { status: 503, headers: { 'Retry-After': '0' } })
    );
    const result = await guard('--npm', registry.respond);

    expect(result.status).not.toBe(0);
    expect(registry.calls.get(first)).toBe(5);
    expect(result.stderr).toContain('503');
  });

  test('a 4xx that is not a rate limit is not retried', async () => {
    const registry = flaky(() => new Response('{"error":"gone"}', { status: 410 }));
    const result = await guard('--npm', registry.respond);

    expect(result.status).not.toBe(0);
    expect(registry.calls.get(first)).toBe(1);
    expect(result.stderr).toContain('410');
  });
});

describe('the mode argument', () => {
  test('an unknown mode fails instead of checking nothing', async () => {
    const result = await guard('--npmm', () => Response.json({}));
    expect(result.status).toBe(2);
    expect(result.stderr).toContain('--npm or --crates');
  });
});
