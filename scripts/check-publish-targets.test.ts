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
async function guard(mode: string, respond: (name: string) => Response, env: object = {}) {
  const server = Bun.serve({
    port: 0,
    hostname: '127.0.0.1',
    fetch: (request) => respond(decodeURIComponent(new URL(request.url).pathname.slice(1)))
  });
  const origin = `http://127.0.0.1:${server.port}`;
  try {
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
    return { stdout, stderr, status };
  } finally {
    server.stop(true);
  }
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

  test('crates.io is checked before anything is uploaded to it', () => {
    expect(steps.indexOf('Check crates.io publish targets')).toBeLessThan(
      steps.indexOf('Publish Rust crates')
    );
    expect(guards[0].env.CRATES_IO_BOOTSTRAP_TOKEN).toBe('${{ secrets.CRATES_IO_BOOTSTRAP_TOKEN }}');
  });

  test('npm is checked after the pin and before changesets publishes', () => {
    expect(steps.indexOf('Pin workspace deps for npm publish')).toBeLessThan(
      steps.indexOf('Check npm publish targets')
    );
    expect(steps.indexOf('Check npm publish targets')).toBeLessThan(
      steps.indexOf('Release PR or publish')
    );
  });
});

describe('the mode argument', () => {
  test('an unknown mode fails instead of checking nothing', async () => {
    const result = await guard('--npmm', () => Response.json({}));
    expect(result.status).toBe(2);
    expect(result.stderr).toContain('--npm or --crates');
  });
});
