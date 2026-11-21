import { afterEach, describe, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkCrates, checkNpm, checkPlan, missingHeldNatives, nativeReleaseVersions, repositoryScan, unaskedMajors, siblingReferences, verifyPublished } from './check-publish-targets.mjs';
import { releaseNotes, releaseTag } from './held-npm-packages.mjs';
import { platformPackageVersions } from './native-bindings.mjs';
import { publishedCrates, publishedPackageVersions, workspaceEntries } from './published-packages.mjs';
import { RUST_PUBLISH_CRATES } from './rust-crates.mjs';

const script = fileURLToPath(new URL('./check-publish-targets.mjs', import.meta.url));
const releaseWorkflow = fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url));

const release = Bun.YAML.parse(readFileSync(releaseWorkflow, 'utf8')) as any;
const packages = publishedPackageVersions();
const agents = packages.find(({ name }) => name === '@betteroffice/agents')!;
const repository = repositoryScan(packages);
const workspaceVersions = new Map(workspaceEntries().map(({ manifest }) => [manifest.name, manifest.version]));
const platforms = platformPackageVersions();
const platformVersions = new Map(platforms.map(({ name, version }) => [name, version]));
const crates = RUST_PUBLISH_CRATES.map((crate) => crate.name);
const fixtureDirectories: string[] = [];
const fonts = { name: '@betteroffice/fonts', version: '1.2.3' };
const latin = { name: '@betteroffice/fonts-latin', version: '1.2.3' };

afterEach(() => {
  for (const directory of fixtureDirectories.splice(0)) rmSync(directory, { recursive: true, force: true });
});

function workspaceFixture(manifests: Array<{ name: string; version: string; [key: string]: unknown }>) {
  const root = mkdtempSync(join(tmpdir(), 'publish-targets-'));
  fixtureDirectories.push(root);
  return manifests.map((manifest) => {
    const directory = join(root, manifest.name.split('/').pop()!);
    mkdirSync(join(directory, 'src'), { recursive: true });
    return { directory, manifest };
  });
}

function fontWorkspaces(field = 'dependencies', range = 'workspace:*') {
  return workspaceFixture([{ ...fonts, [field]: { [latin.name]: range } }, latin]);
}

type RegistryResponder = (name: string, request: Request) => Response | Promise<Response>;

function notFound() {
  return new Response('{"error":"Not found"}', { status: 404 });
}

function versions(...published: string[]) {
  return Response.json({ versions: Object.fromEntries(published.map((v) => [v, {}])) });
}

function registry(respond: RegistryResponder) {
  const server = Bun.serve({
    port: 0,
    hostname: '127.0.0.1',
    fetch: (request) => respond(decodeURIComponent(new URL(request.url).pathname.slice(1)), request)
  });
  return { server, origin: `http://127.0.0.1:${server.port}` };
}

/** Runs a guard in this process against a fake registry, collecting what it prints. */
async function check(
  mode: '--npm' | '--crates' | '--verify-published',
  respond: RegistryResponder,
  options: {
    timeoutMs?: number;
    bootstrap?: boolean;
    packages?: typeof packages;
    workspaces?: ReturnType<typeof workspaceEntries>;
    later?: { name: string; version: string }[];
    platforms?: { name: string; version: string }[];
    natives?: { name: string; version: string }[];
    holdNative?: boolean;
    held?: typeof packages;
    references?: ReturnType<typeof siblingReferences>;
    maxWaitMs?: number;
  } = {}
) {
  const { server, origin } = registry(respond);
  const stdout: string[] = [];
  const stderr: string[] = [];
  let held: typeof packages = [];
  const out = { log: (line: string) => stdout.push(line), error: (line: string) => stderr.push(line) };
  try {
    const started = performance.now();
    const npmOptions = {
      registry: origin, out, timeoutMs: options.timeoutMs,
      packages: options.packages, workspaces: options.workspaces, later: options.later, platforms: options.platforms,
      natives: options.natives, holdNative: options.holdNative, references: options.references,
      onHeld: (entries: typeof packages) => { held = entries; }
    };
    const status =
      mode === '--npm'
        ? await checkNpm(npmOptions)
        : mode === '--verify-published'
          ? await verifyPublished({ ...npmOptions, cdnBase: origin, maxWaitMs: options.maxWaitMs, held: options.held })
          : await checkCrates({ registry: origin, out, timeoutMs: options.timeoutMs, bootstrap: options.bootstrap ?? false });
    return { status, held, stdout: stdout.join('\n'), stderr: stderr.join('\n'), elapsed: performance.now() - started };
  } finally {
    server.stop(true);
  }
}

const SPAWN_BUDGET_MS = 30_000;

// Bun.spawnSync would block the loop this fake registry answers on.
async function guard(
  mode: string,
  respond: RegistryResponder,
  env: object = {}
) {
  const { server, origin } = registry(respond);
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
        CDN_BASE_URL: origin,
        GITHUB_ACTIONS: '',
        GITHUB_OUTPUT: '',
        GITHUB_STEP_SUMMARY: '',
        HELD_NPM_PACKAGES: '',
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
  return versions(workspaceVersions.get(name)!);
}

const FULL_SCAN_BUDGET_MS = 30_000;
const natives = nativeReleaseVersions();

/** npm has only 0.0.0 of every release package and native root, and the current version of everything else. */
function older(name: string) {
  return packages.some((entry) => entry.name === name) || natives.some((entry) => entry.name === name)
    ? versions('0.0.0') : published(name);
}

describe('npm publish targets', () => {
  test('VSDX packages stay out of npm publication', () => {
    for (const name of ['@betteroffice/vsdx', '@betteroffice/vsdx-react', '@betteroffice/vsdx-i18n']) {
      expect(packages.map((entry) => entry.name)).not.toContain(name);
    }
  });

  test('a package that exists at its release version passes', async () => {
    const result = await check('--npm', (name) => {
      const version = workspaceVersions.get(name) ?? platformVersions.get(name);
      return version ? versions(version) : notFound();
    }, repository);
    expect(result.status).toBe(0);
    expect(result.stdout).toContain(`${packages[0]!.name}@${packages[0]!.version} is on npm.`);
  });

  test('a package that exists at an older version passes: that is a normal publish', async () => {
    const result = await check('--npm', (name) =>
      packages.some((entry) => entry.name === name) ? versions('0.0.0') : published(name),
    repository);
    expect(result.status).toBe(0);
    expect(result.stdout).toContain('is a new version of a package on npm.');
    expect(result.stderr).toBe('');
  });

  test('a package that does not exist fails, by name', async () => {
    const missing = packages[0]!.name;
    const result = await check('--npm', (name) => {
      if (name === missing) return notFound();
      return published(name);
    }, repository);
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${missing} is not on npm.`);
    expect(result.stderr).toContain('Initial npm release of a new package');
  });
});

describe('native npm names the release publishes later', () => {
  const root = { name: '@betteroffice/docx-native', version: '1.2.3' };
  const platform = { name: '@betteroffice/docx-native-linux-x64-musl', version: '1.2.3' };
  const native = { packages: [fonts, latin], later: [root], platforms: [platform] };
  const missingLines = (stderr: string) => stderr.split('\n').filter((line) => line.endsWith(' is not on npm.'));

  test.each([
    ['platform package', platform.name],
    ['root', root.name]
  ])('a native %s that does not exist fails, by name', async (_kind, missing) => {
    const result = await check('--npm', (name) => name === missing ? notFound() : versions('1.2.3'), {
      ...native, workspaces: fontWorkspaces()
    });
    expect(result.status).toBe(1);
    expect(missingLines(result.stderr)).toEqual([`${missing} is not on npm.`]);
    expect(result.stderr).toContain('RELEASING.md, "Native packages"');
    expect(result.stderr).not.toContain('Initial npm release of a new package');
  });

  test('native names on npm at the release version pass and print nothing of their own', async () => {
    const result = await check('--npm', () => versions('1.2.3'), { ...native, workspaces: fontWorkspaces() });
    expect(result.status).toBe(0);
    expect(result.stderr).toBe('');
    expect(result.stdout).toBe([
      `${fonts.name}@${fonts.version} is on npm.`,
      `${latin.name}@${latin.version} is on npm.`
    ].join('\n'));
  });

  test('native names on npm only at an older version pass: the release publishes that version', async () => {
    const result = await check('--npm', () => versions('0.0.0'), { ...native, workspaces: fontWorkspaces() });
    expect(result.status).toBe(0);
    expect(result.stderr).toBe('');
    expect(result.stdout).toBe([
      `${fonts.name}@${fonts.version} is a new version of a package on npm.`,
      `${latin.name}@${latin.version} is a new version of a package on npm.`
    ].join('\n'));
  });

  async function withSkipNative<T>(value: string | undefined, run: () => Promise<T>) {
    const prior = process.env.RELEASE_SKIP_NATIVE;
    if (value === undefined) delete process.env.RELEASE_SKIP_NATIVE;
    else process.env.RELEASE_SKIP_NATIVE = value;
    try {
      return await run();
    } finally {
      if (prior === undefined) delete process.env.RELEASE_SKIP_NATIVE;
      else process.env.RELEASE_SKIP_NATIVE = prior;
    }
  }

  test.each(['true', 'TRUE', 'True'])('native names are neither required nor requested when RELEASE_SKIP_NATIVE is %s', async (value) => {
    const requested: string[] = [];
    const skipped = await withSkipNative(value, () => check('--npm', (name) => {
      requested.push(name);
      return name === fonts.name || name === latin.name ? versions('1.2.3') : notFound();
    }, { packages: [fonts, latin], workspaces: fontWorkspaces() }));
    expect(skipped.stderr).toBe('');
    expect(skipped.status).toBe(0);
    expect(requested.sort()).toEqual([fonts.name, latin.name].sort());
  });

  /** Runs the npm check against a fake registry and reports a rejection as `status: 'threw'`, keeping what it printed. */
  async function settledNpmCheck(respond: RegistryResponder, options: Record<string, unknown>) {
    const { server, origin } = registry(respond);
    const stdout: string[] = [];
    const stderr: string[] = [];
    const out = { log: (line: string) => stdout.push(line), error: (line: string) => stderr.push(line) };
    try {
      const status: number | 'threw' = await checkNpm({ ...options, registry: origin, out } as any);
      return { status, stdout: stdout.join('\n'), stderr: stderr.join('\n'), error: undefined as string | undefined };
    } catch (error) {
      return { status: 'threw' as const, stdout: stdout.join('\n'), stderr: stderr.join('\n'), error: String(error) };
    } finally {
      server.stop(true);
    }
  }

  test('a native name npm answers with a body that is not JSON exists', async () => {
    const result = await settledNpmCheck((name) =>
      name === platform.name ? new Response('<!doctype html>', { status: 200 }) : versions('1.2.3'), {
      ...native, workspaces: fontWorkspaces()
    });
    expect(result).toEqual({
      status: 0,
      stdout: [`${fonts.name}@${fonts.version} is on npm.`, `${latin.name}@${latin.version} is on npm.`].join('\n'),
      stderr: '',
      error: undefined
    });
  }, FULL_SCAN_BUDGET_MS);

  test('a native name npm answers with 200 and a body stream that fails exists', async () => {
    const failedBody = () => new Response(new ReadableStream({ start(controller) { controller.error(new Error('stream failed')); } }), { status: 200 });
    expect(await failedBody().body!.cancel().then(() => 'cancelled', (error: Error) => error.message)).toBe('stream failed');
    const realFetch = globalThis.fetch;
    globalThis.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(input instanceof Request ? input.url : String(input));
      return decodeURIComponent(url.pathname.slice(1)) === platform.name
        ? Promise.resolve(failedBody())
        : realFetch(input, init);
    }) as unknown as typeof fetch;
    try {
      const result = await settledNpmCheck(() => versions('1.2.3'), { ...native, workspaces: fontWorkspaces() });
      expect(result).toEqual({
        status: 0,
        stdout: [`${fonts.name}@${fonts.version} is on npm.`, `${latin.name}@${latin.version} is on npm.`].join('\n'),
        stderr: '',
        error: undefined
      });
    } finally {
      globalThis.fetch = realFetch;
    }
  }, FULL_SCAN_BUDGET_MS);

  test('a native name npm answers with a JSON null body exists', async () => {
    const result = await check('--npm', (name) =>
      name === platform.name ? Response.json(null) : versions('1.2.3'), {
      ...native, workspaces: fontWorkspaces()
    });
    expect(result.stderr).toBe('');
    expect(result.status).toBe(0);
  });

  test('a native lookup that keeps failing fails the check, naming the package', async () => {
    const result = await settledNpmCheck((name) =>
      name === platform.name
        ? new Response('{"error":"unavailable"}', { status: 503, headers: { 'Retry-After': '0' } })
        : versions('1.2.3'), {
      ...native, workspaces: fontWorkspaces()
    });
    expect({ status: result.status, error: result.error }).toEqual({ status: 1, error: undefined });
    expect(result.stderr).toContain(`${platform.name} could not be looked up on npm:`);
    expect(result.stderr).toContain('503');
  });

  test('native names found missing are named even when a later lookup fails', async () => {
    const failing = { name: '@betteroffice/docx-native-win32-arm64-msvc', version: '1.2.3' };
    const result = await settledNpmCheck((name) => {
      if (name === platform.name) return notFound();
      if (name === failing.name) return new Response(null, { status: 503, headers: { 'Retry-After': '0' } });
      return versions('1.2.3');
    }, { ...native, platforms: [platform, failing], workspaces: fontWorkspaces() });
    expect({ status: result.status, missing: missingLines(result.stderr), error: result.error }).toEqual({
      status: 1, missing: [`${platform.name} is not on npm.`], error: undefined
    });
    expect(result.stderr).toContain('RELEASING.md, "Native packages"');
    expect(result.stderr).toContain(`${failing.name} could not be looked up on npm:`);
  });

  test('the full repository scan asks npm once for every release package and every native name', async () => {
    const requested: string[] = [];
    const result = await withSkipNative(undefined, () => check('--npm', (name) => {
      requested.push(name);
      return versions(workspaceVersions.get(name) ?? platformVersions.get(name)!);
    }));
    expect(result.status).toBe(0);
    expect(natives).toHaveLength(3);
    expect(requested.sort()).toEqual([...packages, ...natives, ...platforms].map(({ name }) => name).sort());
  }, FULL_SCAN_BUDGET_MS);

  test('a full repository scan that skips native packages asks npm for no platform package, and for the roots only as dependencies', async () => {
    const requested: string[] = [];
    const result = await withSkipNative('true', () => check('--npm', (name) => {
      requested.push(name);
      return name === agents.name ? older(name) : published(name);
    }));
    expect(result.status).toBe(0);
    expect(result.held).toEqual([]);
    expect(requested.sort()).toEqual([...packages, ...natives].map(({ name }) => name).sort());
  }, FULL_SCAN_BUDGET_MS);

  test.each([
    ['platform package', platforms[platforms.length - 1]!.name],
    ['root', natives[0]!.name]
  ])('a native %s of the repository that does not exist fails the full repository scan, by name', async (_kind, missing) => {
    const result = await check('--npm', (name) =>
      name === missing ? notFound() : versions(workspaceVersions.get(name) ?? platformVersions.get(name)!));
    expect(result.status).toBe(1);
    expect(missingLines(result.stderr)).toEqual([`${missing} is not on npm.`]);
    expect(result.stderr).toContain('RELEASING.md, "Native packages"');
  }, FULL_SCAN_BUDGET_MS);
});

describe('workspace dependency versions', () => {
  test('holds agents when its new native versions are missing from npm', async () => {
    expect(natives.map(({ name }) => name).sort()).toEqual([
      '@betteroffice/docx-native', '@betteroffice/pptx-native', '@betteroffice/xlsx-native'
    ]);
    const result = await check('--npm', older, repository);
    expect(result.stderr).toBe('');
    expect(result.status).toBe(0);
    expect(result.held).toEqual([agents]);
  });

  test('fails when agents would publish before its new native versions', async () => {
    const without = await check('--npm', older, { ...repository, holdNative: false });
    expect(without.status).toBe(1);
    expect(without.held).toEqual([]);
    for (const { name, version } of natives) {
      expect(without.stderr).toContain(`${releaseTag(agents)} depends on ${name}@^${version}`);
    }
  });

  test('allows agents to publish when the native manifest versions are already on npm', async () => {
    const result = await check('--npm', (name) =>
      natives.some((native) => native.name === name) ? published(name) : older(name), repository);
    expect(result.status).toBe(0);
    expect(result.held).toEqual([]);
  });

  test('keeps agents held when the release skips native packages', async () => {
    const prior = process.env.RELEASE_SKIP_NATIVE;
    process.env.RELEASE_SKIP_NATIVE = 'true';
    try {
      const skipped = await check('--npm', older, repository);
      expect(skipped.status).toBe(0);
      expect(skipped.held).toEqual([agents]);
    } finally {
      if (prior === undefined) delete process.env.RELEASE_SKIP_NATIVE;
      else process.env.RELEASE_SKIP_NATIVE = prior;
    }
  });

  test('holding applies only to unpublished versions', async () => {
    const result = await check('--npm', (name) => name === agents.name ? published(name) : older(name), repository);
    expect(result.status).toBe(0);
    expect(result.held).toEqual([]);
  });

  test.each(['peerDependencies', 'optionalDependencies'])('holds native dependencies listed in %s', async (field) => {
    const native = natives[0]!;
    const result = await check('--npm', () => versions('0.0.0'), {
      packages: [agents], natives: [native], later: [], platforms: [],
      workspaces: workspaceFixture([{ ...agents, [field]: { [native.name]: 'workspace:*' } }, native])
    });
    expect(result.status).toBe(0);
    expect(result.held).toEqual([agents]);
  });

  test('fails when an immediate publish would depend on a held workspace version', async () => {
    const native = natives[0]!;
    const result = await check('--npm', () => versions('0.0.0'), {
      packages: [agents, fonts], natives: [native], later: [], platforms: [],
      workspaces: workspaceFixture([
        { ...agents, dependencies: { [native.name]: 'workspace:*' } },
        { ...fonts, dependencies: { [agents.name]: 'workspace:*' } },
        native
      ])
    });
    expect(result.status).toBe(1);
    expect(result.held).toEqual([agents]);
    expect(result.stderr).toContain(`${releaseTag(fonts)} depends on ${releaseTag(agents)}`);
  });

  test('the full repository scan holds the same packages as the precomputed scan', async () => {
    const scanned = await check('--npm', older);
    expect(scanned.stderr).toBe('');
    expect(scanned.status).toBe(0);
    const precomputed = await check('--npm', older, repository);
    expect(scanned.held).toEqual(precomputed.held);
    expect([precomputed.status, precomputed.stdout, precomputed.stderr]).toEqual([scanned.status, scanned.stdout, scanned.stderr]);
  }, FULL_SCAN_BUDGET_MS);

  for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
    test.each([
      ['workspace:*', latin.version],
      ['workspace:^', `^${latin.version}`],
      ['workspace:~', `~${latin.version}`]
    ])(`rejects an unavailable %s version in ${field}`, async (range, resolved) => {
      const result = await check('--npm', () => versions('0.0.0'), {
        packages: [fonts], workspaces: fontWorkspaces(field, range)
      });
      expect(result.status).toBe(1);
      expect(result.stderr).toContain(
        `${fonts.name}@${fonts.version} depends on ${latin.name}@${resolved}, which is neither in this release nor on npm`
      );
    });
  }

  test('accepts a dependency version in the release even when npm only has an older version', async () => {
    const result = await check('--npm', () => versions('0.0.0'), {
      packages: [fonts, latin], workspaces: fontWorkspaces()
    });
    expect(result.status).toBe(0);
    expect(result.stderr).toBe('');
  });

  test.each([
    ['workspace:*', latin.version],
    ['workspace:^0', '0.9.0'],
    ['workspace:^', '1.3.0'],
    ['workspace:~', '1.2.4'],
    ['workspace:1.2.3', latin.version],
    ['^1.0.0', '1.3.0'],
    ['>=1.2.0 <2.0.0', '1.4.0'],
    ['^0.1.0 || ~1.2.0', '1.2.4']
  ])('accepts %s when npm has a matching version', async (range, version) => {
    const result = await check('--npm', (name) => versions(name === latin.name ? version : '0.0.0'), {
      packages: [fonts], workspaces: fontWorkspaces('dependencies', range)
    });
    expect(result.status).toBe(0);
    expect(result.stderr).toBe('');
  });

  test('rejects a release dependency whose version does not satisfy a plain range', async () => {
    const result = await check('--npm', () => versions('0.0.0'), {
      packages: [fonts, latin], workspaces: fontWorkspaces('dependencies', '^2.0.0')
    });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`depends on ${latin.name}@^2.0.0`);
  });

  test('still rejects an unclaimed sibling name included in the release', async () => {
    const result = await check('--npm', (name) => name === latin.name ? notFound() : versions('0.0.0'), {
      packages: [fonts, latin], workspaces: fontWorkspaces()
    });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${latin.name} is not on npm.`);
    expect(result.stderr).toContain('Initial npm release of a new package');
  });

  test('ignores development dependencies and dependencies outside this workspace', async () => {
    const workspaces = workspaceFixture([
      { ...fonts, dependencies: { external: '^1' }, devDependencies: { [latin.name]: 'workspace:*' } },
      latin
    ]);
    const result = await check('--npm', () => versions('0.0.0'), { packages: [fonts], workspaces });
    expect(result.status).toBe(0);
  });
});

describe('sibling package references', () => {
  test('finds all three real cdn.ts sibling imports with their source lines', () => {
    const references = siblingReferences(packages.filter(({ name }) => name === fonts.name));
    const cdn = references.filter(({ file }) => file === 'packages/fonts/src/cdn.ts');
    expect(cdn.map(({ name, line }) => [name, line])).toEqual([
      ['@betteroffice/fonts-cjk', 2], [latin.name, 3], ['@betteroffice/fonts-intl', 4]
    ]);
    for (const reference of cdn) {
      expect(reference.dependent).toBe(fonts.name);
      expect(reference.version).toBe(packages.find(({ name }) => name === reference.name)!.version);
    }
  });

  test('scans all source extensions and dynamic imports while excluding tests and non-code references', () => {
    const workspaces = workspaceFixture([fonts, latin]);
    for (const extension of ['ts', 'tsx', 'js', 'mjs']) {
      writeFileSync(join(workspaces[0]!.directory, 'src', `index.${extension}`), [
        '// import { version } from "../../fonts-latin/package.json";',
        '/* from "../../fonts-latin/package.json" */',
        '`from "../../fonts-latin/package.json"`;',
        "import { version } from '../../fonts-latin/package.json';",
        'const sibling = import(/* sibling */ "../../fonts-latin/package.json");',
        'import "../../fonts-latin/package.json";',
        'import { version as own } from "../package.json";',
        'const pattern = /from "..\\/..\\/fonts-latin\\/package.json"/;',
        'const dynamic = import(`../../fonts-latin/package.json`);'
      ].join('\n'));
      writeFileSync(join(workspaces[0]!.directory, 'src', `index.test.${extension}`),
        'import "../../fonts-latin/package.json";');
    }
    mkdirSync(join(workspaces[0]!.directory, 'src', 'nested'));
    writeFileSync(join(workspaces[0]!.directory, 'src', 'nested', 'index.ts'),
      'const sibling = import("../../../fonts-latin/package.json");');
    writeFileSync(join(workspaces[0]!.directory, 'build.mjs'), 'import "../fonts-latin/package.json";');
    writeFileSync(join(workspaces[0]!.directory, 'src', 'view.tsx'),
      'const view = <span>from "../../fonts-latin/package.json"</span>;');
    const references = siblingReferences([fonts], workspaces);
    expect(references).toHaveLength(17);
    expect(references.filter(({ file }) => !file.endsWith('nested/index.ts')).map(({ line }) => line))
      .toEqual([4, 5, 6, 9, 4, 5, 6, 9, 4, 5, 6, 9, 4, 5, 6, 9]);
    expect(references.every(({ name, version }) => name === latin.name && version === latin.version)).toBe(true);
  });

  test.each(['release', 'npm'])('accepts an exact sibling reference available in %s', async (available) => {
    const workspaces = workspaceFixture([fonts, latin]);
    writeFileSync(join(workspaces[0]!.directory, 'src', 'index.ts'),
      'const sibling = import("../../fonts-latin/package.json");');
    const result = await check('--npm', (name) =>
      versions(available === 'npm' && name === latin.name ? latin.version : '0.0.0'), {
      packages: available === 'release' ? [fonts, latin] : [fonts], workspaces
    });
    expect(result.status).toBe(0);
    expect(result.stderr).toBe('');
  });

  test('rejects a sibling reference when the release contains a different version', async () => {
    const workspaces = workspaceFixture([fonts, latin]);
    writeFileSync(join(workspaces[0]!.directory, 'src', 'index.ts'),
      'import { version } from "../../fonts-latin/package.json";');
    const result = await check('--npm', () => versions('0.0.0'), {
      packages: [fonts, { ...latin, version: '1.2.4' }], workspaces
    });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`builds references to ${latin.name}@${latin.version}`);
  });

  test('rejects the real fonts references when fonts-latin is outside the release and absent at its version', async () => {
    const result = await check('--npm', (name) => name === latin.name ? versions('0.0.0') : published(name), {
      packages: packages.filter(({ name }) => name === fonts.name)
    });
    const version = packages.find(({ name }) => name === latin.name)!.version;
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(
      `${fonts.name} builds references to ${latin.name}@${version} (packages/fonts/src/cdn.ts:3), which is neither in this release nor on npm`
    );
  });

  test('rejects sibling references independently of declared dependencies', async () => {
    const workspaces = workspaceFixture([fonts, latin]);
    writeFileSync(join(workspaces[0]!.directory, 'src', 'index.ts'),
      'import { version } from "../../fonts-latin/package.json";');
    const result = await check('--npm', () => versions('0.0.0'), { packages: [fonts], workspaces });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`builds references to ${latin.name}@${latin.version}`);
    expect(result.stderr).not.toContain('depends on');
  });

  test('rejects current manifests minus fonts-latin when npm only has 0.0.0 of latin and intl', async () => {
    const result = await check('--npm', (name) =>
      [fonts.name, latin.name, '@betteroffice/fonts-intl'].includes(name) ? versions('0.0.0') : published(name), {
      ...repository,
      packages: packages.filter(({ name }) => name !== latin.name),
      references: repository.references.filter(({ dependent }) => dependent !== latin.name)
    });
    const version = packages.find(({ name }) => name === latin.name)!.version;
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${fonts.name}@${version} depends on ${latin.name}@${version}`);
    expect(result.stderr).toContain(`${fonts.name} builds references to ${latin.name}@${version}`);
  });
});

describe('published npm and CDN targets', () => {
  test('reports held versions without checking them on npm or the CDN', async () => {
    const workspaces = workspaceFixture([fonts, agents]);
    writeFileSync(join(workspaces[0]!.directory, 'src', 'index.ts'),
      'import { version } from "../../agents/package.json";');
    const requested: string[] = [];
    const result = await check('--verify-published', (name, request) => {
      requested.push(name);
      return request.method === 'HEAD' ? new Response(null) : Response.json(fonts);
    }, { packages: [fonts, agents], held: [agents], workspaces, maxWaitMs: 500 });
    expect(result.status).toBe(0);
    expect(result.stdout).toContain(`${releaseTag(agents)} is held for its native dependencies.`);
    expect(requested.sort()).toEqual([`${fonts.name}/${fonts.version}`, `${releaseTag(fonts)}/package.json`].sort());
  });

  test('reports every missing npm version using version-specific GETs', async () => {
    const requested: string[] = [];
    const result = await check('--verify-published', (name, request) => {
      if (request.method === 'HEAD') return new Response(null);
      requested.push(name);
      return name === `${latin.name}/${latin.version}` ? notFound() : Response.json(fonts);
    }, { packages: [fonts, latin], workspaces: workspaceFixture([fonts, latin]), maxWaitMs: 500 });
    expect(result.status).toBe(1);
    expect(requested).toContain(`${fonts.name}/${fonts.version}`);
    expect(requested).toContain(`${latin.name}/${latin.version}`);
    expect(result.stderr).toBe(`${latin.name}@${latin.version} is not on npm after publishing.`);
  });

  test('retries CDN HEAD 404s until the package and its sibling references propagate', async () => {
    const workspaces = workspaceFixture([fonts, latin]);
    writeFileSync(join(workspaces[0]!.directory, 'src', 'index.ts'), [
      'import { version } from "../../fonts-latin/package.json";',
      'const again = import("../../fonts-latin/package.json");'
    ].join('\n'));
    const heads = new Map<string, number>();
    const result = await check('--verify-published', (name, request) => {
      if (request.method === 'GET') return Response.json(fonts);
      expect(request.method).toBe('HEAD');
      const call = (heads.get(name) ?? 0) + 1;
      heads.set(name, call);
      return call === 1 ? notFound() : new Response(null);
    }, { packages: [fonts], workspaces, maxWaitMs: 2_000 });
    expect(result.status).toBe(0);
    expect(result.stderr).toBe('');
    expect([...heads.entries()].sort()).toEqual([
      [`${latin.name}@${latin.version}/package.json`, 2],
      [`${fonts.name}@${fonts.version}/package.json`, 2]
    ].sort());
    expect(result.elapsed).toBeGreaterThanOrEqual(200);
  });

  test('names CDN versions that stay missing within a shared retry budget', async () => {
    const result = await check('--verify-published', (_name, request) =>
      request.method === 'HEAD' ? notFound() : Response.json(fonts), {
      packages: [fonts, latin], workspaces: workspaceFixture([fonts, latin]), maxWaitMs: 500
    });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${fonts.name}@${fonts.version} is not on the CDN after publishing.`);
    expect(result.stderr).toContain(`${latin.name}@${latin.version} is not on the CDN after publishing.`);
    expect(result.elapsed).toBeLessThan(2_000);
  });

  test('retries npm version propagation and Retry-After responses', async () => {
    let calls = 0;
    const result = await check('--verify-published', (_name, request) => {
      if (request.method === 'HEAD') return new Response(null);
      calls++;
      if (calls === 1) return notFound();
      if (calls === 2) return new Response(null, { status: 429, headers: { 'Retry-After': '0' } });
      return Response.json(fonts);
    }, { packages: [fonts], workspaces: workspaceFixture([fonts]), maxWaitMs: 2_000 });
    expect(result.status).toBe(0);
    expect(calls).toBe(3);
  });
});

describe('publishing held npm packages', () => {
  test('rechecks only the required native versions and rejects an older version', async () => {
    const missing = natives[0]!;
    const requested: string[] = [];
    const { server, origin } = registry((path) => {
      requested.push(path);
      return path === `${missing.name}/${missing.version}` ? notFound() : Response.json({});
    });
    try {
      expect(await missingHeldNatives([agents], { registry: origin })).toEqual([missing]);
      expect(requested.sort()).toEqual(natives.map(({ name, version }) => `${name}/${version}`).sort());
    } finally {
      server.stop(true);
    }
  });

  test('waits for a native version that has just been published', async () => {
    const native = natives[0]!;
    const path = `${native.name}/${native.version}`;
    const response = flaky((name, call) => name === path && call === 1 ? notFound() : Response.json({}));
    const { server, origin } = registry(response.respond);
    try {
      expect(await missingHeldNatives([agents], { registry: origin, wait: true, maxWaitMs: 2_000 })).toEqual([]);
      expect(response.calls.get(path)).toBe(2);
    } finally {
      server.stop(true);
    }
  });

  test('uses the scoped package name and full version for the tag and release title', () => {
    expect(releaseTag({ name: '@betteroffice/agents', version: '0.3.0' })).toBe('@betteroffice/agents@0.3.0');
    expect(releaseTag({ name: '@betteroffice/agents', version: '0.4.0-next.1' })).toBe('@betteroffice/agents@0.4.0-next.1');
  });

  test('uses only the matching changelog section for the release notes', () => {
    const notes = ['### Patch Changes', '', '- Fix an issue.', '', '```md', '## Example', '```'].join('\n');
    const changelog = ['# @betteroffice/agents', '## 0.3.0', '- Newer notes.', '## 0.2.0', notes, '## 0.1.0', '- Older notes.'].join('\r\n');
    expect(releaseNotes(changelog, '0.2.0')).toBe(notes);
    expect(releaseNotes(changelog, '0.1.0')).toBe('- Older notes.');
    expect(releaseNotes(changelog, '0.2')).toBe('');
  });
});

describe('crates.io publish targets', () => {
  test('the release list matches publishable manifests and excludes VSDX', () => {
    expect([...crates].sort()).toEqual(publishedCrates());
    expect(crates.some((name) => name.includes('vsdx'))).toBe(false);
  });

  test('a crate that exists passes', async () => {
    const result = await check('--crates', () => Response.json({}));
    expect(result.status).toBe(0);
    expect(result.stdout).toContain(`${crates[0]} is on crates.io.`);
  });

  test('a crate that does not exist fails, by name', async () => {
    const result = await check('--crates', (name) =>
      name === crates[0] ? notFound() : Response.json({})
    );
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${crates[0]} is not on crates.io.`);
    expect(result.stderr).toContain('Initial crates.io release');
  });

  test('a bootstrap token creates missing crates, so a missing one only prints', async () => {
    const result = await check(
      '--crates',
      (name) => (name === crates[0] ? notFound() : Response.json({})),
      { bootstrap: true }
    );
    expect(result.status).toBe(0);
    expect(result.stdout).toContain(
      `${crates[0]} is not on crates.io; will be created with the bootstrap token.`
    );
    expect(result.stderr).toBe('');
  });
});

describe('release wiring', () => {
  const steps = release.jobs.release.steps.map((step: any) => step.name ?? step.uses);
  const guards = release.jobs.release.steps.filter((step: any) =>
    step.run?.includes('check-publish-targets.mjs')
  );

  test('each guard runs only on its publish path', () => {
    expect(guards.map((step: any) => step.run.trim())).toEqual([
      'node scripts/check-publish-targets.mjs --crates',
      'node scripts/check-publish-targets.mjs --npm',
      'node scripts/check-publish-targets.mjs --plan',
      'node scripts/check-publish-targets.mjs --verify-published'
    ]);
    expect(guards.map((step: any) => step.if)).toEqual([
      "steps.pending.outputs.publishing == 'true' && vars.RELEASE_SKIP_CRATES != 'true'",
      "steps.pending.outputs.publishing == 'true'",
      "steps.pending.outputs.publishing != 'true'",
      "always() && steps.pending.outputs.publishing == 'true'"
    ]);
  });

  test('npm and CDN verification runs after publishing and before deployment', () => {
    const verify = steps.indexOf('Verify npm and CDN publish targets');
    expect(verify).toBeGreaterThan(steps.indexOf('Release PR or publish'));
    expect(verify).toBeLessThan(steps.indexOf('Trigger deploy'));
  });

  test('packs pinned manifests before Changesets and hands the held list to the later job', () => {
    const pack = steps.indexOf('Pack and hold packages that need native versions');
    expect(pack).toBeGreaterThan(steps.indexOf('Pin workspace deps for npm publish'));
    expect(pack + 1).toBe(steps.indexOf('Release PR or publish'));
    expect(release.jobs.release.outputs.held).toBe('${{ steps.npm-targets.outputs.held }}');
    const held = release.jobs['held-npm'];
    expect(held.needs).toEqual(['release', 'native-npm']);
    expect(held.if).toContain("needs.release.result == 'success' && needs.release.outputs.publishing == 'true'");
    expect(held.if).toContain("needs.release.outputs.held != '' && needs.release.outputs.held != '[]'");
    expect(held.if).toContain("needs.native-npm.result == 'success' || needs.native-npm.result == 'skipped'");
    expect(held.permissions).toEqual({ contents: 'write', 'id-token': 'write' });
  });

  test('a verification miss reports without gating deploy, native or Python publishing', () => {
    const step = (name: string) => release.jobs.release.steps.find((s: any) => s.name === name);
    const verify = step('Verify npm and CDN publish targets');
    expect(verify['continue-on-error']).toBe(true);
    expect(verify.id).toBeUndefined();
  });

  test('a failed publish still gates deploy, native and Python publishing', () => {
    const publish = release.jobs.release.steps.find((s: any) => s.name === 'Release PR or publish');
    expect(publish['continue-on-error']).toBeUndefined();
    const deploy = release.jobs.release.steps.find((s: any) => s.name === 'Trigger deploy');
    expect(deploy.if).toBe("steps.changesets.outputs.published == 'true'");
    expect(release.jobs['native-npm'].needs).toBe('release');
    expect(release.jobs['native-npm'].if).toBe(
      "needs.release.result == 'success' && needs.release.outputs.publishing == 'true' && vars.RELEASE_SKIP_NATIVE != 'true'"
    );
    expect(release.jobs['python-bindings'].needs).toBe('release');
    expect(release.jobs['python-bindings'].if).toBe(
      "needs.release.result == 'success' && needs.release.outputs.publishing == 'true' && vars.RELEASE_SKIP_PYTHON != 'true'"
    );
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

describe('crates OIDC-first wiring', () => {
  const named = new Map(release.jobs.release.steps.map((step: any) => [step.name, step]));

  test('the OIDC exchange runs on the publish path unless crates are skipped', () => {
    const auth = named.get('Authenticate to crates.io');
    expect(auth.uses).toMatch(/^rust-lang\/crates-io-auth-action@[0-9a-f]{40}$/);
    expect(auth.if).toBe(
      "steps.pending.outputs.publishing == 'true' && vars.RELEASE_SKIP_CRATES != 'true'"
    );
  });

  test('an OIDC failure continues only when the bootstrap token is detected', () => {
    const auth = named.get('Authenticate to crates.io');
    expect(auth['continue-on-error']).toBe(
      "${{ steps.crates-bootstrap.outputs.enabled == 'true' }}"
    );
  });

  test('the crates publish gets the OIDC token and the bootstrap token separately', () => {
    const publish = named.get('Publish Rust crates');
    expect(publish.env.CARGO_REGISTRY_TOKEN).toBe('${{ steps.crates-auth.outputs.token }}');
    expect(publish.env.CRATES_IO_BOOTSTRAP_TOKEN).toBe('${{ secrets.CRATES_IO_BOOTSTRAP_TOKEN }}');
  });

  test('the release job can mint OIDC tokens', () => {
    expect(release.jobs.release.permissions['id-token']).toBe('write');
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
    const result = await check('--npm', registry.respond, repository);

    expect(result.status).toBe(0);
    expect(registry.calls.get(first)).toBe(2);
    expect(result.elapsed).toBeGreaterThan(900);
  });

  test('gives up on a request that hangs, and retries it', async () => {
    const registry = flaky((name, call) =>
      name === first && call === 1 ? new Promise<Response>(() => {}) : published(name)
    );
    const result = await check('--npm', registry.respond, { ...repository, timeoutMs: 400 });

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
    const result = await check('--npm', registry.respond, repository);

    expect(result.status).toBe(0);
    expect(registry.calls.get(first)).toBe(2);
  });

  test('stops after five attempts instead of publishing on a guess', async () => {
    const registry = flaky(
      () =>
        new Response('{"error":"unavailable"}', { status: 503, headers: { 'Retry-After': '0' } })
    );
    await expect(check('--npm', registry.respond, repository)).rejects.toThrow('503');
    expect(registry.calls.get(first)).toBe(5);
  });

  test('a 4xx that is not a rate limit is not retried', async () => {
    const registry = flaky(() => new Response('{"error":"gone"}', { status: 410 }));
    await expect(check('--npm', registry.respond, repository)).rejects.toThrow('410');
    expect(registry.calls.get(first)).toBe(1);
  });
});

describe('the mode argument', () => {
  test('an unknown mode fails instead of checking nothing', async () => {
    const result = await guard('--npmm', () => Response.json({}));
    expect(result.status).toBe(2);
    expect(result.stderr).toContain('expected exactly one of --npm, --crates, --verify-published or --plan');
  }, SPAWN_BUDGET_MS);

  test('--npm checks the registry NPM_REGISTRY_URL names and exits with the result', async () => {
    const missing = packages[0]!.name;
    const result = await guard('--npm', (name) => (name === missing ? notFound() : published(name)));
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${missing} is not on npm.`);
  }, SPAWN_BUDGET_MS);

  test('--npm scans the repository itself and reports what the precomputed scan reports', async () => {
    const directory = mkdtempSync(join(tmpdir(), 'publish-output-'));
    fixtureDirectories.push(directory);
    const output = join(directory, 'output');
    const cli = await guard('--npm', older, {
      GITHUB_OUTPUT: output, RELEASE_SKIP_NATIVE: 'true', NODE_OPTIONS: '', NODE_NO_WARNINGS: '1'
    });
    const prior = process.env.RELEASE_SKIP_NATIVE;
    process.env.RELEASE_SKIP_NATIVE = 'true';
    try {
      const precomputed = await check('--npm', older, repository);
      expect(cli.status).toBe(0);
      expect(readFileSync(output, 'utf8')).toBe(`held=${JSON.stringify([agents])}\n`);
      expect([cli.status, cli.stdout.trimEnd(), cli.stderr.trimEnd()])
        .toEqual([precomputed.status, precomputed.stdout, precomputed.stderr]);
    } finally {
      if (prior === undefined) delete process.env.RELEASE_SKIP_NATIVE;
      else process.env.RELEASE_SKIP_NATIVE = prior;
    }
  }, SPAWN_BUDGET_MS);

  test('--verify-published uses the registry and CDN environment overrides', async () => {
    const gets: string[] = [];
    const heads: string[] = [];
    const result = await guard('--verify-published', (name, request) => {
      (request.method === 'HEAD' ? heads : gets).push(name);
      return request.method === 'HEAD' ? new Response(null) : Response.json({});
    });
    expect(result.status).toBe(0);
    expect(gets.sort()).toEqual(packages.map(({ name, version }) => `${name}/${version}`).sort());
    expect(heads.sort()).toEqual(packages.map(({ name, version }) => `${name}@${version}/package.json`).sort());
  }, SPAWN_BUDGET_MS);

  test('--verify-published on GitHub Actions annotates and summarizes what is missing', async () => {
    const directory = mkdtempSync(join(tmpdir(), 'publish-summary-'));
    fixtureDirectories.push(directory);
    const summary = join(directory, 'summary.md');
    const { name, version } = packages[0]!;
    const result = await guard('--verify-published', (path, request) =>
      request.method === 'HEAD'
        ? new Response(null)
        : path === `${name}/${version}` ? new Response(null, { status: 403 }) : Response.json({}),
    { GITHUB_ACTIONS: 'true', GITHUB_STEP_SUMMARY: summary });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`::error title=Release incomplete::${name}@${version} is not on npm after publishing`);
    const written = readFileSync(summary, 'utf8');
    expect(written).toContain('### Release incomplete');
    expect(written).toContain(`- \`${name}@${version}\` on npm`);
    expect(written).not.toContain('on the CDN');
  }, SPAWN_BUDGET_MS);
});

describe('release plan', () => {
  const fixture = (peer: string) => [
    { directory: '/fixture/fonts', manifest: { name: '@betteroffice/fonts', version: '0.4.2', peerDependencies: { '@betteroffice/fonts-cjk': peer } } },
    { directory: '/fixture/fonts-cjk', manifest: { name: '@betteroffice/fonts-cjk', version: '0.4.2' } }
  ];
  const plan = (type: 'minor' | 'major') => ({
    changesets: [{ id: 'a', summary: '', releases: [{ name: '@betteroffice/fonts-cjk', type: 'minor' }, { name: '@betteroffice/fonts', type: 'minor' }] }],
    releases: [
      { name: '@betteroffice/fonts', type, oldVersion: '0.4.2', newVersion: type === 'major' ? '1.0.0' : '0.5.0', changesets: ['a'] },
      { name: '@betteroffice/fonts-cjk', type, oldVersion: '0.4.2', newVersion: type === 'major' ? '1.0.0' : '0.5.0', changesets: ['a'] }
    ],
    preState: undefined
  });

  test('a major bump no changeset asked for is named with the peer range that forced it', async () => {
    const majors = unaskedMajors(plan('major'), fixture('workspace:*'), []);
    expect(majors.map(({ name }) => name)).toEqual(['@betteroffice/fonts', '@betteroffice/fonts-cjk']);
    expect(majors[0]!.causes).toEqual(['peer @betteroffice/fonts-cjk@workspace:* does not cover its next version 1.0.0']);
    expect(majors[1]!.causes).toEqual([]);
    const lines: string[] = [];
    expect(await checkPlan({ plan: plan('major'), workspaces: fixture('workspace:*'), fixed: [], out: { log: () => {}, error: (line: string) => lines.push(line) } })).toBe(1);
    expect(lines[0]).toContain('@betteroffice/fonts 0.4.2 -> 1.0.0 is a major bump no changeset asked for: peer @betteroffice/fonts-cjk@workspace:*');
    expect(lines[1]).toContain('(raised with its fixed group)');
  });

  test('minor bumps inside a caret peer range pass', async () => {
    expect(unaskedMajors(plan('minor'), fixture('workspace:^0'), [])).toEqual([]);
  });

  test('a major asked for one member of a fixed group covers the whole group', async () => {
    const asked = plan('major');
    asked.changesets[0]!.releases[0]!.type = 'major';
    expect(unaskedMajors(asked, fixture('workspace:*'), [['@betteroffice/fonts', '@betteroffice/fonts-cjk']])).toEqual([]);
    expect(unaskedMajors(plan('major'), fixture('workspace:*'), [['@betteroffice/fonts', '@betteroffice/fonts-cjk']]).map(({ name }) => name))
      .toEqual(['@betteroffice/fonts', '@betteroffice/fonts-cjk']);
  });
});
