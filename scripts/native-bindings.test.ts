import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  NATIVE_PACKAGES,
  NATIVE_FORMATS,
  bindingVersion,
  pendingPublishNames,
  platformPackageVersions,
  validateNativeVersions
} from './native-bindings.mjs';

const releaseWorkflow = fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url));
const publishWorkflow = fileURLToPath(
  new URL('../.github/workflows/publish-native.yml', import.meta.url)
);
const distWorkflow = fileURLToPath(new URL('../.github/workflows/native-dist.yml', import.meta.url));

describe('native package registry', () => {
  const roots = NATIVE_PACKAGES.map((path) =>
    JSON.parse(readFileSync(new URL(`../${path}/package.json`, import.meta.url), 'utf8'))
  );
  const packages = [...roots, ...platformPackageVersions()];
  const versions = new Map<string, string>(packages.map(({ name, version }) => [name, version]));
  const packageName = (url: string) => decodeURIComponent(new URL(url).pathname.slice(1));
  const published = (name: string) => Response.json({ versions: { [versions.get(name)!]: {} } });

  test('registers every native format and platform package', () => {
    expect(NATIVE_FORMATS).toEqual(['docx', 'pptx', 'xlsx']);
    expect(NATIVE_PACKAGES).toEqual(NATIVE_FORMATS.map((format) => `bindings/${format}-native`));
    const platforms = [
      'darwin-arm64',
      'darwin-x64',
      'linux-arm64-gnu',
      'linux-x64-gnu',
      'win32-x64-msvc'
    ];
    expect(
      platformPackageVersions()
        .map(({ name }) => name)
        .sort()
    ).toEqual(
      NATIVE_FORMATS.flatMap((format) =>
        platforms.map((platform) => `@betteroffice/${format}-native-${platform}`)
      ).sort()
    );
    expect(platformPackageVersions()).toHaveLength(NATIVE_PACKAGES.length * 5);
    expect(() => validateNativeVersions()).not.toThrow();
  });

  test('keeps roots private and platform packages public', () => {
    for (const path of NATIVE_PACKAGES) {
      const root = JSON.parse(
        readFileSync(new URL(`../${path}/package.json`, import.meta.url), 'utf8')
      );
      expect(root.private).toBe(true);
      for (const identity of readdirSync(new URL(`../${path}/npm`, import.meta.url))) {
        const platform = JSON.parse(
          readFileSync(new URL(`../${path}/npm/${identity}/package.json`, import.meta.url), 'utf8')
        );
        expect(platform.private).not.toBe(true);
        expect(platform.publishConfig.access).toBe('public');
      }
    }
  });

  test('detects only unpublished root versions', async () => {
    const requested: string[] = [];
    const pending = await pendingPublishNames({
      fetchImpl: async (url: string, options: RequestInit) => {
        const name = packageName(url);
        requested.push(name);
        expect(options.headers).toEqual({ Accept: 'application/vnd.npm.install-v1+json' });
        expect(options.signal).toBeInstanceOf(AbortSignal);
        return name === '@betteroffice/pptx-native' || !roots.some((root) => root.name === name)
          ? Response.json({ versions: {} })
          : published(name);
      }
    });
    expect(pending).toEqual(['pptx']);
    expect(requested.sort()).toEqual(packages.map(({ name }) => name).sort());
  });

  test('returns no formats when every root version is published', async () => {
    expect(
      await pendingPublishNames({
        fetchImpl: async (url: string) => published(packageName(url))
      })
    ).toEqual([]);
  });

  const missingSets: [string, string[]][] = [
    ['root name', ['@betteroffice/docx-native']],
    ['platform name', ['@betteroffice/pptx-native-linux-arm64-gnu']],
    ['root names', roots.map(({ name }) => name)],
    ['platform names', platformPackageVersions().map(({ name }) => name)],
    ['root and platform names', packages.map(({ name }) => name)]
  ];

  test.each(missingSets)('reports every missing %s', async (_label, missing) => {
    const requested: string[] = [];
    const failure = await pendingPublishNames({
      fetchImpl: async (url: string) => {
        const name = packageName(url);
        requested.push(name);
        return missing.includes(name)
          ? new Response('{"error":"missing"}', { status: 404 })
          : published(name);
      }
    }).catch((error: Error) => error);
    expect(failure).toBeInstanceOf(Error);
    expect((failure as Error).message).toBe(
      `Missing npm packages: ${requested.filter((name) => missing.includes(name)).join(', ')}. ` +
        'They must be published by hand first (RELEASING.md, "Native packages"), ' +
        'because Trusted Publishing cannot create a package.'
    );
    expect(requested.sort()).toEqual(packages.map(({ name }) => name).sort());
  });

  test('registry errors stop release detection', async () => {
    for (const failed of ['@betteroffice/docx-native', '@betteroffice/docx-native-linux-x64-gnu']) {
      await expect(
        pendingPublishNames({
          fetchImpl: async (url: string) => {
            const name = packageName(url);
            return name === failed ? new Response('unavailable', { status: 503 }) : published(name);
          }
        })
      ).rejects.toThrow(`npm answered 503 for ${failed}`);
    }
  });
});

describe('native version synchronization', () => {
  test('each native package joins its core fixed release group', () => {
    const config = JSON.parse(
      readFileSync(new URL('../.changeset/config.json', import.meta.url), 'utf8')
    );
    expect(config.privatePackages.version).toBe(true);
    for (const format of NATIVE_FORMATS) {
      const group = config.fixed.find((names: string[]) => names.includes(`@betteroffice/${format}`));
      expect(group).toContain(`@betteroffice/${format}-native`);
      expect(bindingVersion(`bindings/${format}-native`)).toBe(bindingVersion(`packages/${format}`));
    }
  });
});

describe('native package release wiring', () => {
  const release = Bun.YAML.parse(readFileSync(releaseWorkflow, 'utf8')) as any;
  const publish = Bun.YAML.parse(readFileSync(publishWorkflow, 'utf8')) as any;
  const dist = Bun.YAML.parse(readFileSync(distWorkflow, 'utf8')) as any;
  const ci = Bun.YAML.parse(
    readFileSync(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8')
  ) as any;
  const packages = Bun.YAML.parse(
    readFileSync(new URL('../.github/workflows/native-packages.yml', import.meta.url), 'utf8')
  ) as any;

  test('keeps native publishing outside the main release job', () => {
    const job = release.jobs.release;
    expect(job['timeout-minutes']).toBe(45);
    expect(JSON.stringify(job)).not.toContain('native');
    expect(JSON.stringify(job)).not.toContain('withhold');
    expect(JSON.stringify(job)).not.toContain('RELEASE_SKIP_NATIVE');
    const steps = job.steps.map((step: any) => step.name);
    expect(steps.indexOf('Authenticate to crates.io')).toBeGreaterThan(
      steps.indexOf('Detect crates.io bootstrap token')
    );
    expect(steps.indexOf('Authenticate to crates.io')).toBeLessThan(
      steps.indexOf('Check crates.io publish targets')
    );
  });

  test('release dispatches and waits for the dedicated publisher', () => {
    const job = release.jobs['native-npm'];
    expect(job.needs).toBe('release');
    expect(job.if).toBe(
      "needs.release.result == 'success' && needs.release.outputs.publishing == 'true' && vars.RELEASE_SKIP_NATIVE != 'true'"
    );
    expect(job['runs-on']).toBe('ubuntu-latest');
    expect(job['timeout-minutes']).toBe(75);
    expect(job.permissions).toEqual({ contents: 'read', actions: 'write' });
    expect(job.steps[0].uses).toBe('actions/checkout@v7');
    const step = job.steps[1];
    expect(step.env.SHA).toBe('${{ github.sha }}');
    expect(step.run).toStartWith('set -euo pipefail\n');
    expect(step.run).toContain('pending=$(node scripts/native-bindings.mjs --pending)');
    expect(step.run).toContain('publish-native.yml/dispatches');
    expect(step.run).toContain('X-GitHub-Api-Version: 2026-03-10');
    expect(step.run).toContain('-f ref=main');
    expect(step.run).toContain('inputs[binding]=$binding');
    expect(step.run).toContain('inputs[sha]=$SHA');
    expect(step.run).toContain('inputs[dry_run]=false');
    expect(step.run).toContain("--jq '.workflow_run_id'");
    expect(step.run).toContain('gh run watch "$run" --exit-status');
    expect(step.run).toContain('@betteroffice/$binding-native');
    expect(step.run).toContain('exit $status');
  });

  test('a failed registry lookup fails before dispatch', () => {
    const directory = mkdtempSync(join(tmpdir(), 'native-dispatch-'));
    try {
      writeFileSync(join(directory, 'node'), '#!/bin/sh\nexit 42\n', { mode: 0o755 });
      writeFileSync(join(directory, 'gh'), '#!/bin/sh\necho unexpected-dispatch\nexit 99\n', {
        mode: 0o755
      });
      const step = release.jobs['native-npm'].steps[1];
      const result = spawnSync('bash', ['-c', step.run], {
        encoding: 'utf8',
        env: { ...process.env, PATH: `${directory}:${process.env.PATH}` }
      });
      expect(result.status).toBe(42);
      expect(result.stdout).not.toContain('unexpected-dispatch');
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test('publisher is dispatch-only and uses token-free OIDC', () => {
    expect(Object.keys(publish.on)).toEqual(['workflow_dispatch']);
    expect(publish.on.workflow_dispatch.inputs.dry_run.default).toBe(true);
    expect(publish.name).toBe('Publish a native package');
    expect(publish['run-name']).toContain('native package');
    const job = publish.jobs.publish;
    expect(job.permissions['id-token']).toBe('write');
    expect(job.environment).toBe('npm-${{ inputs.binding }}');
    const refusal = job.steps.find((step: any) => step.name === 'Refuse an npm token');
    expect(refusal.env.TOKEN).toBe('${{ secrets.NPM_TOKEN }}');
    expect(refusal.run).toContain('[ -n "$TOKEN" ]');
    expect(refusal.run).toContain('NPM_TOKEN');
    expect(refusal.run).toContain('exit 1');
    const validate = job.steps.find((step: any) => step.name === 'Validate packages');
    const upload = job.steps.find((step: any) => step.name === 'Publish packages');
    expect(validate.if).toBe('inputs.dry_run');
    expect(validate.run).toContain('--dry-run --no-gh-release');
    expect(upload.if).toBe('${{ !inputs.dry_run }}');
    expect(upload.run).toBe('npm publish --access public');
  });

  test('publisher downloads the loader before assembling and checking packages', () => {
    const steps = publish.jobs.publish.steps;
    const download = steps.find(
      (step: any) =>
        step.uses === 'actions/download-artifact@v4' &&
        step.with.pattern === '${{ inputs.binding }}-native.*'
    );
    expect(download.with).toEqual({
      pattern: '${{ inputs.binding }}-native.*',
      path: 'bindings/${{ inputs.binding }}-native/artifacts',
      'merge-multiple': true
    });
    const loader = steps.find(
      (step: any) =>
        step.uses === 'actions/download-artifact@v4' &&
        step.with.name === '${{ inputs.binding }}-native-loader'
    );
    expect(loader.with).toEqual({
      name: '${{ inputs.binding }}-native-loader',
      path: 'bindings/${{ inputs.binding }}-native'
    });
    expect(steps.indexOf(loader)).toBe(steps.indexOf(download) + 1);
    expect(steps.indexOf(loader)).toBeLessThan(
      steps.findIndex((step: any) => step.name === 'Assemble packages')
    );
    expect(steps.indexOf(loader)).toBeLessThan(
      steps.findIndex((step: any) => step.name === 'Check root contents')
    );
  });

  test('publisher checks root contents after dropping private and before publishing', () => {
    const steps = publish.jobs.publish.steps;
    const makePublic = steps.find((step: any) => step.name === 'Make the root publishable');
    expect(makePublic.run).toBe('npm pkg delete private');
    expect(makePublic['working-directory']).toBe('bindings/${{ inputs.binding }}-native');
    expect(makePublic.if).toBeUndefined();
    const contents = steps.find((step: any) => step.name === 'Check root contents');
    expect(contents.run).toBe(
      `npm pack --dry-run --json | jq -e '["index.js", "index.d.ts"] - [.[0].files[].path] == []'`
    );
    expect(contents['working-directory']).toBe('bindings/${{ inputs.binding }}-native');
    expect(contents.if).toBeUndefined();
    expect(steps.indexOf(contents)).toBe(steps.indexOf(makePublic) + 1);
    expect(steps.indexOf(contents)).toBeLessThan(
      steps.findIndex((step: any) => step.name === 'Validate packages')
    );
    expect(steps.indexOf(contents)).toBeLessThan(
      steps.findIndex((step: any) => step.name === 'Publish packages')
    );
  });

  test('dist builds exactly the declared targets at the requested commit', () => {
    const job = dist.jobs.bindings;
    const targets = job.strategy.matrix.platform.map((platform: any) => platform.target);
    expect(targets).toHaveLength(5);
    for (const path of NATIVE_PACKAGES) {
      const manifest = JSON.parse(
        readFileSync(new URL(`../${path}/package.json`, import.meta.url), 'utf8')
      );
      expect([...targets].sort()).toEqual([...manifest.napi.targets].sort());
    }
    expect(job.steps[0].with.ref).toBe('${{ inputs.sha }}');
    expect(publish.jobs.dist.uses).toBe('./.github/workflows/native-dist.yml');
    expect(publish.jobs.dist.with).toEqual({
      binding: '${{ inputs.binding }}',
      sha: '${{ inputs.sha }}'
    });
    expect(publish.jobs.publish.steps[0].with.ref).toBe('${{ inputs.sha }}');
    const toolchain = job.steps.find((step: any) => step.uses === 'dtolnay/rust-toolchain@stable');
    expect(toolchain.with.targets).toBe('${{ matrix.platform.target }}');
    const build = job.steps.find((step: any) => step.name === 'Build binding');
    expect(build.run).toBe(
      'bunx napi build --cwd bindings/${{ inputs.binding }}-native --platform --target ${{ matrix.platform.target }} --release -- --locked'
    );
  });

  test('dist uploads the generated loader from GNU Linux x64 after testing', () => {
    const steps = dist.jobs.bindings.steps;
    const loader = steps.find(
      (step: any) =>
        step.uses === 'actions/upload-artifact@v4' &&
        step.with.name === '${{ inputs.binding }}-native-loader'
    );
    expect(loader.if).toBe("matrix.platform.target == 'x86_64-unknown-linux-gnu'");
    expect(loader.with).toEqual({
      name: '${{ inputs.binding }}-native-loader',
      path: 'bindings/${{ inputs.binding }}-native/index.js',
      'if-no-files-found': 'error'
    });
    const binary = steps.find(
      (step: any) =>
        step.uses === 'actions/upload-artifact@v4' &&
        step.with.name === '${{ inputs.binding }}-native.${{ matrix.platform.target }}'
    );
    const testing = steps.find((step: any) => step.name === 'Test binding');
    expect(testing.run).toBe('bun test bindings/${{ inputs.binding }}-native/index.test.ts');
    expect(steps.indexOf(loader)).toBe(steps.indexOf(binary) + 1);
    expect(steps.indexOf(loader)).toBeGreaterThan(steps.indexOf(testing));
  });

  test('CI builds every path and checks declarations and all three runtimes', () => {
    const job = ci.jobs.native;
    expect(job.name).toBe('Native gate');
    expect(job['runs-on']).toBe('ubuntu-latest');
    expect(job['timeout-minutes']).toBe(45);
    expect(job.strategy).toBeUndefined();
    expect(job.steps[0].uses).toBe('actions/checkout@v7');
    const bun = job.steps.find((step: any) => step.uses === 'oven-sh/setup-bun@v2');
    expect(bun.with['bun-version']).toBe('1.4.3');
    const toolchain = job.steps.find((step: any) => step.uses === 'dtolnay/rust-toolchain@stable');
    expect(toolchain.with.targets).toBe('wasm32-unknown-unknown');
    const cache = job.steps.find((step: any) => step.uses === 'Swatinem/rust-cache@v2');
    expect(cache.with).toEqual({ workspaces: '.\nbindings\n', key: 'native' });
    expect(job.steps.some((step: any) => step.uses === './.github/actions/wasm-toolchain')).toBe(true);
    expect(job.steps.some((step: any) => step.with?.tool === 'wasm-pack@0.15.0')).toBe(true);
    const wasmCache = job.steps.find((step: any) => step.with?.path === 'target/wasm-pack');
    expect(wasmCache.with.key).toBe('wasm-pack-${{ steps.wasm.outputs.key }}');
    const runtime = job.steps.find((step: any) => step.uses === 'actions/setup-node@v7.0.0');
    expect(runtime.with['node-version']).toBe('24');
    const deno = job.steps.find((step: any) => step.uses === 'denoland/setup-deno@v2');
    expect(deno.with['deno-version']).toBe('v2.9.0');
    expect(job.steps.some((step: any) => step.run === 'bun install --frozen-lockfile')).toBe(true);
    const build = job.steps.find((step: any) => step.name === 'Build and test');
    expect(build.run).toContain('for package in $(node scripts/native-bindings.mjs --paths); do');
    expect(build.run).toContain('bunx napi build --cwd "$package" --platform -- --locked');
    expect(build.run).toContain('bun test "$package/index.test.ts"');
    const wasm = job.steps.find((step: any) => step.name === 'Build wasm for parity');
    expect(wasm.run).toContain('bun run --filter @betteroffice/xlsx build:wasm');
    expect(wasm.run).toContain('bun run --filter @betteroffice/pptx build:wasm');
    const parity = job.steps.find((step: any) => step.name === 'Check native and wasm parity');
    expect(parity.run).toBe('bun test bindings/xlsx-native/parity.test.ts bindings/pptx-native/parity.test.ts');
    expect(job.steps.indexOf(wasm)).toBeLessThan(job.steps.indexOf(build));
    expect(job.steps.indexOf(build)).toBeLessThan(job.steps.indexOf(parity));
    const root = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
    expect(root.scripts.test).not.toContain('bindings');
    const declarations = job.steps.find((step: any) => step.name === 'Check generated declarations');
    expect(declarations.run).toBe("git diff --exit-code -- 'bindings/*-native/index.d.ts'");
    const smoke = job.steps.find((step: any) => step.name === 'Check runtimes');
    expect(smoke.run).toContain('node scripts/native-bindings-smoke.mjs');
    expect(smoke.run).toContain('bun scripts/native-bindings-smoke.mjs');
    expect(smoke.run).toContain('for package in $(node scripts/native-bindings.mjs --paths); do');
    expect(smoke.run).toContain('$smoke_dir/node_modules/@betteroffice/$format');
    expect(smoke.run).toContain('cp scripts/native-bindings-smoke.mjs "$smoke_dir/smoke.mjs"');
    expect(smoke.run).toContain(
      'deno run --no-config --node-modules-dir=manual --allow-read --allow-env --allow-ffi "$smoke_dir/smoke.mjs" "$PWD"'
    );
    expect(job.steps.some((step: any) => step.uses?.startsWith('actions/upload-artifact@'))).toBe(
      false
    );
  });

  test('package assembly is path-filtered and builds every format', () => {
    expect(packages.name).toBe('Native packages');
    expect(Object.keys(packages.on).sort()).toEqual(['pull_request', 'workflow_dispatch']);
    expect(packages.on.pull_request.paths).toEqual([
      'bindings/*-native/**',
      'scripts/native-bindings.mjs',
      'scripts/native-bindings-smoke.mjs',
      'scripts/version-packages.mjs',
      '.github/workflows/native-packages.yml',
      '.github/workflows/native-dist.yml',
      '.github/workflows/publish-native.yml'
    ]);
    expect(packages.permissions).toEqual({ contents: 'read' });
    expect(packages.concurrency.group).toBe(
      '${{ github.workflow }}-${{ github.event.pull_request.number || github.ref }}'
    );
    expect(packages.concurrency['cancel-in-progress']).toBe(
      "${{ github.event_name == 'pull_request' }}"
    );
    const job = packages.jobs.dist;
    expect(job.strategy.matrix.binding).toEqual(NATIVE_FORMATS);
    expect(job.uses).toBe('./.github/workflows/native-dist.yml');
    expect(job.with).toEqual({
      binding: '${{ matrix.binding }}',
      sha: '${{ github.event.pull_request.head.sha || github.sha }}'
    });
  });

  test('package assembly validates and packs all six packages without uploading', () => {
    const job = packages.jobs.assemble;
    expect(job.needs).toBe('dist');
    expect(job.name).toBe('Assemble ${{ matrix.binding }}-native');
    expect(job.strategy.matrix.binding).toEqual(NATIVE_FORMATS);
    expect(job.steps[0].with.ref).toBe('${{ github.event.pull_request.head.sha || github.sha }}');
    const bun = job.steps.find((step: any) => step.uses === 'oven-sh/setup-bun@v2');
    expect(bun.with['bun-version']).toBe('1.4.3');
    const runtime = job.steps.find((step: any) => step.uses === 'actions/setup-node@v7.0.0');
    expect(runtime.with['node-version']).toBe('24');
    expect(job.steps.some((step: any) => step.run === 'bun install --frozen-lockfile')).toBe(true);
    const download = job.steps.find(
      (step: any) =>
        step.uses === 'actions/download-artifact@v4' &&
        step.with.pattern === '${{ matrix.binding }}-native.*'
    );
    expect(download.with).toEqual({
      pattern: '${{ matrix.binding }}-native.*',
      path: 'bindings/${{ matrix.binding }}-native/artifacts',
      'merge-multiple': true
    });
    const loader = job.steps.find(
      (step: any) =>
        step.uses === 'actions/download-artifact@v4' &&
        step.with.name === '${{ matrix.binding }}-native-loader'
    );
    expect(loader.with).toEqual({
      name: '${{ matrix.binding }}-native-loader',
      path: 'bindings/${{ matrix.binding }}-native'
    });
    const artifacts = job.steps.find((step: any) => step.name === 'Assemble packages');
    expect(artifacts.run).toBe('bunx napi artifacts --cwd bindings/${{ matrix.binding }}-native');
    const makePublic = job.steps.find((step: any) => step.name === 'Make the root publishable');
    expect(makePublic.run).toBe('npm pkg delete private');
    expect(makePublic['working-directory']).toBe('bindings/${{ matrix.binding }}-native');
    const validate = job.steps.find((step: any) => step.name === 'Validate packages');
    expect(validate.run).toBe(
      'bunx napi pre-publish --cwd bindings/${{ matrix.binding }}-native --tag-style npm --dry-run --no-gh-release'
    );
    expect(job.steps.indexOf(loader)).toBe(job.steps.indexOf(download) + 1);
    expect(job.steps.indexOf(loader)).toBeLessThan(job.steps.indexOf(artifacts));
    expect(job.steps.indexOf(artifacts)).toBeLessThan(job.steps.indexOf(makePublic));
    expect(job.steps.indexOf(makePublic)).toBeLessThan(job.steps.indexOf(validate));
    const pack = job.steps.find((step: any) => step.name === 'Check package contents');
    expect(pack['working-directory']).toBe('bindings/${{ matrix.binding }}-native');
    expect(pack.run).toStartWith(
      `npm pack --dry-run --json | jq -e '["index.js", "index.d.ts"] - [.[0].files[].path] == []'\n` +
        'npm pack --dry-run\n'
    );
    expect(pack.run).toContain('for package in npm/*; do');
    expect(pack.run).toContain('(cd "$package" && npm pack --dry-run)');
    expect(job.steps.indexOf(loader)).toBeLessThan(job.steps.indexOf(pack));
    expect(job.steps.indexOf(validate)).toBeLessThan(job.steps.indexOf(pack));
  });
});

describe('native package public names', () => {
  const read = (path: string) => readFileSync(new URL(`../${path}`, import.meta.url), 'utf8');
  const distWorkflowText = read('.github/workflows/native-dist.yml');

  test('root manifests carry the -native public name and binary', () => {
    for (const path of NATIVE_PACKAGES) {
      const manifest = JSON.parse(read(`${path}/package.json`));
      const format = path.replace(/^bindings\//, '').replace(/-native$/, '');
      expect(manifest.name).toBe(`@betteroffice/${format}-native`);
      expect(manifest.napi.packageName).toBe(`@betteroffice/${format}-native`);
      expect(manifest.napi.binaryName).toBe(`betteroffice-${format}-native`);
      expect(manifest.description).toContain('Node.js, Bun and Deno');
      expect(manifest.keywords).toContain('deno');
    }
  });

  test('platform packages match the root binary and version', () => {
    for (const path of NATIVE_PACKAGES) {
      const manifest = JSON.parse(read(`${path}/package.json`));
      const format = path.replace(/^bindings\//, '').replace(/-native$/, '');
      for (const identity of readdirSync(new URL(`../${path}/npm`, import.meta.url))) {
        const platform = JSON.parse(read(`${path}/npm/${identity}/package.json`));
        expect(platform.version).toBe(manifest.version);
        expect(platform.name.startsWith(`@betteroffice/${format}-native-`)).toBe(true);
        const triple = platform.name.replace(`@betteroffice/${format}-native-`, '');
        expect(identity).toBe(triple);
        expect(platform.main).toBe(`${manifest.napi.binaryName}.${triple}.node`);
        expect(platform.files).toEqual([platform.main]);
        expect(platform.description).toBe(manifest.description);
        expect(platform.keywords).toEqual(manifest.keywords);
      }
    }
  });

  test('ignores generated native loaders and keeps declarations', () => {
    expect(read('.gitignore').split('\n')).toContain('bindings/*-native/index.js');
    const result = spawnSync(
      'git',
      [
        'check-ignore',
        '--no-index',
        '--',
        ...NATIVE_PACKAGES.flatMap((path) => [`${path}/index.js`, `${path}/index.d.ts`])
      ],
      { cwd: fileURLToPath(new URL('../', import.meta.url)), encoding: 'utf8' }
    );
    expect(result.status).toBe(0);
    expect(result.stdout.trim().split('\n')).toEqual(
      NATIVE_PACKAGES.map((path) => `${path}/index.js`)
    );
  });

  test('the dist workflow collects the renamed binaries', () => {
    expect(distWorkflowText).toContain('betteroffice-${{ inputs.binding }}-native.*.node');
  });
});
