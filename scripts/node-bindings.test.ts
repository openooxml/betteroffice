import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  NODE_BINDINGS,
  NODE_BINDING_NAMES,
  bindingVersion,
  pendingPublishNames,
  platformPackageVersions,
  synchronizeNodeLoader,
  validateNodeVersions
} from './node-bindings.mjs';

const releaseWorkflow = fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url));
const publishWorkflow = fileURLToPath(
  new URL('../.github/workflows/publish-node-binding.yml', import.meta.url)
);
const distWorkflow = fileURLToPath(new URL('../.github/workflows/node-dist.yml', import.meta.url));

describe('Node binding registry', () => {
  test('registers every native format and platform package', () => {
    expect(NODE_BINDING_NAMES).toEqual(['docx', 'pptx', 'xlsx']);
    expect(NODE_BINDINGS).toEqual(NODE_BINDING_NAMES.map((name) => `bindings/node-${name}`));
    expect(platformPackageVersions()).toHaveLength(NODE_BINDINGS.length * 5);
    expect(() => validateNodeVersions()).not.toThrow();
  });

  test('detects only unpublished root versions', async () => {
    const pending = await pendingPublishNames({
      fetchImpl: async (url: string) => {
        const name = decodeURIComponent(new URL(url).pathname.slice(1));
        const format = name.replace('@betteroffice/', '').replace('-native', '');
        return name.includes('pptx')
          ? new Response('{"error":"missing"}', { status: 404 })
          : Response.json({ versions: { [bindingVersion(`bindings/node-${format}`)]: {} } });
      }
    });
    expect(pending).toEqual(['pptx']);
  });

  test('registry failures stop release detection', async () => {
    await expect(pendingPublishNames({
      fetchImpl: async () => new Response('unavailable', { status: 503 })
    })).rejects.toThrow('npm answered 503');
  });
});

describe('native version synchronization', () => {
  test('each native package joins its core fixed release group', () => {
    const config = JSON.parse(readFileSync(new URL('../.changeset/config.json', import.meta.url), 'utf8'));
    for (const format of NODE_BINDING_NAMES) {
      const group = config.fixed.find((names: string[]) => names.includes(`@betteroffice/${format}`));
      expect(group).toContain(`@betteroffice/${format}-native`);
      expect(bindingVersion(`bindings/node-${format}`)).toBe(bindingVersion(`packages/${format}`));
    }
  });

  test('updates generated checks and error messages for every platform', () => {
    for (const binding of NODE_BINDINGS) {
      const source = readFileSync(new URL(`../${binding}/index.js`, import.meta.url), 'utf8');
      const before = bindingVersion(binding);
      const updated = synchronizeNodeLoader(source, before, '999.999.999');
      expect(updated).not.toContain(`bindingPackageVersion !== '${before}'`);
      expect(updated).not.toContain(`version mismatch, expected ${before} but got`);
      expect(updated).toContain("bindingPackageVersion !== '999.999.999'");
      expect(synchronizeNodeLoader(updated, '999.999.999', before)).toBe(source);
      expect(() => synchronizeNodeLoader(updated, before, '999.999.999')).toThrow('not synchronized');
    }
  });
});

describe('Node binding release wiring', () => {
  const release = Bun.YAML.parse(readFileSync(releaseWorkflow, 'utf8')) as any;
  const publish = Bun.YAML.parse(readFileSync(publishWorkflow, 'utf8')) as any;
  const dist = Bun.YAML.parse(readFileSync(distWorkflow, 'utf8')) as any;

  test('release dispatches and waits for the dedicated publisher', () => {
    const step = release.jobs.release.steps.find(
      (value: any) => value.name === 'Publish Node native bindings'
    );
    expect(step.run).toContain('scripts/node-bindings.mjs --pending');
    expect(step.run).toContain('publish-node-binding.yml/dispatches');
    expect(step.run).toContain('gh run watch');
    expect(step.run).toContain('pending=$(node scripts/node-bindings.mjs --pending)');
    const steps = release.jobs.release.steps.map((value: any) => value.name);
    expect(steps.indexOf('Authenticate to crates.io')).toBeGreaterThan(steps.indexOf(step.name));
  });

  test('publisher uses OIDC and builds all declared platforms at the requested commit', () => {
    expect(publish.jobs.publish.permissions['id-token']).toBe('write');
    expect(publish.jobs.publish.environment).toBe('npm-${{ inputs.binding }}');
    expect(publish.jobs.publish.steps.some((step: any) => step.run?.includes('NPM_TOKEN'))).toBe(
      true
    );
    expect(dist.jobs.bindings.strategy.matrix.platform).toHaveLength(5);
    expect(dist.jobs.bindings.steps[0].with.ref).toBe('${{ inputs.sha }}');
    const build = dist.jobs.bindings.steps.find((step: any) => step.name === 'Build binding');
    expect(build.run).toContain('--platform');
  });

  test('a failed registry lookup fails the dispatch step', () => {
    const directory = mkdtempSync(join(tmpdir(), 'native-dispatch-'));
    try {
      writeFileSync(join(directory, 'node'), '#!/bin/sh\nexit 42\n', { mode: 0o755 });
      const step = release.jobs.release.steps.find((value: any) => value.name === 'Publish Node native bindings');
      const result = spawnSync('bash', ['-e', '-o', 'pipefail', '-c', step.run], {
        encoding: 'utf8',
        env: { ...process.env, PATH: `${directory}:${process.env.PATH}` }
      });
      expect(result.status).toBe(42);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test('CI tests every published platform and assembles the complete artifact set', () => {
    const ci = Bun.YAML.parse(readFileSync(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8')) as any;
    expect(ci.jobs['node-native'].strategy.matrix.platform).toEqual(dist.jobs.bindings.strategy.matrix.platform);
    expect(ci.jobs['node-packages'].needs).toBe('node-native');
    expect(ci.jobs['node-packages'].strategy.matrix.binding).toEqual(NODE_BINDING_NAMES);
    expect(ci.jobs['node-packages'].steps.at(-1).run).toContain('napi pre-publish');
    expect(ci.jobs['node-packages'].steps.at(-1).run).toContain('--dry-run');
  });
});

describe('Node binding public names', () => {
  const read = (path: string) => readFileSync(new URL(`../${path}`, import.meta.url), 'utf8');
  const distWorkflowText = read('.github/workflows/node-dist.yml');

  test('root manifests carry the -native public name and binary', () => {
    for (const path of NODE_BINDINGS) {
      const manifest = JSON.parse(read(`${path}/package.json`));
      const format = path.replace('bindings/node-', '');
      expect(manifest.name).toBe(`@betteroffice/${format}-native`);
      expect(manifest.napi.packageName).toBe(`@betteroffice/${format}-native`);
      expect(manifest.napi.binaryName).toBe(`betteroffice-${format}-native`);
    }
  });

  test('platform packages match the root binary and version', () => {
    for (const path of NODE_BINDINGS) {
      const manifest = JSON.parse(read(`${path}/package.json`));
      const format = path.replace('bindings/node-', '');
      for (const identity of readdirSync(new URL(`../${path}/npm`, import.meta.url))) {
        const platform = JSON.parse(read(`${path}/npm/${identity}/package.json`));
        expect(platform.version).toBe(manifest.version);
        expect(platform.name.startsWith(`@betteroffice/${format}-native-`)).toBe(true);
        const triple = platform.name.replace(`@betteroffice/${format}-native-`, '');
        expect(identity).toBe(triple);
        expect(platform.main).toBe(`${manifest.napi.binaryName}.${triple}.node`);
        expect(platform.files).toEqual([platform.main]);
      }
    }
  });

  test('generated loaders require only the -native platform packages', () => {
    for (const path of NODE_BINDINGS) {
      const manifest = JSON.parse(read(`${path}/package.json`));
      const format = path.replace('bindings/node-', '');
      const loader = read(`${path}/index.js`);
      const required = [...loader.matchAll(/require\('(@betteroffice\/[^']+)'\)/g)].map((m) => m[1]);
      expect(required.length).toBeGreaterThan(0);
      for (const name of required) {
        expect(name.startsWith(`@betteroffice/${format}-native`)).toBe(true);
      }
      expect(loader).toContain(`./${manifest.napi.binaryName}.`);
      expect(loader).not.toContain(`@betteroffice/${format}-node`);
      expect(loader).not.toContain(`betteroffice-${format}-node`);
    }
  });

  test('the dist workflow collects the renamed binaries', () => {
    expect(distWorkflowText).toContain('betteroffice-${{ inputs.binding }}-native.*.node');
    expect(distWorkflowText).not.toContain('-node.*.node');
  });
});
