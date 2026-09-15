import { describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import {
  NODE_BINDINGS,
  NODE_BINDING_NAMES,
  bindingVersion,
  pendingPublishNames,
  platformPackageVersions
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
    const versions = new Set(NODE_BINDINGS.map(bindingVersion));
    expect(versions.size).toBe(1);
    expect(new Set(platformPackageVersions().map((entry) => entry.version))).toEqual(versions);
  });

  test('detects only unpublished root versions', async () => {
    const pending = await pendingPublishNames({
      fetchImpl: async (url: string) => {
        const name = decodeURIComponent(new URL(url).pathname.slice(1));
        return name.includes('pptx')
          ? new Response('{"error":"missing"}', { status: 404 })
          : Response.json({ versions: { '0.0.1': {} } });
      }
    });
    expect(pending).toEqual(['pptx']);
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
  });

  test('publisher uses OIDC and builds all declared platforms at the requested commit', () => {
    expect(publish.jobs.publish.permissions['id-token']).toBe('write');
    expect(publish.jobs.publish.environment).toBe('npm-${{ inputs.binding }}');
    expect(publish.jobs.publish.steps.some((step: any) => step.run?.includes('NPM_TOKEN'))).toBe(
      true
    );
    expect(dist.jobs.bindings.strategy.matrix.platform).toHaveLength(5);
    expect(dist.jobs.bindings.steps[0].with.ref).toBe('${{ inputs.sha }}');
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
