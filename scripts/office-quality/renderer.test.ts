import { expect, test } from 'bun:test';
import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { createServer } from 'vite';
import { qualityRendererPlugin } from './renderer.mjs';

for (const format of ['pptx', 'xlsx', 'vsdx']) {
  for (const channel of ['commit', 'published']) {
    test(`${format} ${channel} harness resolves only its selected renderer`, async () => {
      const directory = await mkdtemp(resolve(tmpdir(), 'office-quality-renderer-'));
      const renderer = resolve(directory, channel, 'index.js');
      await mkdir(resolve(directory, channel));
      await writeFile(renderer, 'export const initWasm = async () => {};');
      const requested: string[] = [];
      const server = await createServer({
        configFile: false,
        root: import.meta.dir,
        cacheDir: resolve(directory, 'cache'),
        logLevel: 'silent',
        server: {
          middlewareMode: true,
          hmr: false,
          watch: null,
          preTransformRequests: false,
        },
        optimizeDeps: { noDiscovery: true, include: [] },
        plugins: [
          qualityRendererPlugin(format),
          {
            name: 'only-selected-renderer-is-built',
            enforce: 'pre',
            resolveId(source) {
              if (!/^@betteroffice\/(?:pptx|xlsx|vsdx)$/.test(source)) return;
              requested.push(source);
              if (source !== `@betteroffice/${format}`)
                throw new Error(`Unbuilt renderer: ${source}`);
              return renderer;
            },
          },
        ],
        resolve: {
          alias:
            channel === 'published'
              ? [{ find: `@betteroffice/${format}`, replacement: renderer }]
              : [],
        },
      });
      try {
        const harness = await server.transformRequest('/harness.ts');
        expect(harness?.code).toContain('virtual:office-quality-renderer');
        const selected = await server.transformRequest(
          'virtual:office-quality-renderer'
        );
        expect(selected?.code).toContain(renderer);
        expect(requested.every((source) => source === `@betteroffice/${format}`)).toBe(true);
        if (channel === 'commit') expect(requested).toContain(`@betteroffice/${format}`);
      } finally {
        await server.close();
        await rm(directory, { recursive: true, force: true });
      }
    });
  }
}
