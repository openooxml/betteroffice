import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts', 'src/mcp.ts', 'src/render.ts', 'src/cli.ts', 'src/engines.node.ts', 'src/engines.browser.ts'],
  format: ['esm'],
  dts: true,
  splitting: true,
  clean: true,
  external: [/^@betteroffice\//, '#agents-engines', '@napi-rs/canvas'],
});
