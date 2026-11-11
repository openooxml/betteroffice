import { plugin } from 'bun';

/**
 * Preload for the display-list parity child: PNG export hands each page to the
 * engine's `renderPng`, which this replaces with one that records the page and
 * draws nothing. List it before the native engine preload so it wins that module.
 */
(globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.displayListParity')] = [];

plugin({
  name: 'betteroffice-display-list-parity',
  setup(build) {
    build.onLoad({ filter: /[\\/]packages[\\/]docx[\\/]src[\\/]wasm[\\/]enginePng\.ts$/ }, () => ({
      loader: 'ts',
      contents: `
        const key = Symbol.for('betteroffice.displayListParity');
        export const engineKind = globalThis[Symbol.for('betteroffice.testEngine')] === 'native' ? 'native' : 'wasm';
        export async function renderPng(engine) {
          globalThis[key].push(JSON.stringify({ fontChains: engine.fontChains, page: engine.page }));
          return new Uint8Array();
        }
        export default renderPng;
      `,
    }));
  },
});
