import { plugin } from 'bun';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { NATIVE_ENGINE_FILTER, nativeEngineSource } from './native-engine';

plugin({
  name: 'betteroffice-native-engine',
  setup(build) {
    build.onLoad({ filter: NATIVE_ENGINE_FILTER }, ({ path }) => {
      return { contents: nativeEngineSource(path) ?? readFileSync(path, 'utf8'), loader: 'ts' };
    });
  },
});

const preload = fileURLToPath(import.meta.url);
(globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')] = 'native';
(globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testPreload')] = preload;

const BaseWorker = globalThis.Worker;
globalThis.Worker = class extends BaseWorker {
  constructor(url: string | URL, options: WorkerOptions & { preload?: string | string[] } = {}) {
    const preloads = options.preload === undefined ? [] : [options.preload].flat();
    super(url, { ...options, preload: [...preloads, preload] } as WorkerOptions);
  }
} as typeof Worker;
const option = `--preload ${preload.replace(/\s/g, '\\$&')}`;
if (!process.env.BUN_OPTIONS?.includes(option)) {
  process.env.BUN_OPTIONS = [process.env.BUN_OPTIONS, option].filter(Boolean).join(' ');
}
