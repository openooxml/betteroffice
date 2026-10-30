import { plugin } from 'bun';
import { fileURLToPath } from 'node:url';
import { NATIVE_ENGINE_FILTER, nativeEngineSource } from './native-engine';

plugin({
  name: 'betteroffice-native-engine',
  setup(build) {
    build.onLoad({ filter: NATIVE_ENGINE_FILTER }, ({ path }) => {
      const contents = nativeEngineSource(path);
      return contents === undefined ? undefined : { contents, loader: 'ts' };
    });
  },
});

(globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')] = 'native';

const preload = fileURLToPath(import.meta.url);
const BaseWorker = globalThis.Worker;
globalThis.Worker = class extends BaseWorker {
  constructor(url: string | URL, options: WorkerOptions & { preload?: string | string[] } = {}) {
    const preloads = options.preload === undefined ? [] : [options.preload].flat();
    super(url, { ...options, preload: [...preloads, preload] } as WorkerOptions);
  }
} as typeof Worker;
process.env.BUN_OPTIONS = [process.env.BUN_OPTIONS, `--preload ${preload}`].filter(Boolean).join(' ');
