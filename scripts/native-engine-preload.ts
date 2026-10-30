import { plugin } from 'bun';
import { NATIVE_ENGINE_FILTER, nativeEngineSource } from './native-engine';

// Links the -native engine modules in place of the wasm ones, as the -native package builds do.
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
