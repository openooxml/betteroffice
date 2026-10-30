import { plugin } from 'bun';
import { NATIVE_ENGINE_FILTER, nativeEngineFor } from './native-engine';

// Links the -native engine modules in place of the wasm ones, as the -native package builds do.
plugin({
  name: 'betteroffice-native-engine',
  setup(build) {
    build.onResolve({ filter: NATIVE_ENGINE_FILTER }, (args) => {
      if (!args.path.startsWith('.') && !args.path.startsWith('/')) return undefined;
      const path = nativeEngineFor(args.path, args.importer);
      return path === undefined ? undefined : { path };
    });
  },
});

(globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')] = 'native';
