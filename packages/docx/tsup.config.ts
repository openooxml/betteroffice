import { defineConfig } from 'tsup';

export default defineConfig([
  {
    entry: {
      core: 'src/core.ts',
      'docx/index': 'src/docx/index.ts',
      'docx/wrapTypes': 'src/docx/wrapTypes.ts',
      'styles/index': 'src/styles/index.ts',
      'utils/index': 'src/utils/index.ts',
      'utils/cardStyles': 'src/utils/cardStyles.ts',
      'types/document': 'src/types/document.ts',
      'types/content': 'src/types/content.ts',
      'layout/pagination/index': 'src/layout/pagination/index.ts',
      'layout/render/index': 'src/layout/render/index.ts',
      'layout/index': 'src/layout/index.ts',
      'plugin-api/index': 'src/plugin-api/index.ts',
      'plugin-api/RenderedDomContext': 'src/plugin-api/RenderedDomContext.ts',
      'plugin-api/resolveItemPositions': 'src/plugin-api/resolveItemPositions.ts',
      'utils/comments': 'src/utils/comments.ts',
      'utils/findReplace': 'src/utils/findReplace.ts',
      'utils/findVerticalScrollParent': 'src/utils/findVerticalScrollParent.ts',
      'utils/fontOptions': 'src/utils/fontOptions.ts',
      'utils/stylePreview': 'src/utils/stylePreview.ts',
      'utils/listState': 'src/utils/listState.ts',
      'utils/reportIssue': 'src/utils/reportIssue.ts',
      'utils/sidebarConstants': 'src/utils/sidebarConstants.ts',
      'utils/units': 'src/utils/units.ts',
      'editor/index': 'src/editor/index.ts',
      'utils/autoScroll': 'src/utils/autoScroll.ts',
      // The yrs editing-core facade (the only JS entry to crates/docx-edit).
      // Its embedded wasm stays out of every other entry: the facade reaches
      // ./wasm only via dynamic import, and hosts load the facade lazily.
      'yrs/index': 'src/yrs/index.ts',
      // Dedicated browser worker loaded relative to the yrs facade bundle.
      'yrs/residentEngineWorker': 'src/yrs/residentEngineWorker.ts',
      // The four wasm loaders are ROOT-NAMED entries on purpose: their
      // `new URL('./generated/…', import.meta.url)` literals must resolve
      // against a root-level chunk next to dist/generated/ (copy-assets puts
      // the gitignored binaries there). Keep them at the dist root.
      'wasm-opc': 'src/wasm/opc.ts',
      'wasm-layout': 'src/wasm/layout.ts',
      'wasm-edit': 'src/wasm/edit.ts',
      'wasm-parse': 'src/wasm/parse.ts',
    },
    format: ['esm'],
    // cjs builds need the import.meta.url shim for the wasm asset URLs.
    shims: true,
    dts: true,
    splitting: true,
    sourcemap: false,
    clean: true,
    treeshake: true,
    minify: true,
    injectStyle: false,
  },
]);
