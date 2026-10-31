import { writeFile } from 'node:fs/promises';
import { defineConfig } from 'tsup';
import { nativeEnginePlugin } from '../../../scripts/native-engine';

export default defineConfig([
  {
    entry: {
      core: '../../../packages/docx/src/core.ts',
      worker: '../../../packages/docx/src/worker.ts',
      'docx/index': '../../../packages/docx/src/docx/index.ts',
      'docx/wrapTypes': '../../../packages/docx/src/docx/wrapTypes.ts',
      'styles/index': '../../../packages/docx/src/styles/index.ts',
      'utils/index': '../../../packages/docx/src/utils/index.ts',
      'utils/cardStyles': '../../../packages/docx/src/utils/cardStyles.ts',
      'types/document': '../../../packages/docx/src/types/document.ts',
      'types/content': '../../../packages/docx/src/types/content.ts',
      'layout/pagination/index': '../../../packages/docx/src/layout/pagination/index.ts',
      'layout/render/index': '../../../packages/docx/src/layout/render/index.ts',
      'layout/index': '../../../packages/docx/src/layout/index.ts',
      'plugin-api/index': '../../../packages/docx/src/plugin-api/index.ts',
      'plugin-api/RenderedDomContext': '../../../packages/docx/src/plugin-api/RenderedDomContext.ts',
      'plugin-api/resolveItemPositions': '../../../packages/docx/src/plugin-api/resolveItemPositions.ts',
      'utils/comments': '../../../packages/docx/src/utils/comments.ts',
      'utils/findReplace': '../../../packages/docx/src/utils/findReplace.ts',
      'utils/findVerticalScrollParent': '../../../packages/docx/src/utils/findVerticalScrollParent.ts',
      'utils/fontOptions': '../../../packages/docx/src/utils/fontOptions.ts',
      'utils/stylePreview': '../../../packages/docx/src/utils/stylePreview.ts',
      'utils/listState': '../../../packages/docx/src/utils/listState.ts',
      'utils/reportIssue': '../../../packages/docx/src/utils/reportIssue.ts',
      'utils/sidebarConstants': '../../../packages/docx/src/utils/sidebarConstants.ts',
      'utils/units': '../../../packages/docx/src/utils/units.ts',
      'editor/index': '../../../packages/docx/src/editor/index.ts',
      'utils/autoScroll': '../../../packages/docx/src/utils/autoScroll.ts',
      'yrs/index': '../../../packages/docx/src/yrs/index.ts',
      'collaboration/index': '../../../packages/docx/src/collaboration/index.ts',
      'yrs/docxWorker': '../../../packages/docx/src/yrs/residentEngineWorker.ts',
      'wasm-opc': '../../../packages/docx/src/wasm/opc.ts',
      'wasm-layout': '../../../packages/docx/src/wasm/layout.ts',
      'wasm-edit': '../../../packages/docx/src/wasm/edit.ts',
      'wasm-parse': '../../../packages/docx/src/wasm/parse.ts',
      'wasm-preflight': '../../../packages/docx/src/wasm/preflight.ts',
      'wasm-query': '../../../packages/docx/src/wasm/query.ts',
    },
    outDir: '../dist',
    tsconfig: '../../../packages/docx/tsconfig.json',
    format: ['esm'],
    shims: true,
    dts: true,
    splitting: true,
    sourcemap: false,
    clean: true,
    treeshake: true,
    minify: true,
    injectStyle: false,
    external: ['@betteroffice/fonts'],
    esbuildPlugins: [nativeEnginePlugin('docx')],
    esbuildOptions(options) {
      options.chunkNames = 'chunk-[hash]';
    },
    onSuccess: async () => {
      await writeFile('../dist/browser.js', 'throw new Error("use @betteroffice/docx in browsers");\n');
      await writeFile('../dist/docxWorker.mjs', "import './yrs/docxWorker.mjs';\n");
    },
  },
]);
