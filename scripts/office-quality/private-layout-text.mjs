import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const [root, fontsRoot, input, output] = process.argv.slice(2);
if (!root || !fontsRoot || !input || !output || process.argv.length !== 6)
  throw new Error('usage: private-layout-text.mjs docx-package-root fonts-dist-dir source.docx out.json');
const load = (path) => import(pathToFileURL(resolve(path)).href);
const [yrs, editor, layout, fonts] = await Promise.all([
  load(resolve(root, 'dist/yrs/index.mjs')),
  load(resolve(root, 'dist/editor/index.mjs')),
  load(resolve(root, 'dist/layout/index.mjs')),
  load(resolve(fontsRoot, 'index.js')),
]);
const bytes = new Uint8Array(await readFile(input));
layout.configureDefaultFonts({ fonts });
const session = await yrs.createYrsSession({ clientId: 7 });

function blockLines(blocks) {
  return blocks.flatMap((block) => {
    if (block.kind === 'paragraph')
      return [(block.runs ?? []).map((run) => run.text ?? '').join('')];
    if (block.kind === 'table')
      return block.rows.flatMap((row) => row.cells.flatMap((cell) => blockLines(cell.blocks)));
    return [];
  });
}

try {
  const host = await session.openDocx(bytes, true);
  const document = host.document;
  const source = layout.createRustMeasureSource({ engine: session });
  source.setCompat(document.package.settings?.compatibilityFlags);
  const renderEnv = {
    themeColors: {},
    defaultTabStopTwips: document.package.settings?.defaultTabStop ?? null,
    numericIds: {},
    showHiddenText: false,
  };
  const request = editor.buildResidentRegionLayoutRequest(document, 24, renderEnv);
  const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request)));
  await source.prepareFontRequirements(requirements);
  const measurement = source.measurementConfigForRequirements(requirements);
  const comp = editor.computeLayout({ document, pageGap: 24, session, renderEnv, measurement });
  const blocks = new Map(session.yrsBlocksForStory('body', renderEnv).map((block) => [String(block.id), block]));
  const pages = comp.layout.pages.map((page) => page.fragments.flatMap((fragment) => {
    if (fragment.kind === 'paragraph') {
      if (!Array.isArray(fragment.resolvedLines))
        throw new Error(`Missing resolved lines for body block ${fragment.blockId}`);
      return fragment.resolvedLines.map((line) => line.segments.map((segment) => segment.text ?? '').join(''));
    }
    if (fragment.kind === 'table') {
      const table = blocks.get(String(fragment.blockId));
      if (table?.kind !== 'table') throw new Error(`Missing body table ${fragment.blockId}`);
      return table.rows.slice(fragment.rowStart, fragment.rowEnd)
        .flatMap((row) => row.cells.flatMap((cell) => blockLines(cell.blocks)));
    }
    return [];
  }));
  await writeFile(output, JSON.stringify(pages) + '\n');
  console.log(JSON.stringify({ pages: pages.length }));
} finally {
  session.destroy();
}
