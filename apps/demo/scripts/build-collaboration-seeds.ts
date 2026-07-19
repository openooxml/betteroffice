import { mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { parseDocx } from "../../../packages/docx/src/docx/index.ts";
import {
  createYrsSession,
  documentToYrs,
} from "../../../packages/docx/src/yrs/index.ts";
import { preloadEditWasm } from "../../../packages/docx/src/wasm/edit.ts";
import { preloadOpcWasm } from "../../../packages/docx/src/wasm/opc.ts";
import { preloadParseWasm } from "../../../packages/docx/src/wasm/parse.ts";
import {
  initWasm as initXlsxWasm,
  openWorkbook,
} from "../../../packages/xlsx/src/index.ts";
import {
  initWasm as initPptxWasm,
  openPresentation,
} from "../../../packages/pptx/src/index.ts";

const demo = resolve(import.meta.dir, "..");
const root = resolve(demo, "../..");
const seeds = resolve(demo, "public/seeds");

function equalBytes(left: Uint8Array, right: Uint8Array): boolean {
  return (
    left.byteLength === right.byteLength &&
    left.every((value, index) => value === right[index])
  );
}

await mkdir(seeds, { recursive: true });

await Promise.all([
  preloadOpcWasm(
    await readFile(
      resolve(
        root,
        "packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm",
      ),
    ),
  ),
  preloadParseWasm(
    await readFile(
      resolve(
        root,
        "packages/docx/src/wasm/generated/parse/docx_parse_bg.wasm",
      ),
    ),
  ),
  preloadEditWasm(
    await readFile(
      resolve(
        root,
        "packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm",
      ),
    ),
  ),
]);

const docxBytes = new Uint8Array(
  await readFile(resolve(demo, "public/betteroffice-demo.docx")),
);
const document = await parseDocx(docxBytes);
const docxSession = await createYrsSession({ clientId: 1 });
documentToYrs(docxSession, document);
const docxSeed = docxSession.encodeState();
const docxStateVector = docxSession.encodeStateVector();
docxSession.destroy();
const docxVerification = await createYrsSession({ clientId: 2 });
docxVerification.loadState(docxSeed);
if (!equalBytes(docxVerification.encodeStateVector(), docxStateVector)) {
  throw new Error("DOCX collaboration seed did not round-trip");
}
docxVerification.destroy();
await writeFile(resolve(seeds, "docx.bin"), docxSeed);

await initXlsxWasm(
  await readFile(
    resolve(root, "packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm"),
  ),
);
const xlsxBytes = new Uint8Array(
  await readFile(resolve(demo, "public/showcase.xlsx")),
);
const workbook = openWorkbook(xlsxBytes, {
  collaborative: true,
  clientId: 1,
});
const xlsxSeed = workbook.encodeStateAsUpdate();
const xlsxStateVector = workbook.encodeStateVector();
workbook.dispose();
const xlsxVerification = openWorkbook(xlsxBytes, {
  collaborative: true,
  clientId: 2,
});
xlsxVerification.applyUpdate(xlsxSeed);
if (!equalBytes(xlsxVerification.encodeStateVector(), xlsxStateVector)) {
  throw new Error("XLSX collaboration seed did not round-trip");
}
xlsxVerification.dispose();
await writeFile(resolve(seeds, "xlsx.bin"), xlsxSeed);

await initPptxWasm(
  await readFile(
    resolve(root, "packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm"),
  ),
);
const pptxBytes = new Uint8Array(
  await readFile(resolve(demo, "public/betteroffice-demo.pptx")),
);
const presentation = openPresentation(pptxBytes, { clientId: 1 });
const pptxSeed = presentation.encodeStateAsUpdate();
const pptxStateVector = presentation.encodeStateVector();
presentation.dispose();
const pptxVerification = openPresentation(pptxBytes, {
  clientId: 2,
  initialUpdate: pptxSeed,
});
if (!equalBytes(pptxVerification.encodeStateVector(), pptxStateVector)) {
  throw new Error("PPTX collaboration seed did not round-trip");
}
if (!equalBytes(pptxVerification.encodeStateAsUpdate(), pptxSeed)) {
  throw new Error("PPTX collaboration seed changed during round-trip");
}
pptxVerification.dispose();
await writeFile(resolve(seeds, "pptx.bin"), pptxSeed);

console.log(
  `Wrote DOCX (${docxSeed.byteLength} bytes), XLSX (${xlsxSeed.byteLength} bytes), and PPTX (${pptxSeed.byteLength} bytes) collaboration seeds`,
);
