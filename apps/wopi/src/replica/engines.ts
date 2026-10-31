import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  createYrsSession,
  saveYrsDocx,
  type YrsSession,
} from "../../../../packages/docx/src/yrs";
import { preloadEditWasm } from "../../../../packages/docx/src/wasm/edit";
import { preloadOpcWasm } from "../../../../packages/docx/src/wasm/opc";
import {
  initWasm as initXlsxWasm,
  openWorkbook,
  type WorkbookHandle,
} from "../../../../packages/xlsx/src/wasm/loader";
import {
  initWasm as initPptxWasm,
  openPresentation,
  type PresentationHandle,
} from "../../../../packages/pptx/src/wasm/loader";

export type { PresentationHandle, WorkbookHandle, YrsSession };

const ROOT = resolve(import.meta.dir, "../../../..");

function wasmBytes(relativePath: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(ROOT, relativePath)));
}

let enginesReady: Promise<void> | undefined;

/** Preloads the wasm every format needs. Idempotent. */
export function initEngines(): Promise<void> {
  if (!enginesReady) {
    enginesReady = (async () => {
      await preloadEditWasm(
        wasmBytes("packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm"),
      );
      await preloadOpcWasm(
        wasmBytes("packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm"),
      );
      await initXlsxWasm(
        wasmBytes("packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm"),
      );
      await initPptxWasm(
        wasmBytes("packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm"),
      );
    })();
    enginesReady.catch(() => {
      enginesReady = undefined;
    });
  }
  return enginesReady;
}

/** A docx replica session opened on the source bytes. */
export async function openDocxReplica(
  bytes: Uint8Array,
  clientId: number,
): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  try {
    session.openDocx(bytes.slice(), true);
  } catch (error) {
    session.destroy();
    throw error;
  }
  return session;
}

/** The headless docx save path. */
export async function saveDocxReplica(session: YrsSession): Promise<Uint8Array> {
  return (await saveYrsDocx(session)).bytes;
}

/** An xlsx replica opened collaboratively on the source bytes. */
export function openXlsxReplica(bytes: Uint8Array, clientId: number): WorkbookHandle {
  return openWorkbook(bytes.slice(), { collaborative: true, clientId });
}

/** A pptx replica opened collaboratively on the source bytes. */
export function openPptxReplica(
  bytes: Uint8Array,
  clientId: number,
): PresentationHandle {
  return openPresentation(bytes.slice(), { clientId });
}
