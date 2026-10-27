import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  createYrsSession,
  saveYrsDocx,
  type YrsSession,
} from "../../../../packages/docx/src/yrs";
import { unzipContainer } from "../../../../packages/docx/src/wasm/opc";
import {
  openWorkbook,
  type WorkbookHandle,
} from "../../../../packages/xlsx/src/wasm/loader";
import {
  openPresentation,
  type PresentationHandle,
} from "../../../../packages/pptx/src/wasm/loader";
import type { Format, JoinPayload } from "../../src/replica/replica";

const ROOT = resolve(import.meta.dir, "../../../..");

const FIXTURE_PATHS: Record<Format, string> = {
  docx: resolve(ROOT, "apps/demo/public/betteroffice-demo.docx"),
  xlsx: resolve(ROOT, "packages/xlsx/test-fixtures/sample.xlsx"),
  pptx: resolve(ROOT, "apps/demo/public/betteroffice-demo.pptx"),
};

/** Source bytes for a format fixture. */
export function sourceBytes(format: Format): Uint8Array {
  return new Uint8Array(readFileSync(FIXTURE_PATHS[format]));
}

/** Whether two byte arrays are byte-equal. */
export function byteEqual(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((byte, i) => byte === b[i]);
}

/** Unzipped parts of a package, keyed by part name. */
export function unzipParts(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipContainer(bytes.slice());
}

/** Part names whose bytes differ between two saves, or that exist in only one. */
export function differingParts(a: Uint8Array, b: Uint8Array): string[] {
  if (byteEqual(a, b)) return [];
  const pa = unzipParts(a);
  const pb = unzipParts(b);
  const names = [...new Set([...Object.keys(pa), ...Object.keys(pb)])].sort();
  return names.filter((name) => {
    const x = pa[name];
    const y = pb[name];
    if (!x || !y) return true;
    return !byteEqual(x, y);
  });
}

export interface TestClient {
  readonly format: Format;
  readonly clientId: number;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): void;
  edit(marker: string, slot?: number): void;
  remove(marker: string): void;
  contains(marker: string): boolean;
  save(): Promise<Uint8Array>;
  destroy(): void;
}

function docxText(session: YrsSession): string {
  return session
    .storyIds()
    .sort()
    .flatMap((story) => session.paragraphs(story).map((p) => p.text))
    .join("\n");
}

function wrapDocx(session: YrsSession, extra?: YrsSession): TestClient {
  return {
    format: "docx",
    clientId: session.clientId,
    encodeStateVector: () => session.encodeStateVector(),
    encodeStateAsUpdate: (remote) => session.encodeStateAsUpdate(remote),
    applyUpdate: (update) => {
      session.applyUpdate(update.slice());
    },
    edit: (marker, slot = 0) => {
      const paragraphs = session.paragraphs("body");
      const para =
        slot === 0
          ? paragraphs[0]!
          : slot === 1
            ? paragraphs.at(-1)!
            : paragraphs[slot - 1]!;
      session.insertText({ story: "body", paraId: para.paraId, offset: 0 }, marker);
    },
    remove: (marker) => {
      for (const para of session.paragraphs("body")) {
        const at = para.text.indexOf(marker);
        if (at === -1) continue;
        session.deleteRange({
          story: "body",
          start: { paraId: para.paraId, offset: at },
          end: { paraId: para.paraId, offset: at + marker.length },
        });
        return;
      }
      throw new Error("marker not found in docx client");
    },
    contains: (marker) => docxText(session).includes(marker),
    save: async () => (await saveYrsDocx(session)).bytes,
    destroy: () => {
      session.destroy();
      extra?.destroy();
    },
  };
}

function wrapXlsx(handle: WorkbookHandle): TestClient {
  let lastCell = { row: 25, col: 0 };
  return {
    format: "xlsx",
    clientId: handle.clientId,
    encodeStateVector: () => handle.encodeStateVector(),
    encodeStateAsUpdate: (remote) => handle.encodeStateAsUpdate(remote),
    applyUpdate: (update) => {
      handle.applyUpdate(update.slice());
    },
    edit: (marker, slot = 0) => {
      lastCell = { row: 25 + slot, col: 0 };
      handle.editCell(0, lastCell.row, lastCell.col, marker);
    },
    remove: (marker) => {
      for (let row = 20; row <= 32; row += 1) {
        for (let col = 0; col <= 2; col += 1) {
          if (handle.cell(0, row, col).input.includes(marker)) {
            handle.editCell(0, row, col, "");
            return;
          }
        }
      }
      throw new Error("marker not found in xlsx client");
    },
    contains: (marker) => {
      for (let row = 20; row <= 32; row += 1) {
        for (let col = 0; col <= 2; col += 1) {
          if (handle.cell(0, row, col).input.includes(marker)) return true;
        }
      }
      return false;
    },
    save: async () => handle.save(),
    destroy: () => handle.dispose(),
  };
}

function wrapPptx(handle: PresentationHandle): TestClient {
  const firstStory = () =>
    handle.snapshot().slides[0]!.shapes.find((s) => s.textStories.length > 0)!
      .textStories[0]!;
  return {
    format: "pptx",
    clientId: handle.clientId,
    encodeStateVector: () => handle.encodeStateVector(),
    encodeStateAsUpdate: (remote) => handle.encodeStateAsUpdate(remote),
    applyUpdate: (update) => {
      handle.applyUpdate(update.slice());
    },
    edit: (marker, slot = 0) => {
      const stories = handle
        .snapshot()
        .slides.flatMap((slide) =>
          slide.shapes.flatMap((shape) => shape.textStories),
        );
      const story = stories[slot % stories.length] ?? firstStory();
      handle.insertText(story.id, 0, marker);
    },
    remove: (marker) => {
      for (const slide of handle.snapshot().slides) {
        for (const shape of slide.shapes) {
          for (const ref of shape.textStories) {
            const story = handle.story(ref.id);
            let base = 0;
            for (const para of story.paragraphs) {
              const text = para.runs.map((run) => run.text).join("");
              const at = text.indexOf(marker);
              if (at !== -1) {
                handle.deleteText(ref.id, base + at, base + at + marker.length);
                return;
              }
              base += text.length + 1;
            }
          }
        }
      }
      throw new Error("marker not found in pptx client");
    },
    contains: (marker) => JSON.stringify(handle.snapshot()).includes(marker),
    save: async () => handle.save(),
    destroy: () => handle.dispose(),
  };
}

/** Client ids present as keys in an encoded Yrs v1 state vector. */
export function stateVectorClientIds(stateVector: Uint8Array): number[] {
  let offset = 0;
  const readVar = (): number => {
    let value = 0n;
    let shift = 0n;
    for (;;) {
      const byte = stateVector[offset]!;
      offset += 1;
      value |= BigInt(byte & 0x7f) << shift;
      if (byte < 0x80) return Number(value);
      shift += 7n;
    }
  };
  const count = readVar();
  const ids: number[] = [];
  for (let i = 0; i < count; i += 1) {
    ids.push(readVar());
    readVar();
  }
  return ids;
}

/** A client hydrated from a server join payload, the way the spike built them. */
export async function openTestClient(
  payload: JoinPayload,
  source: Uint8Array,
  clientId: number,
): Promise<TestClient> {
  if (payload.format === "docx") {
    const session = await createYrsSession({ clientId });
    session.openDocx(source.slice(), false);
    session.loadState(payload.initialUpdate.slice());
    return wrapDocx(session);
  }
  if (payload.format === "xlsx") {
    const handle = openWorkbook(source.slice(), { collaborative: true, clientId });
    handle.applyUpdate(payload.initialUpdate.slice());
    return wrapXlsx(handle);
  }
  const handle = openPresentation(source.slice(), {
    clientId,
    initialUpdate: payload.initialUpdate.slice(),
  });
  return wrapPptx(handle);
}

/** A fresh, non-collaborative open of saved bytes, for save round-trip checks. */
export async function openFreshClient(
  format: Format,
  saved: Uint8Array,
  clientId: number,
): Promise<TestClient> {
  if (format === "docx") {
    const session = await createYrsSession({ clientId });
    session.openDocx(saved.slice(), true);
    return wrapDocx(session);
  }
  if (format === "xlsx") return wrapXlsx(openWorkbook(saved.slice()));
  return wrapPptx(openPresentation(saved.slice(), { clientId }));
}
