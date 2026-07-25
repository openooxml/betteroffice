import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import { createYrsSession } from "../../../packages/docx/src/yrs/index.ts";
import { openWorkbook } from "../../../packages/xlsx/src/index.ts";
import { openPresentation } from "../../../packages/pptx/src/index.ts";
import { buildCollaborationSeeds } from "./build-collaboration-seeds.ts";

const demo = resolve(import.meta.dir, "..");
const committedSeeds = resolve(demo, "public/seeds");

function equalBytes(left: Uint8Array, right: Uint8Array): boolean {
  return (
    left.byteLength === right.byteLength &&
    left.every((value, index) => value === right[index])
  );
}

function assertEqualBytes(
  format: string,
  value: string,
  committed: Uint8Array,
  generated: Uint8Array,
): void {
  if (equalBytes(committed, generated)) return;
  if (committed.byteLength !== generated.byteLength) {
    throw new Error(
      `${format} collaboration seed ${value} length differs: committed ${committed.byteLength}, generated ${generated.byteLength}`,
    );
  }
  const firstDifference = committed.findIndex(
    (byte, index) => byte !== generated[index],
  );
  throw new Error(
    `${format} collaboration seed ${value} differs at byte ${firstDifference}`,
  );
}

async function assertDocxEquivalent(
  committed: Uint8Array,
  generated: Uint8Array,
): Promise<number> {
  const committedSession = await createYrsSession({ clientId: 2 });
  const generatedSession = await createYrsSession({ clientId: 3 });
  try {
    committedSession.loadState(committed);
    generatedSession.loadState(generated);
    assertEqualBytes(
      "DOCX",
      "state vector",
      committedSession.encodeStateVector(),
      generatedSession.encodeStateVector(),
    );
    const storyIds = committedSession.storyIds();
    assert.deepStrictEqual(
      storyIds,
      generatedSession.storyIds(),
      "DOCX collaboration seed story IDs differ",
    );
    for (const storyId of storyIds) {
      assert.deepStrictEqual(
        committedSession.storySegments(storyId),
        generatedSession.storySegments(storyId),
        `DOCX collaboration seed story content differs for ${storyId}`,
      );
    }
    return storyIds.length;
  } finally {
    committedSession.destroy();
    generatedSession.destroy();
  }
}

function assertXlsxEquivalent(
  workbookBytes: Uint8Array,
  committed: Uint8Array,
  generated: Uint8Array,
): number {
  const committedWorkbook = openWorkbook(workbookBytes, {
    collaborative: true,
    clientId: 2,
  });
  const generatedWorkbook = openWorkbook(workbookBytes, {
    collaborative: true,
    clientId: 3,
  });
  try {
    committedWorkbook.applyUpdate(committed);
    generatedWorkbook.applyUpdate(generated);
    assertEqualBytes(
      "XLSX",
      "state vector",
      committedWorkbook.encodeStateVector(),
      generatedWorkbook.encodeStateVector(),
    );
    const sheetInfo = committedWorkbook.sheetInfo();
    assert.deepStrictEqual(
      sheetInfo,
      generatedWorkbook.sheetInfo(),
      "XLSX collaboration seed sheet metadata differs",
    );
    return sheetInfo.sheetNames.length;
  } finally {
    committedWorkbook.dispose();
    generatedWorkbook.dispose();
  }
}

function assertPptxEquivalent(
  presentationBytes: Uint8Array,
  committed: Uint8Array,
  generated: Uint8Array,
): number {
  const committedPresentation = openPresentation(presentationBytes, {
    clientId: 2,
    initialUpdate: committed,
  });
  const generatedPresentation = openPresentation(presentationBytes, {
    clientId: 3,
    initialUpdate: generated,
  });
  try {
    assertEqualBytes(
      "PPTX",
      "state vector",
      committedPresentation.encodeStateVector(),
      generatedPresentation.encodeStateVector(),
    );
    const snapshot = committedPresentation.snapshot();
    assert.deepStrictEqual(
      snapshot,
      generatedPresentation.snapshot(),
      "PPTX collaboration seed content differs",
    );
    return snapshot.slides.length;
  } finally {
    committedPresentation.dispose();
    generatedPresentation.dispose();
  }
}

function encodingStatus(committed: Uint8Array, generated: Uint8Array): string {
  return equalBytes(committed, generated)
    ? "byte-identical"
    : "different bytes, equivalent state";
}

const generatedSeeds = await mkdtemp(
  resolve(tmpdir(), "betteroffice-demo-seeds-"),
);

try {
  await buildCollaborationSeeds(generatedSeeds);
  const [
    committedDocx,
    generatedDocx,
    committedXlsx,
    generatedXlsx,
    committedPptx,
    generatedPptx,
    workbookBytes,
    presentationBytes,
  ] = await Promise.all([
    readFile(resolve(committedSeeds, "docx.bin")),
    readFile(resolve(generatedSeeds, "docx.bin")),
    readFile(resolve(committedSeeds, "xlsx.bin")),
    readFile(resolve(generatedSeeds, "xlsx.bin")),
    readFile(resolve(committedSeeds, "pptx.bin")),
    readFile(resolve(generatedSeeds, "pptx.bin")),
    readFile(resolve(demo, "public/showcase.xlsx")),
    readFile(resolve(demo, "public/betteroffice-demo.pptx")),
  ]);

  const storyCount = await assertDocxEquivalent(
    committedDocx,
    generatedDocx,
  );
  const sheetCount = assertXlsxEquivalent(
    workbookBytes,
    committedXlsx,
    generatedXlsx,
  );
  const slideCount = assertPptxEquivalent(
    presentationBytes,
    committedPptx,
    generatedPptx,
  );

  console.log(
    `DOCX seed: ${encodingStatus(committedDocx, generatedDocx)}, ${storyCount} stories`,
  );
  console.log(
    `XLSX seed: ${encodingStatus(committedXlsx, generatedXlsx)}, ${sheetCount} sheets`,
  );
  console.log(
    `PPTX seed: ${encodingStatus(committedPptx, generatedPptx)}, ${slideCount} slides`,
  );
} finally {
  await rm(generatedSeeds, { recursive: true, force: true });
}
