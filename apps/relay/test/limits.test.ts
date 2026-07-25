import { expect, test } from "bun:test";
import { MAX_COLLABORATION_FRAME_BYTES } from "../../../shared/collaboration-limits";
import { DEFAULT_MAX_FRAME_BYTES as DOCX_MAX_FRAME_BYTES } from "../../../packages/docx/src/collaboration/protocol";
import { DEFAULT_MAX_FRAME_BYTES as PPTX_MAX_FRAME_BYTES } from "../../../packages/pptx/src/collaboration/protocol";
import { DEFAULT_MAX_FRAME_BYTES as XLSX_MAX_FRAME_BYTES } from "../../../packages/xlsx/src/collaboration/protocol";

test("every client caps frames at the relay's ingress limit", () => {
  expect(DOCX_MAX_FRAME_BYTES).toBe(MAX_COLLABORATION_FRAME_BYTES);
  expect(PPTX_MAX_FRAME_BYTES).toBe(MAX_COLLABORATION_FRAME_BYTES);
  expect(XLSX_MAX_FRAME_BYTES).toBe(MAX_COLLABORATION_FRAME_BYTES);
});
