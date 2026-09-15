import { expect, test } from "bun:test";
import {
  MAX_AWARENESS_PAYLOAD_BYTES,
  MAX_COLLABORATION_FRAME_BYTES,
  MAX_JOIN_REPLAY_BYTES,
  MAX_RETAINED_HISTORY_BYTES,
} from "../../../shared/collaboration-limits";
import { DEFAULT_MAX_FRAME_BYTES as DOCX_MAX_FRAME_BYTES } from "../../../packages/docx/src/collaboration/protocol";
import { encodeAwarenessMessage as encodeDocxAwareness } from "../../../packages/docx/src/collaboration/protocol";
import { DEFAULT_MAX_FRAME_BYTES as PPTX_MAX_FRAME_BYTES } from "../../../packages/pptx/src/collaboration/protocol";
import { encodeAwarenessUpdate as encodePptxAwareness } from "../../../packages/pptx/src/collaboration/protocol";
import { DEFAULT_MAX_FRAME_BYTES as XLSX_MAX_FRAME_BYTES } from "../../../packages/xlsx/src/collaboration/protocol";
import { encodeAwarenessUpdate as encodeXlsxAwareness } from "../../../packages/xlsx/src/collaboration/protocol";
import { DEFAULT_MAX_FRAME_BYTES as VSDX_MAX_FRAME_BYTES } from "../../../packages/vsdx/src/collaboration/protocol";
import { encodeAwarenessUpdate as encodeVsdxAwareness } from "../../../packages/vsdx/src/collaboration/protocol";

test("every client caps frames at the relay's ingress limit", () => {
  expect(DOCX_MAX_FRAME_BYTES).toBe(MAX_COLLABORATION_FRAME_BYTES);
  expect(PPTX_MAX_FRAME_BYTES).toBe(MAX_COLLABORATION_FRAME_BYTES);
  expect(XLSX_MAX_FRAME_BYTES).toBe(MAX_COLLABORATION_FRAME_BYTES);
  expect(VSDX_MAX_FRAME_BYTES).toBe(MAX_COLLABORATION_FRAME_BYTES);
});

test("one maximum-size frame cannot evict the whole retained history", () => {
  expect(MAX_RETAINED_HISTORY_BYTES).toBeGreaterThanOrEqual(
    4 * MAX_COLLABORATION_FRAME_BYTES,
  );
});

test("join replay is bounded by the retained budget", () => {
  expect(MAX_JOIN_REPLAY_BYTES).toBeLessThanOrEqual(MAX_RETAINED_HISTORY_BYTES);
});

test("every provider's largest awareness frame fits the relay awareness cap", () => {
  const docx = encodeDocxAwareness([
    {
      clientId: 1,
      clock: 1,
      state: {
        user: { name: "n".repeat(1024), color: "#0B57D0" },
        cursor: {
          story: "s".repeat(1024),
          anchor: new Uint8Array(1024).fill(255),
          head: new Uint8Array(1024).fill(255),
        },
      },
    },
  ]);
  const xlsx = encodeXlsxAwareness([
    {
      clientId: 1,
      clock: 1,
      state: {
        user: { name: "n".repeat(128), color: "#0B57D0" },
        cursor: {
          sheet: "s".repeat(256),
          anchor: { row: 999999999, col: 999999999 },
          head: { row: 999999999, col: 999999999 },
        },
      },
    },
  ]);
  const pptx = encodePptxAwareness([
    {
      clientId: 1,
      clock: 1,
      state: {
        clientId: 1,
        clock: 1,
        user: { name: "n".repeat(1024), color: "#B3261E" },
        cursor: { slideId: "s".repeat(1024), shapeId: "h".repeat(1024) },
      },
    },
  ]);
  const vsdx = encodeVsdxAwareness([
    {
      clientId: 1,
      clock: 1,
      state: {
        clientId: 1,
        clock: 1,
        user: { name: "n".repeat(1024), color: "#B3261E" },
        cursor: { pageId: "s".repeat(1024), shapeId: "h".repeat(1024) },
      },
    },
  ]);
  for (const frame of [docx, xlsx, pptx, vsdx]) {
    expect(frame.byteLength).toBeLessThanOrEqual(MAX_AWARENESS_PAYLOAD_BYTES);
  }
});
