import { expect, test } from "bun:test";
import type { FileSession } from "./session/fileSession";
import { decodeSaveBarrier, requestOrderedSave } from "./sessionApi";

test("legacy state vectors cannot authorize saves", () => {
  expect(() => decodeSaveBarrier(new TextEncoder().encode('{"stateVector":"AA=="}'))).toThrow("invalidSaveBarrier");
});

test("ordered saves wait for the captured socket sequence", async () => {
  let applied = 1;
  let saved = false;
  const session = { requestSave: async () => { saved = true; } } as unknown as FileSession;
  const barrier = { connectionId: crypto.randomUUID(), seq: 2 };
  const pending = requestOrderedSave(session, "user", barrier, {
    timeoutMs: 500, relay: { appliedSequence: () => applied },
  });
  await new Promise((resolve) => setTimeout(resolve, 60));
  expect(saved).toBe(false);
  applied = 2;
  await pending;
  expect(saved).toBe(true);
});

test("unknown or disconnected sockets cannot authorize saves", async () => {
  let saved = false;
  const session = { requestSave: async () => { saved = true; } } as unknown as FileSession;
  await expect(requestOrderedSave(session, "user", { connectionId: crypto.randomUUID(), seq: 0 }, {
    timeoutMs: 0, relay: { appliedSequence: () => null },
  })).rejects.toMatchObject({ kind: "notSynced" });
  expect(saved).toBe(false);
});
