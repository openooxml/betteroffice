import { expect, spyOn, test } from "bun:test";
import { closeDocument, finishClose, saveDocument } from "./documentActions";
import type { SaveBarrier } from "../../src/relay/sequence";

test.each(["flush", "vector", "save"])("rejoining rejects a pending %s without waiting for its deadline", async (stage) => {
  const h = setup();
  const controller = new AbortController();
  let entered!: () => void;
  let release!: () => void;
  const started = new Promise<void>((resolve) => { entered = resolve; });
  const stalled = new Promise<void>((resolve) => { release = resolve; });
  if (stage === "flush") h.editor.flushPendingInput = () => { entered(); return stalled; };
  if (stage === "vector") h.editor.encodeStateVector = async () => { entered(); await stalled; return Uint8Array.of(9); };
  let requestSignal: AbortSignal | undefined;
  const api = {
    save: async (barrier: SaveBarrier, signal?: AbortSignal) => {
      requestSignal = signal;
      if (stage === "save") { entered(); await stalled; }
      return h.api.save(barrier);
    },
  };
  const warning = spyOn(console, "error").mockImplementation(() => {});
  const pending = saveDocument(h.editor, api, { isModified: h.isModified, signal: controller.signal });
  const outcome = pending.then(() => "saved", (error: Error) => error.message);
  let timeout: ReturnType<typeof setTimeout> | undefined;
  try {
    await started;
    controller.abort(new Error("rejoinRequired"));
    expect(await Promise.race([
      outcome,
      new Promise<string>((resolve) => { timeout = setTimeout(() => resolve("still pending"), 50); }),
    ])).toBe("rejoinRequired");
    if (stage === "save") expect(requestSignal?.aborted).toBe(true);
    release();
    await pending.catch(() => {});
    if (stage !== "save") expect(h.calls).not.toContain("save");
  } finally {
    clearTimeout(timeout);
    release();
    await pending.catch(() => {});
    warning.mockRestore();
  }
});

function setup() {
  const calls: string[] = [];
  let dirty = false;
  const editor = {
    async flushPendingInput() { calls.push("flush"); dirty = true; },
    async encodeStateVector() { calls.push("vector"); return Uint8Array.of(9); },
  };
  const api = {
    async save(barrier: SaveBarrier) { expect(barrier).toEqual(Uint8Array.of(9)); calls.push("save"); return { itemVersion: "2" }; },
    async close() { calls.push("close"); },
  };
  const isModified = () => dirty;
  return { calls, editor, api, isModified, closeOptions: { isModified, setInputBlocked() {} } };
}

test("save flushes pending input before capturing the ordered save vector", async () => {
  const h = setup();
  await saveDocument(h.editor, h.api, { isModified: h.isModified });
  expect(h.calls).toEqual(["flush", "vector", "save"]);
});

test.each(["flushPendingInput", "encodeStateVector"] as const)("a stalled docx %s fails without posting a save", async (method) => {
  const h = setup();
  h.editor[method] = () => new Promise<never>(() => {});
  const progress: Record<string, unknown>[] = [];
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    await expect(saveDocument(h.editor, h.api, {
      isModified: h.isModified,
      inputTimeoutMs: 5,
      onProgress: (patch) => { progress.push(patch); },
    })).rejects.toThrow(`${method}Timeout`);
    expect(h.calls).not.toContain("save");
    expect(warning).toHaveBeenCalledTimes(1);
    expect(progress.at(-1)).toMatchObject({ saveStep: "result", saveResult: "failed", saveError: `${method}Timeout` });
  } finally {
    warning.mockRestore();
  }
});

test("a throwing docx state vector API reports its failure", async () => {
  const h = setup();
  h.editor.encodeStateVector = () => { throw new Error("worker-hosted replica"); };
  const progress: Record<string, unknown>[] = [];
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    await expect(saveDocument(h.editor, h.api, {
      isModified: h.isModified,
      onProgress: (patch) => { progress.push(patch); },
    })).rejects.toThrow("worker-hosted replica");
    expect(h.calls).toEqual(["flush"]);
    expect(progress.at(-1)).toMatchObject({ saveStep: "result", saveResult: "failed", saveFailedStep: "stateVector" });
  } finally {
    warning.mockRestore();
  }
});

test("save diagnostics record every stage and capture changes committed by the flush", async () => {
  const h = setup();
  const progress: Record<string, unknown>[] = [];
  let modifiedAfterFlush = false;
  await saveDocument(h.editor, h.api, {
    isModified: h.isModified,
    onFlushed: () => { modifiedAfterFlush = h.isModified(); },
    onProgress: (patch) => { progress.push(patch); },
  });
  expect(modifiedAfterFlush).toBe(true);
  expect(progress.map((patch) => patch.saveStep)).toEqual(["started", "flushed", "stateVector", "posted", "result"]);
  expect(progress[2]).toMatchObject({ saveStateVectorBytes: 1 });
  expect(progress.at(-1)).toMatchObject({ saveResult: "saved" });
});

test("a stalled save response times out and aborts the request", async () => {
  const h = setup();
  let signal: AbortSignal | undefined;
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    await expect(saveDocument(h.editor, {
      save(_barrier: SaveBarrier, requestSignal?: AbortSignal) {
        signal = requestSignal;
        return new Promise<never>(() => {});
      },
    }, { isModified: h.isModified, saveTimeoutMs: 5 })).rejects.toThrow("saveTimeout");
    expect(signal?.aborted).toBe(true);
    expect(h.isModified()).toBe(true);
  } finally {
    warning.mockRestore();
  }
});

test("close saves an edit committed by the final flush before closing", async () => {
  const h = setup();
  await closeDocument(h.editor, h.api, h.closeOptions, () => { h.calls.push("host-close"); });
  expect(h.calls).toEqual(["flush", "vector", "save", "close", "host-close"]);
});

test("failed final saves leave the session and host open", async () => {
  const h = setup();
  await expect(closeDocument(h.editor, {
    async save() { h.calls.push("save"); throw new Error("notSynced"); }, close: () => h.api.close(),
  }, h.closeOptions, () => { h.calls.push("host-close"); })).rejects.toThrow("notSynced");
  expect(h.calls).toEqual(["flush", "vector", "save"]);
});

test("failed input flush does not close or save", async () => {
  const h = setup();
  h.editor.flushPendingInput = async () => { throw new Error("input failed"); };
  await expect(closeDocument(h.editor, { save: async () => {}, close: () => h.api.close() }, h.closeOptions, () => {})).rejects.toThrow("input failed");
  expect(h.calls).toEqual([]);
});

test("unmodified close still performs an ordered save", async () => {
  const h = setup();
  h.editor.flushPendingInput = async () => { h.calls.push("flush"); };
  await closeDocument(h.editor, h.api, h.closeOptions, () => {});
  expect(h.calls).toEqual(["flush", "vector", "save", "close"]);
});

test("host save options preserve editing and skip unmodified saves", async () => {
  const h = setup();
  await saveDocument(h.editor, h.api, { DontTerminateEdit: true, DontSaveIfUnmodified: true, isModified: h.isModified });
  expect(h.calls).toEqual([]);
});

test("close selects UI_Close only when requested and otherwise navigates the top window", () => {
  const calls: string[] = [];
  const win = { top: { location: { href: "" } } };
  const bridge = { enabled: true, uiClose() { calls.push("UI_Close"); } };
  const closed = () => { calls.push("closed"); };
  finishClose({ closePostMessage: false, closeUrl: "https://host.example/close" }, bridge, win, closed);
  expect(win.top.location.href).toBe("https://host.example/close");
  expect(calls).toEqual([]);
  finishClose({ closePostMessage: true, closeUrl: "https://host.example/other" }, bridge, win, closed);
  expect(calls).toEqual(["UI_Close"]);
  expect(win.top.location.href).toBe("https://host.example/close");
  finishClose({ closePostMessage: true, closeUrl: "" }, { ...bridge, enabled: false }, win, closed);
  expect(calls).toEqual(["UI_Close", "closed"]);
});

test("a failed close request leaves the host open", async () => {
  const h = setup();
  await expect(closeDocument(h.editor, {
    save: h.api.save,
    async close() { h.calls.push("close"); throw new Error("Close failed"); },
  }, h.closeOptions, () => { h.calls.push("host-close"); })).rejects.toThrow("Close failed");
  expect(h.calls).toEqual(["flush", "vector", "save", "close"]);
});

test("save captures the socket barrier immediately after input flush", async () => {
  const h = setup();
  const connectionId = crypto.randomUUID();
  let seq = 0;
  h.editor.flushPendingInput = async () => { h.calls.push("flush"); seq += 1; };
  await saveDocument(h.editor, {
    async save(barrier: SaveBarrier) {
      expect(barrier).toEqual({ connectionId, seq: 1 });
      h.calls.push("save");
    },
  }, {
    isModified: h.isModified,
    captureBarrier: () => ({ connectionId, seq }),
  });
  expect(h.calls).toEqual(["flush", "save"]);
});

test("close waits for a delayed update when idle modified notifications are withheld", async () => {
  const h = setup();
  const barrier = { connectionId: crypto.randomUUID(), seq: 1 };
  let apply!: () => void;
  const applied = new Promise<void>((resolve) => { apply = resolve; });
  h.editor.flushPendingInput = async () => { h.calls.push("flush"); };
  const closing = closeDocument(h.editor, {
    async save(captured: SaveBarrier) {
      expect(captured).toEqual(barrier);
      h.calls.push("save");
      await applied;
      h.calls.push("applied");
    },
    close: h.api.close,
  }, { ...h.closeOptions, captureBarrier: () => ({ ...barrier }) }, () => { h.calls.push("host-close"); });
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(h.isModified()).toBe(false);
  expect(h.calls).toEqual(["flush", "save"]);
  apply();
  await closing;
  expect(h.calls).toEqual(["flush", "save", "applied", "close", "host-close"]);
});

test("a stalled closing flush restores input and reports the error without saving or closing", async () => {
  const h = setup();
  let settle!: () => void;
  h.editor.flushPendingInput = () => new Promise<void>((resolve) => { settle = resolve; });
  const blocked: boolean[] = [];
  let error: unknown;
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    await expect(closeDocument(h.editor, h.api, {
      ...h.closeOptions, inputTimeoutMs: 5,
      setInputBlocked: (value) => { blocked.push(value); },
      onFailure: (failure) => { error = failure; },
    }, () => { h.calls.push("host-close"); })).rejects.toThrow("flushPendingInputTimeout");
    expect(blocked).toEqual([true, false]);
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toBe("flushPendingInputTimeout");
    expect(h.calls).toEqual([]);
    settle();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(h.calls).toEqual([]);
  } finally {
    warning.mockRestore();
  }
});

test("close saves once through its captured barrier when every save triggers a remote edit", async () => {
  const h = setup();
  const connectionId = crypto.randomUUID();
  let seq = 1;
  let captures = 0;
  await closeDocument(h.editor, {
    async save(barrier: SaveBarrier) {
      expect(barrier).toEqual({ connectionId, seq: 1 });
      h.calls.push("save");
      seq += 1;
    },
    close: h.api.close,
  }, {
    ...h.closeOptions,
    captureBarrier: () => { captures += 1; return { connectionId, seq }; },
  }, () => { h.calls.push("host-close"); });
  expect(h.isModified()).toBe(true);
  expect(captures).toBe(1);
  expect(h.calls).toEqual(["flush", "save", "close", "host-close"]);
});
