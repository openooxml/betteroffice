import { describe, expect, spyOn, test } from "bun:test";

import { saveDocument } from "./documentActions";
import { PostMessageBridge, type PostMessageWindow, type WopiView } from "./postMessage";

const ORIGIN = "https://cloud.example.com";

class FakeWindow implements PostMessageWindow {
  sent: Array<{ message: string; origin: string }> = [];
  listeners = new Set<(event: MessageEvent) => void>();
  parent = {
    postMessage: (message: string, origin: string): void => {
      this.sent.push({ message, origin });
    },
  };

  addEventListener(_type: "message", listener: (event: MessageEvent) => void): void {
    this.listeners.add(listener);
  }

  removeEventListener(_type: "message", listener: (event: MessageEvent) => void): void {
    this.listeners.delete(listener);
  }

  emit(event: { origin: string; data: unknown }): void {
    for (const listener of [...this.listeners]) listener(event as MessageEvent);
  }

  parsed(): Array<{ MessageId: string; SendTime: number; Values: unknown }> {
    return this.sent.map((entry) => JSON.parse(entry.message));
  }
}

function setup(origin = ORIGIN) {
  const win = new FakeWindow();
  const calls: string[] = [];
  const views: WopiView[] = [{ ViewId: 11, UserName: "Alice A" }, { ViewId: 22, UserName: "Bob B" }];
  const bridge = new PostMessageBridge(win, origin, {
    onSave: () => { calls.push("save"); },
    onClose: () => calls.push("close"),
    onBlurFocus: () => calls.push("blur"),
    onGrabFocus: () => calls.push("grab"),
    onGetViews: () => views,
    onHostReady: () => calls.push("ready"),
  }, () => 1700000000000);
  bridge.start();
  return { win, bridge, calls, views };
}

function hostMessage(messageId: string, values: Record<string, unknown> = {}): string {
  return JSON.stringify({ MessageId: messageId, SendTime: 1700000000000, Values: values });
}

describe("PostMessage bridge sending", () => {
  test("sends JSON strings only to the configured origin, never *", () => {
    const { win, bridge } = setup();
    bridge.frameReady();
    bridge.documentLoaded(1700000001234);
    bridge.modified(true);
    bridge.modified(false);
    bridge.uiClose();
    expect(win.sent).toHaveLength(5);
    for (const entry of win.sent) {
      expect(entry.origin).toBe(ORIGIN);
      expect(entry.origin).not.toBe("*");
      expect(typeof entry.message).toBe("string");
    }
    const messages = win.parsed();
    expect(messages[0]).toEqual({ MessageId: "App_LoadingStatus", SendTime: 1700000000000, Values: { Status: "Frame_Ready" } });
    expect(messages[1]).toEqual({ MessageId: "App_LoadingStatus", SendTime: 1700000000000, Values: { Status: "Document_Loaded", DocumentLoadedTime: 1700000001234 } });
    expect(messages[2]!.Values).toEqual({ Modified: true });
    expect(messages[3]!.Values).toEqual({ Modified: false });
    expect(messages[4]!.MessageId).toBe("UI_Close");
  });

  test("boot failure posts a Failed loading status", () => {
    const { win, bridge } = setup();
    bridge.bootFailed();
    expect(win.parsed()).toEqual([
      { MessageId: "App_LoadingStatus", SendTime: 1700000000000, Values: { Status: "Failed" } },
    ]);
  });

  test("sends nothing when the origin is empty", () => {
    const { win, bridge } = setup("");
    expect(bridge.enabled).toBe(false);
    bridge.frameReady();
    bridge.documentLoaded();
    bridge.uiClose();
    expect(win.sent).toEqual([]);
  });

  test("view added and removed carry the view", () => {
    const { win, bridge } = setup();
    bridge.viewAdded({ ViewId: 7, UserName: "Alice A" });
    bridge.viewRemoved({ ViewId: 7, UserName: "Alice A" });
    const messages = win.parsed();
    expect(messages[0]).toMatchObject({ MessageId: "View_Added", Values: { ViewId: 7, UserName: "Alice A" } });
    expect(messages[1]).toMatchObject({ MessageId: "View_Removed", Values: { ViewId: 7 } });
  });
});

describe("PostMessage bridge receiving", () => {
  test("dispatches every known host action", () => {
    const { win, calls } = setup();
    win.emit({ origin: ORIGIN, data: hostMessage("Host_PostmessageReady") });
    win.emit({ origin: ORIGIN, data: hostMessage("Action_Save") });
    win.emit({ origin: ORIGIN, data: hostMessage("Action_Close") });
    win.emit({ origin: ORIGIN, data: hostMessage("Blur_Focus") });
    win.emit({ origin: ORIGIN, data: hostMessage("Grab_Focus") });
    expect(calls).toEqual(["ready", "save", "close", "blur", "grab"]);
  });

  test("answers Get_Views with every current view", () => {
    const { win } = setup();
    win.emit({ origin: ORIGIN, data: hostMessage("Get_Views") });
    const replies = win.parsed().filter((message) => message.MessageId === "Get_Views_Resp");
    expect(replies).toHaveLength(1);
    expect(replies[0]!.Values).toEqual([{ ViewId: 11, UserName: "Alice A" }, { ViewId: 22, UserName: "Bob B" }]);
  });

  test("rejects messages from other origins", () => {
    const { win, calls } = setup();
    win.emit({ origin: "https://evil.example.com", data: hostMessage("Action_Save") });
    win.emit({ origin: "", data: hostMessage("Action_Save") });
    expect(calls).toEqual([]);
  });

  test("rejects malformed shapes and unknown ids", () => {
    const { win, calls } = setup();
    win.emit({ origin: ORIGIN, data: "not json" });
    win.emit({ origin: ORIGIN, data: JSON.stringify({ MessageId: 42 }) });
    win.emit({ origin: ORIGIN, data: JSON.stringify({ MessageId: "Action_Save", SendTime: "soon" }) });
    win.emit({ origin: ORIGIN, data: JSON.stringify({ MessageId: "Action_Save", Values: "x" }) });
    win.emit({ origin: ORIGIN, data: JSON.stringify({ MessageId: "Action_DeleteEverything" }) });
    win.emit({ origin: ORIGIN, data: 42 });
    win.emit({ origin: ORIGIN, data: null });
    expect(calls).toEqual([]);
  });

  test("accepts object form as well as JSON strings", () => {
    const { win, calls } = setup();
    win.emit({ origin: ORIGIN, data: { MessageId: "Action_Save", SendTime: 1, Values: {} } });
    expect(calls).toEqual(["save"]);
  });

  test("a star origin disables the bridge entirely", () => {
    const { win, bridge, calls } = setup("*");
    expect(bridge.enabled).toBe(false);
    bridge.frameReady();
    win.emit({ origin: "*", data: hostMessage("Action_Save") });
    expect(win.sent).toEqual([]);
    expect(calls).toEqual([]);
  });

  test("an unnormalized origin is normalized to its url origin", () => {
    const { win, bridge } = setup("https://cloud.example.com/");
    expect(bridge.enabled).toBe(true);
    bridge.frameReady();
    expect(win.sent[0]!.origin).toBe("https://cloud.example.com");
  });

  test("rejects NaN send times and array values", () => {
    const { win, calls } = setup();
    win.emit({ origin: ORIGIN, data: { MessageId: "Action_Save", SendTime: Number.NaN, Values: {} } });
    win.emit({ origin: ORIGIN, data: { MessageId: "Action_Save", SendTime: 1, Values: [1, 2, 3] } });
    expect(calls).toEqual([]);
  });

  test("dispose stops dispatch", () => {
    const { win, bridge, calls } = setup();
    bridge.dispose();
    win.emit({ origin: ORIGIN, data: hostMessage("Action_Save") });
    expect(calls).toEqual([]);
  });
});

test("opening failure is sent once and prevents a later loaded message", () => {
  const { win, bridge } = setup();
  bridge.bootFailed();
  bridge.bootFailed();
  bridge.documentLoaded();
  expect(win.parsed().map((message) => message.Values)).toEqual([{ Status: "Failed" }]);
});

test("Action_Save Notify responds only after the ordered save succeeds", async () => {
  const win = new FakeWindow();
  let finish!: () => void;
  const saving = new Promise<void>((resolve) => { finish = resolve; });
  const bridge = new PostMessageBridge(win, ORIGIN, {
    onSave: () => saving, onClose() {}, onBlurFocus() {}, onGrabFocus() {}, onGetViews: () => [],
  });
  bridge.handleMessage({ origin: ORIGIN, data: hostMessage("Action_Save", { Notify: true }) });
  expect(win.parsed()).toEqual([]);
  finish();
  await saving;
  await Promise.resolve();
  expect(win.parsed()).toMatchObject([{ MessageId: "Action_Save_Resp", Values: { success: true } }]);
});

test("Action_Save failure notifies only when requested and forwards edit options", async () => {
  const win = new FakeWindow();
  const options: unknown[] = [];
  const bridge = new PostMessageBridge(win, ORIGIN, {
    onSave: async (values) => { options.push(values); throw new Error("notSynced"); },
    onClose() {}, onBlurFocus() {}, onGrabFocus() {}, onGetViews: () => [],
  });
  bridge.handleMessage({ origin: ORIGIN, data: hostMessage("Action_Save", { Notify: true, DontTerminateEdit: true, DontSaveIfUnmodified: true }) });
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(win.parsed()).toMatchObject([{ MessageId: "Action_Save_Resp", Values: { success: false, result: "notSynced" } }]);
  expect(options).toEqual([{ DontTerminateEdit: true, DontSaveIfUnmodified: true }]);
  bridge.handleMessage({ origin: ORIGIN, data: hostMessage("Action_Save") });
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(win.parsed()).toHaveLength(1);
});

test.each(["flushPendingInput", "encodeStateVector"] as const)("Action_Save Notify reports a stalled docx %s", async (method) => {
  const win = new FakeWindow();
  let saving: Promise<void> | undefined;
  let posted = false;
  const editor = {
    flushPendingInput: async () => {},
    encodeStateVector: async () => Uint8Array.of(0),
  };
  editor[method] = () => new Promise<never>(() => {});
  const bridge = new PostMessageBridge(win, ORIGIN, {
    onSave: (options) => {
      saving = saveDocument(editor, { save: async () => { posted = true; } }, {
        ...options, isModified: () => true, inputTimeoutMs: 5,
      });
      return saving;
    },
    onClose() {}, onBlurFocus() {}, onGrabFocus() {}, onGetViews: () => [],
  });
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    bridge.handleMessage({ origin: ORIGIN, data: hostMessage("Action_Save", { Notify: true }) });
    await saving?.catch(() => {});
    await Promise.resolve();
    expect(posted).toBe(false);
    expect(win.parsed()).toMatchObject([{
      MessageId: "Action_Save_Resp", Values: { success: false, result: `${method}Timeout` },
    }]);
  } finally {
    warning.mockRestore();
  }
});
