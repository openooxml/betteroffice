import { expect, test } from "bun:test";
import { closeDocument, finishClose, saveDocument } from "./documentActions";

function setup() {
  const calls: string[] = [];
  let dirty = false;
  const editor = {
    async flushPendingInput() { calls.push("flush"); dirty = true; },
    async encodeStateVector() { calls.push("vector"); return Uint8Array.of(9); },
  };
  const api = {
    async save(vector: Uint8Array) { expect(vector).toEqual(Uint8Array.of(9)); calls.push("save"); return { itemVersion: "2" }; },
    async close() { calls.push("close"); },
  };
  return { calls, editor, api, isModified: () => dirty };
}

test("save flushes pending input before capturing the ordered save vector", async () => {
  const h = setup();
  await saveDocument(h.editor, h.api, { isModified: h.isModified });
  expect(h.calls).toEqual(["flush", "vector", "save"]);
});

test("close saves an edit committed by the final flush before closing", async () => {
  const h = setup();
  await closeDocument(h.editor, {
    save: () => saveDocument(h.editor, h.api, { DontTerminateEdit: true, isModified: h.isModified }),
    close: () => h.api.close(),
  }, h.isModified, () => { h.calls.push("host-close"); });
  expect(h.calls).toEqual(["flush", "vector", "save", "close", "host-close"]);
});

test("failed final saves leave the session and host open", async () => {
  const h = setup();
  await expect(closeDocument(h.editor, {
    async save() { h.calls.push("save"); throw new Error("notSynced"); }, close: () => h.api.close(),
  }, h.isModified, () => { h.calls.push("host-close"); })).rejects.toThrow("notSynced");
  expect(h.calls).toEqual(["flush", "save"]);
});

test("failed input flush does not close or save", async () => {
  const h = setup();
  h.editor.flushPendingInput = async () => { throw new Error("input failed"); };
  await expect(closeDocument(h.editor, { save: async () => {}, close: () => h.api.close() }, h.isModified, () => {})).rejects.toThrow("input failed");
  expect(h.calls).toEqual([]);
});

test("unmodified close skips saving", async () => {
  const h = setup();
  h.editor.flushPendingInput = async () => { h.calls.push("flush"); };
  await closeDocument(h.editor, { save: async () => { h.calls.push("save"); }, close: () => h.api.close() }, h.isModified, () => {});
  expect(h.calls).toEqual(["flush", "close"]);
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
    save: () => saveDocument(h.editor, h.api, { DontTerminateEdit: true, isModified: h.isModified }),
    async close() { h.calls.push("close"); throw new Error("Close failed"); },
  }, h.isModified, () => { h.calls.push("host-close"); })).rejects.toThrow("Close failed");
  expect(h.calls).toEqual(["flush", "vector", "save", "close"]);
});
