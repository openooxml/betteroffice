import { describe, expect, test } from "bun:test";
import { SessionRevision, savedBytes } from "./session";
import { encodeBytes, formatOf } from "./bridge";

describe("desktop document sessions", () => {
  test("a save cannot clear edits made while the dialog or write is pending", () => {
    const session = new SessionRevision();
    session.change();
    const checkpoint = session.checkpoint();
    session.change();
    session.didSave(checkpoint);
    expect(session.dirty).toBe(true);
    session.didSave(session.checkpoint());
    expect(session.dirty).toBe(false);
  });

  test("cancelling a save retains the dirty state", () => {
    const session = new SessionRevision();
    session.change();
    session.checkpoint();
    expect(session.dirty).toBe(true);
  });

  test("a failed save reports its failure, not that the file is opening", () => {
    const failure = new Error("The document changed while saving");
    expect(() => savedBytes(null, failure)).toThrow(failure);
    expect(() => savedBytes(undefined, null)).toThrow("still opening");
    expect(savedBytes(Uint8Array.of(7).buffer, null)).toEqual(
      Uint8Array.of(7)
    );
  });

  test("validates supported file types case insensitively", () => {
    expect(formatOf("Budget.2026.XLSX")).toBe("xlsx");
    expect(() => formatOf("document.docx.exe")).toThrow();
    expect(() => formatOf("document")).toThrow();
  });

  test("encodes large binary packages without overflowing argument limits", () => {
    const bytes = Uint8Array.from(
      { length: 150_000 },
      (_, index) => index % 256
    );
    const decoded = Uint8Array.from(atob(encodeBytes(bytes)), (value) =>
      value.charCodeAt(0)
    );
    expect(decoded).toEqual(bytes);
  });
});
