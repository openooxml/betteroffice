import { describe, expect, test } from "bun:test";
import { savedBytes } from "./session";

describe("editor host session", () => {
  test("a failed save reports its failure, not that the file is opening", () => {
    const failure = new Error("The document changed while saving");
    expect(() => savedBytes(null, failure)).toThrow(failure);
    expect(() => savedBytes(undefined, null)).toThrow("still opening");
    expect(savedBytes(Uint8Array.of(7).buffer, null)).toEqual(
      Uint8Array.of(7)
    );
  });

  test("an unfinished save uses the host's translated message", () => {
    expect(() => savedBytes(null, null, "Der Editor öffnet die Datei noch.")).toThrow("Der Editor öffnet die Datei noch.");
  });
});
