import { expect, test } from "bun:test";
import { isFatalDocxError } from "./docxError";

test("recoverable errors leave the demo document mounted without a fatal banner", () => {
  expect(isFatalDocxError({ recoverable: true })).toBe(false);
});

test("fatal and legacy errors still replace the demo document with an error", () => {
  expect(isFatalDocxError()).toBe(true);
  expect(isFatalDocxError({ recoverable: false })).toBe(true);
});
