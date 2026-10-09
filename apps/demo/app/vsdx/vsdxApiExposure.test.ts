import { expect, test } from "bun:test";
import { clearVsdxApi, exposeVsdxApi, shouldExposeVsdxApi } from "./vsdxApiExposure";

test("the debug api stays off the window unless explicitly requested", () => {
  expect(shouldExposeVsdxApi("")).toBe(false);
  expect(shouldExposeVsdxApi("?room=a")).toBe(false);
  expect(shouldExposeVsdxApi("?vsdxApi=0")).toBe(false);
  expect(shouldExposeVsdxApi("?vsdxApi=1")).toBe(true);
});

test("clearing removes only the api instance it installed", () => {
  const target: { __vsdxApi?: unknown } = {};
  const first = { handle: 1 };
  exposeVsdxApi(target, first as never);
  expect(target.__vsdxApi).toBe(first);
  clearVsdxApi(target, { handle: 2 } as never);
  expect(target.__vsdxApi).toBe(first);
  clearVsdxApi(target, first as never);
  expect("__vsdxApi" in target).toBe(false);
});
