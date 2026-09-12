import { expect, test } from "bun:test";
import { liveFormats } from "./formats.ts";

test("middleware matcher literal covers the homepage and every live format", async () => {
  const source = await Bun.file(new URL("../middleware.ts", import.meta.url)).text();
  const match = source.match(/export const config = \{ matcher: (\[[^\n]+\]) \};/);

  expect(match).not.toBeNull();
  expect(JSON.parse(match[1])).toEqual([
    "/",
    ...liveFormats.map((format) => `/${format.id}`),
  ]);
});
