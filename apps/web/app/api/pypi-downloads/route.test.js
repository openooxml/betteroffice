import { describe, expect, test } from "bun:test";
import { PYTHON_PUBLISH_NAMES } from "../../../../../scripts/python-bindings.mjs";
import { PYPI_PACKAGES } from "../../../lib/pypi-downloads.ts";

describe("PyPI downloads", () => {
  test("tracks every published binding under its PyPI name", () => {
    expect(PYPI_PACKAGES).toEqual(PYTHON_PUBLISH_NAMES.map((name) => `betteroffice-${name}`));
  });
});
