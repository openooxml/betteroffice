import { expect, test } from "bun:test";
import { mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";

test("benchmark generation leaves source fixtures untouched", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "vsdx-fixture-isolation-"));
  try {
    const scripts = resolve(directory, "scripts");
    const fixtures = resolve(directory, "crates/vsdx-parse/tests/fixtures");
    await mkdir(scripts, { recursive: true });
    await mkdir(fixtures, { recursive: true });
    await symlink(resolve(import.meta.dir, "../node_modules"), resolve(directory, "node_modules"));
    await writeFile(resolve(scripts, "create-vsdx-fixture.ts"), await readFile(resolve(import.meta.dir, "create-vsdx-fixture.ts")));
    await writeFile(resolve(fixtures, "foundation.vsdx"), "local fixture edits");
    const child = Bun.spawn([process.execPath, "scripts/create-vsdx-fixture.ts", "--benchmark-dir=bench"], {
      cwd: directory,
      stdout: "pipe",
      stderr: "pipe",
    });
    const [exitCode, stderr] = await Promise.all([child.exited, new Response(child.stderr).text()]);
    expect(stderr).toBe("");
    expect(exitCode).toBe(0);
    expect(await readdir(fixtures)).toEqual(["foundation.vsdx"]);
    expect(await readFile(resolve(fixtures, "foundation.vsdx"), "utf8")).toBe("local fixture edits");
    expect((await readdir(resolve(directory, "bench"))).sort()).toEqual([
      "deep-inheritance.vsdx",
      "dense-glue.vsdx",
      "many-pages.vsdx",
      "nested-groups.vsdx",
      "shape-heavy.vsdx",
      "text-heavy.vsdx",
    ]);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
