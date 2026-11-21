import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";

const docsPath = join(import.meta.dir, "..", "..", "..", "apps", "docs", "content", "docs", "wopi.mdx");

const HOST_OWNED_WOPI_VARS = new Set(["WOPI_ACCESS_TOKEN_EXPIRATION"]);
const REPOSITORY_CI_VARS = new Set(["WOPI_IMAGE_PUBLISH"]);

function wopiVars(source: string): Set<string> {
  return new Set(source.match(/(?<![A-Z0-9_])WOPI_[A-Z][A-Z0-9_]*/g) ?? []);
}

function sourceVars(): Set<string> {
  const vars = new Set<string>();
  for (const name of ["config.ts", "server.ts"]) {
    for (const entry of wopiVars(readFileSync(join(import.meta.dir, name), "utf8"))) vars.add(entry);
  }
  return vars;
}

describe("wopi admin docs", () => {
  test("documents every env var read in config.ts and server.ts and no unknown app vars", () => {
    const expected = sourceVars();
    expect(expected.size).toBeGreaterThan(0);
    const docs = readFileSync(docsPath, "utf8");
    const docsVars = wopiVars(docs);
    for (const hostVar of HOST_OWNED_WOPI_VARS) docsVars.delete(hostVar);
    const paragraphs = docs.split(/\n\n+/);
    for (const ciVar of REPOSITORY_CI_VARS) {
      expect(expected.has(ciVar)).toBe(false);
      const mentioning = paragraphs.filter((paragraph) => wopiVars(paragraph).has(ciVar));
      expect(mentioning).toHaveLength(1);
      expect(mentioning[0]).toContain("The published container image is");
      docsVars.delete(ciVar);
    }
    expect([...docsVars].sort()).toEqual([...expected].sort());
    expect(docs).toContain("`PORT`");
  });

  test("docs cover the required operational topics without dashes or tokens", () => {
    const docs = readFileSync(docsPath, "utf8");
    expect(docs.includes("\u2014")).toBe(false);
    expect(docs.includes("\u2013")).toBe(false);
    expect(docs).not.toMatch(/access_token=[A-Za-z0-9]/);
    for (const topic of [
      "512 MiB",
      "WOPI_PROOF_OLD_KEY_FILE",
      "WOPI_JOURNAL_DIR",
      "Strict-Transport-Security",
      "affinity cookie",
      "allow_local_remote_servers",
      "COLLABORATION_APP_PRODUCT",
      "COLLABORATION_WOPI_SRC",
      "richdocuments",
      "OFFICE_WEB_APP_BASE_URL",
      "ENABLE_OFFICE_WEB_APP_EDIT",
      "sticky",
      "/healthz",
      "docker build",
    ]) {
      expect(docs).toContain(topic);
    }
  });
});
