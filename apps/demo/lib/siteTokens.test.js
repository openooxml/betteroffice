import { expect, test } from "bun:test";

const demo = await Bun.file(new URL("../app/globals.css", import.meta.url)).text();
const shared = await Bun.file(new URL("../../../shared/site-tokens.css", import.meta.url)).text();

test("launcher inherits every shared palette token within the site scope", () => {
  expect(demo).toContain('@import "../../../shared/site-tokens.css";');
  const siteBlocks = [...demo.matchAll(/\.site\s*\{([^}]+)\}/g)];
  expect(siteBlocks).toHaveLength(1);
  const site = siteBlocks[0][1];
  const tokens = new Set([...shared.matchAll(/--bo-([\w-]+)\s*:/g)].map((match) => match[1]));

  for (const token of tokens) {
    const alias = token === "ease" ? "--ease" : `--color-${token}`;
    expect(site).toContain(`${alias}: var(--bo-${token});`);
  }
  expect(demo).not.toMatch(/--bo-[\w-]+\s*:/);
});

test("launcher uses the shared frame and registration marks without local overrides", () => {
  expect(shared).toMatch(/\.wrap\s*\{/);
  expect(shared).toMatch(/\.rule::before/);
  expect(demo).not.toMatch(/\.(?:wrap|rule)\b/);
});
