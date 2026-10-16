import { expect, test } from "bun:test";

const demo = (await Bun.file(new URL("../app/globals.css", import.meta.url)).text())
  .replace(/\/\*[\s\S]*?\*\//g, "");
const shared = await Bun.file(new URL("../../../shared/site-tokens.css", import.meta.url)).text();

test("launcher inherits every shared palette token within the site scope", () => {
  expect(demo).toContain('@import "../../../shared/site-tokens.css";');
  const siteBlocks = [...demo.matchAll(/\.site\s*\{([^}]+)\}/g)];
  expect(siteBlocks).toHaveLength(1);
  const site = siteBlocks[0][1];
  const tokens = new Set([...shared.matchAll(/--bo-([\w-]+)\s*:/g)].map((match) => match[1]));

  for (const token of tokens) {
    const alias = token === "ease" ? "--ease" : `--color-${token}`;
    expect(site).toMatch(new RegExp(
      `(?:^|;)\\s*${alias}\\s*:\\s*var\\(\\s*--bo-${token}\\s*\\)\\s*(?=;|$)`,
    ));
  }
  expect(demo).not.toMatch(/--bo-[\w-]+\s*:/);
});

test("launcher uses the shared frame and registration marks without local overrides", () => {
  expect(shared).toMatch(/\.wrap\s*\{/);
  expect(shared).toMatch(/\.rule::before/);
  expect(demo).not.toMatch(/(?:^|})\s*\.(?:wrap|rule(?:::before|::after)?)\s*(?:,|\{)/);
});

test("launcher uses the shared background for the body in dark mode", () => {
  const darkBody = demo.match(
    /@media\s*\(\s*prefers-color-scheme\s*:\s*dark\s*\)\s*\{\s*html:has\(\s*\.site\s*\)\s+body\s*\{([^}]*)\}/,
  )?.[1];
  expect(darkBody).toMatch(/(?:^|;)\s*background\s*:\s*var\(\s*--bo-bg\s*\)\s*(?=;|$)/);
});
