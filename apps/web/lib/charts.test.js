import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseAgentEvals, agentEvalsSection } from "./agent-evals.ts";
import { benchmarkRows, benchmarksChartSvg, benchmarksHeadline, evalsChartSvg, evalsHeadline, themeOf } from "./charts.ts";
import { parseFidelity } from "./fidelity.ts";

const README = readFileSync(join(import.meta.dir, "..", "..", "..", "README.md"), "utf8");
const BASIC = parseAgentEvals(README, ["basic"]);
const COMPLEX = parseAgentEvals(README, ["complex"]);
const FORMATS = parseFidelity(README);

const altOf = (name) => README.match(new RegExp(`<img src="https://betteroffice\\.dev/api/charts/${name}\\.svg" alt="([^"]+)" width="720">`))?.[1];

describe("README chart embeds", () => {
  test("the data tables are still read from inside the collapsed blocks", () => {
    expect(BASIC.rows.map((row) => row.format)).toEqual(["docx", "xlsx", "pptx"]);
    expect(COMPLEX.rows.map((row) => row.format)).toEqual(["docx", "xlsx", "pptx"]);
    expect(FORMATS.map((f) => f.format)).toEqual(["docx", "pptx", "xlsx"]);
    for (const format of FORMATS) expect(format.rows.length).toBeGreaterThan(2);
    expect(agentEvalsSection(README)).toContain("<details>");
  });

  test("each chart is a picture with a dark source and a light image, and one sentence of alt text", () => {
    for (const name of ["evals", "benchmarks"]) {
      expect(README).toContain(`<source media="(prefers-color-scheme: dark)" srcset="https://betteroffice.dev/api/charts/${name}.svg?theme=dark">`);
      const alt = altOf(name);
      expect(alt).toBeTruthy();
      expect(alt.split(/\.\s/)).toHaveLength(1);
      expect(alt).not.toMatch(/[–—]/);
    }
    expect(README.match(/<details>\n<summary>Data<\/summary>/g)).toHaveLength(2);
    expect(README.match(/<\/details>/g)).toHaveLength(2);
  });

  test("alt text carries the numbers of the tables next to it", () => {
    expect(altOf("evals")).toBe(evalsHeadline(BASIC, COMPLEX));
    expect(altOf("benchmarks")).toBe(benchmarksHeadline(FORMATS));
  });
});

describe("chart SVGs", () => {
  const evals = (theme) => evalsChartSvg(BASIC, COMPLEX, theme);
  const benchmarks = (theme) => benchmarksChartSvg(FORMATS, theme);

  test("are well-formed, sized and carry every published value", () => {
    for (const svg of [evals("light"), benchmarks("light"), evals("dark"), benchmarks("dark")]) {
      expect(svg.startsWith("<svg xmlns=")).toBe(true);
      expect(svg.trimEnd().endsWith("</svg>")).toBe(true);
      expect(svg).not.toMatch(/undefined|NaN|[–—]/);
      expect(svg).toMatch(/<title>[^<]+<\/title>/);
    }
    const evalSvg = evals("light");
    expect(evalSvg.match(/<rect x=/g).length).toBeGreaterThanOrEqual(3 * 4);
    for (const text of ["79%", "98%", "96%", "45%", "94%", "58%"]) expect(evalSvg).toContain(`>${text}<`);
    const bench = benchmarks("light");
    for (const row of benchmarkRows(FORMATS)) {
      expect(bench).toContain(`>${row.ours.text}<`);
      expect(bench).toContain(`>${row.theirs.text}<`);
    }
  });

  test("basic bars are translucent with a full-colour outline, complex bars solid", () => {
    const svg = evals("light");
    expect(svg).toContain('fill-opacity="0.2"');
    expect(svg).toMatch(/fill="#047857" fill-opacity="1"/);
    expect(svg).toMatch(/fill-opacity="0.2" stroke="#047857"/);
  });

  test("the theme changes the palette and anything unknown is light", () => {
    expect(evals("dark")).toContain("#34d399");
    expect(evals("dark")).not.toContain("#047857");
    expect(themeOf("dark")).toBe("dark");
    expect(themeOf("sepia")).toBe("light");
    expect(themeOf(null)).toBe("light");
  });

  test("no data draws nothing", () => {
    expect(evalsChartSvg(null, null, "light")).toBeNull();
    expect(benchmarksChartSvg([], "light")).toBeNull();
  });
});

describe("chart routes", () => {
  process.env.BETTEROFFICE_AGENT_EVALS = agentEvalsSection(README);
  process.env.BETTEROFFICE_BENCHMARKS = README.slice(README.indexOf("<!-- BEGIN GENERATED VISUAL FIDELITY -->"), README.indexOf("<!-- END GENERATED VISUAL FIDELITY -->") + "<!-- END GENERATED VISUAL FIDELITY -->".length);

  test("serve cacheable SVG per theme", async () => {
    for (const name of ["evals", "benchmarks"]) {
      const { GET } = await import(`../app/api/charts/${name}.svg/route.ts`);
      const light = await GET(new Request(`https://betteroffice.dev/api/charts/${name}.svg`));
      expect(light.status).toBe(200);
      expect(light.headers.get("content-type")).toContain("image/svg+xml");
      expect(light.headers.get("cache-control")).toContain("s-maxage=3600");
      const dark = await GET(new Request(`https://betteroffice.dev/api/charts/${name}.svg?theme=dark`));
      expect(await dark.text()).not.toBe(await light.text());
    }
  });

  test("answer 404 without results", async () => {
    const saved = process.env.BETTEROFFICE_AGENT_EVALS;
    process.env.BETTEROFFICE_AGENT_EVALS = "";
    const { GET } = await import("../app/api/charts/evals.svg/route.ts");
    expect((await GET(new Request("https://betteroffice.dev/api/charts/evals.svg"))).status).toBe(404);
    process.env.BETTEROFFICE_AGENT_EVALS = saved;
  });
});
