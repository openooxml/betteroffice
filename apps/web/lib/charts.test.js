import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseAgentEvals, agentEvalsSection } from "./agent-evals.ts";
import { benchmarksChartSvg, benchmarksHeadline, evalsChartSvg, evalsHeadline, evalsModel, perfChartSvg, perfHeadline, tableRows, themeOf } from "./charts.ts";
import { parseFidelity } from "./fidelity.ts";

const README = readFileSync(join(import.meta.dir, "..", "..", "..", "README.md"), "utf8");
const SECTION = agentEvalsSection(README);
const MODEL = evalsModel(SECTION);
const BASIC = parseAgentEvals(README, ["basic"]);
const COMPLEX = parseAgentEvals(README, ["complex"]);
const FORMATS = parseFidelity(README);

const altOf = (name) => README.match(new RegExp(`<img src="https://betteroffice\\.dev/api/charts/${name}\\.svg" alt="([^"]+)" width="100%">`))?.[1];

const SVGS = () => {
  const out = [];
  for (const theme of ["light", "dark"]) {
    out.push(["evals", evalsChartSvg(BASIC, COMPLEX, theme, MODEL)], ["benchmarks", benchmarksChartSvg(FORMATS, theme)], ["perf", perfChartSvg(FORMATS, theme)]);
  }
  return out;
};

describe("README chart embeds", () => {
  test("the data tables are still read from inside the collapsed blocks", () => {
    expect(BASIC.rows.map((row) => row.format)).toEqual(["docx", "xlsx", "pptx"]);
    expect(COMPLEX.rows.map((row) => row.format)).toEqual(["docx", "xlsx", "pptx"]);
    expect(FORMATS.map((f) => f.format)).toEqual(["docx", "pptx", "xlsx"]);
    for (const format of FORMATS) expect(format.rows.length).toBeGreaterThan(2);
    expect(agentEvalsSection(README)).toContain("<details>");
    expect(MODEL).toBe("gpt-6-luna");
  });

  test("each chart is a plain flush-left picture at full width, with a dark source and one-sentence alt", () => {
    for (const name of ["evals", "benchmarks", "perf"]) {
      expect(README).toContain(`<source media="(prefers-color-scheme: dark)" srcset="https://betteroffice.dev/api/charts/${name}.svg?theme=dark">`);
      const alt = altOf(name);
      expect(alt).toBeTruthy();
      expect(alt.split(/\.\s/)).toHaveLength(1);
      expect(alt).not.toMatch(/[–—]/);
    }
    const pictures = README.match(/<picture>[\s\S]*?<\/picture>/g);
    expect(pictures).toHaveLength(3);
    expect(README).not.toMatch(/<p align="center">\s*<picture>/);
    expect(README.match(/<details>\n<summary>Data<\/summary>/g)).toHaveLength(3);
    expect(README.match(/<\/details>/g)).toHaveLength(3);
  });

  test("alt text carries the numbers of the tables next to it", () => {
    expect(altOf("evals")).toBe(evalsHeadline(BASIC, COMPLEX, MODEL));
    expect(altOf("benchmarks")).toBe(benchmarksHeadline(FORMATS));
    expect(altOf("perf")).toBe(perfHeadline(FORMATS));
  });

  test("the perf data table lists every timing row the benchmark tables carry", () => {
    const from = README.indexOf("perf.svg");
    const block = README.slice(from, README.indexOf("</details>", from));
    const rows = [...block.matchAll(/^\| (DOCX|PPTX|XLSX) ([^|]+?) \| ([^|]+?) \| ([^|]+?) \|$/gm)].map((m) => [m[1].toLowerCase(), m[2], m[3], m[4]]);
    const timing = FORMATS.flatMap((f) => f.rows.filter((r) => /time/i.test(r.label)).map((r) => [f.format, r.label, r.ours, r.theirs]));
    expect(rows).toEqual(timing);
  });
});

describe("chart SVGs", () => {
  const evals = (theme) => evalsChartSvg(BASIC, COMPLEX, theme, MODEL);

  test("are well-formed, sized, transparent and free of invalid values", () => {
    for (const [, svg] of SVGS()) {
      expect(svg.startsWith("<svg xmlns=")).toBe(true);
      expect(svg.trimEnd().endsWith("</svg>")).toBe(true);
      expect(svg).toMatch(/viewBox="0 0 560 \d+"/);
      expect(svg).not.toMatch(/undefined|NaN|[–—]/);
      expect(svg).toMatch(/<title>[^<]+<\/title>/);
      expect(svg).not.toMatch(/<rect width="560"/);
    }
  });

  test("every text sits inside the canvas with padding, even in a wide fallback font", () => {
    for (const [name, svg] of SVGS()) {
      const width = Number(svg.match(/viewBox="0 0 (\d+) /)[1]);
      const height = Number(svg.match(/viewBox="0 0 \d+ (\d+)"/)[1]);
      const texts = [...svg.matchAll(/<text x="([\d.]+)" y="([\d.]+)" font-size="(\d+)"([^>]*)>([^<]*)<\/text>/g)];
      expect(texts.length).toBeGreaterThan(8);
      for (const [, x, y, size, attrs, content] of texts) {
        const wide = content.replace(/&\w+;/g, "x").length * Number(size) * 0.62;
        const anchor = attrs.match(/text-anchor="(\w+)"/)?.[1] ?? "start";
        const from = anchor === "middle" ? Number(x) - wide / 2 : anchor === "end" ? Number(x) - wide : Number(x);
        expect({ name, content, ok: from >= 12 && from + wide <= width - 12 && Number(y) <= height - 6 && Number(y) >= 10 }).toEqual({ name, content, ok: true });
      }
    }
  });

  test("the evals chart names the model and has a Total group next to every format", () => {
    const svg = evals("light");
    expect(svg).toContain("Agent edits with gpt-6-luna: pass rate by format");
    for (const label of ["DOCX", "XLSX", "PPTX", "Total"]) expect(svg).toContain(`>${label}<`);
    const total = (key) => BASIC.rows.reduce((t, r) => [t[0] + Number(r[key].split("/")[0]), t[1] + Number(r[key].split("/")[1])], [0, 0]);
    const [pass, runs] = total("ours");
    expect(svg).toContain(`>${Math.round((100 * pass) / runs)}%<`);
    for (const text of ["79%", "98%", "96%", "45%", "94%", "58%"]) expect(svg).toContain(`>${text}<`);
    expect(evalsChartSvg(BASIC, COMPLEX, "light")).toContain("Agent edits: pass rate by format");
  });

  test("basic bars are translucent with a full-colour outline, complex bars solid", () => {
    const svg = evals("light");
    expect(svg).toContain('fill-opacity="0.2"');
    expect(svg).toMatch(/fill="#047857" fill-opacity="1"/);
    expect(svg).toMatch(/fill-opacity="0.2" stroke="#047857"/);
  });

  test("every row of the README benchmark tables has a bar pair in the charts, with both values", () => {
    const bench = benchmarksChartSvg(FORMATS, "light");
    const perf = perfChartSvg(FORMATS, "light");
    let seen = 0;
    for (const format of FORMATS) {
      for (const row of format.rows) {
        const svg = /time/i.test(row.label) ? perf : bench;
        expect(svg).toContain(`data-row="${format.format}:${row.label}"`);
        const group = svg.slice(svg.indexOf(`data-row="${format.format}:${row.label}"`), svg.indexOf("</g>", svg.indexOf(`data-row="${format.format}:${row.label}"`)));
        expect(group).toContain(`>${row.ours}<`);
        expect(group).toContain(`>${row.theirs}<`);
        expect(group.match(/<rect /g)).toHaveLength(2);
        seen += 1;
      }
      expect(bench).toContain(`>${format.format.toUpperCase()}<`);
    }
    expect(seen).toBeGreaterThanOrEqual(15);
  });

  test("lower-is-better rows say so and size against the larger value", () => {
    const rows = tableRows(FORMATS, "other");
    const error = rows.find((row) => row.label === "Absolute page error");
    expect(error.better).toBe("lower");
    expect(error.ours.share).toBe(0);
    expect(error.theirs.share).toBe(1);
    expect(benchmarksChartSvg(FORMATS, "light")).toContain("Absolute page error (lower is better)");
    expect(perfChartSvg(FORMATS, "light")).toContain("Render time (avg) (lower is better)");
    const ssim = rows.find((row) => row.label === "SSIM");
    expect(ssim.ours.share).toBeCloseTo(0.8519, 4);
  });

  test("a value that is not a number is shown as text without a bar", () => {
    const formats = structuredClone(FORMATS);
    formats[0].rows[0] = { label: "SSIM", ours: "0.9000", theirs: "—" };
    const svg = benchmarksChartSvg(formats, "light");
    const start = svg.indexOf('data-row="docx:SSIM"');
    const group = svg.slice(start, svg.indexOf("</g>", start));
    expect(group.match(/<rect /g)).toHaveLength(1);
    expect(group).toContain(">—<");
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
    expect(perfChartSvg([], "light")).toBeNull();
  });
});

describe("chart routes", () => {
  process.env.BETTEROFFICE_AGENT_EVALS = SECTION;
  process.env.BETTEROFFICE_BENCHMARKS = README.slice(README.indexOf("<!-- BEGIN GENERATED VISUAL FIDELITY -->"), README.indexOf("<!-- END GENERATED VISUAL FIDELITY -->") + "<!-- END GENERATED VISUAL FIDELITY -->".length);

  test("serve cacheable SVG per theme", async () => {
    for (const name of ["evals", "benchmarks", "perf"]) {
      const { GET } = await import(`../app/api/charts/${name}.svg/route.ts`);
      const light = await GET(new Request(`https://betteroffice.dev/api/charts/${name}.svg`));
      expect(light.status).toBe(200);
      expect(light.headers.get("content-type")).toContain("image/svg+xml");
      expect(light.headers.get("cache-control")).toContain("s-maxage=3600");
      const dark = await GET(new Request(`https://betteroffice.dev/api/charts/${name}.svg?theme=dark`));
      expect(await dark.text()).not.toBe(await light.text());
    }
    const { GET } = await import("../app/api/charts/evals.svg/route.ts");
    expect(await (await GET(new Request("https://betteroffice.dev/api/charts/evals.svg"))).text()).toContain("gpt-6-luna");
  });

  test("answer 404 without results", async () => {
    const saved = process.env.BETTEROFFICE_AGENT_EVALS;
    process.env.BETTEROFFICE_AGENT_EVALS = "";
    const { GET } = await import("../app/api/charts/evals.svg/route.ts");
    expect((await GET(new Request("https://betteroffice.dev/api/charts/evals.svg"))).status).toBe(404);
    process.env.BETTEROFFICE_AGENT_EVALS = saved;
  });
});
