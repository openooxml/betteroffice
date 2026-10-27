import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseAgentEvals, agentEvalsSection } from "./agent-evals.ts";
import { benchmarksChartSvg, benchmarksHeadline, evalsChartSvg, evalsHeadline, evalsModel, tableRows, themeOf } from "./charts.ts";
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
    out.push(["evals", evalsChartSvg(BASIC, COMPLEX, theme, MODEL)], ["benchmarks", benchmarksChartSvg(FORMATS, theme)]);
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
    for (const name of ["evals", "benchmarks"]) {
      expect(README).toContain(`<source media="(prefers-color-scheme: dark)" srcset="https://betteroffice.dev/api/charts/${name}.svg?theme=dark">`);
      const alt = altOf(name);
      expect(alt).toBeTruthy();
      expect(alt.split(/\.\s/)).toHaveLength(1);
      expect(alt).not.toMatch(/[–—]/);
    }
    const pictures = README.match(/<picture>[\s\S]*?<\/picture>/g);
    expect(pictures).toHaveLength(2);
    expect(README).not.toContain("perf.svg");
    expect(README).not.toMatch(/<p align="center">\s*<picture>/);
    expect(README.match(/<details>\n<summary>Data<\/summary>/g)).toHaveLength(2);
    expect(README.match(/<\/details>/g)).toHaveLength(2);
  });

  test("alt text carries the numbers of the tables next to it", () => {
    expect(altOf("evals")).toBe(evalsHeadline(BASIC, COMPLEX, MODEL));
    expect(altOf("benchmarks")).toBe(benchmarksHeadline(FORMATS));
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
        expect({ name, content, ok: from >= 0 && from + wide <= width && Number(y) <= height - 2 && Number(y) >= 10 }).toEqual({ name, content, ok: true });
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

  const TABLE_ONLY = /^(scored\/total|parse success)$/i;

  test("every quality and timing row of the README tables has a bar pair, with both values; only scored/total and parse success stay in the tables", () => {
    const bench = benchmarksChartSvg(FORMATS, "light");
    let seen = 0;
    for (const format of FORMATS) {
      for (const row of format.rows) {
        const marker = `data-row="${format.format}:${row.label}"`;
        if (TABLE_ONLY.test(row.label)) {
          expect(bench).not.toContain(marker);
          continue;
        }
        expect(bench).toContain(marker);
        const group = bench.slice(bench.indexOf(marker), bench.indexOf("</g>", bench.indexOf(marker)));
        expect(group).toContain(`>${row.ours}<`);
        expect(group).toContain(`>${row.theirs}<`);
        expect(group.match(/<rect /g)).toHaveLength(2);
        seen += 1;
      }
      expect(bench).toContain(`>${format.format.toUpperCase()}<`);
    }
    expect(seen).toBe(9);
  });

  test("within a format the quality rows come before the timing rows", () => {
    const bench = benchmarksChartSvg(FORMATS, "light");
    for (const format of FORMATS) {
      const at = (label) => bench.indexOf(`data-row="${format.format}:${label}"`);
      const time = format.rows.filter((row) => /time/i.test(row.label)).map((row) => at(row.label));
      const quality = format.rows.filter((row) => !/time/i.test(row.label) && !TABLE_ONLY.test(row.label)).map((row) => at(row.label));
      expect(Math.max(...quality)).toBeLessThan(Math.min(...time));
    }
  });

  test("lower-is-better rows are marked and sized against the larger value", () => {
    const rows = tableRows(FORMATS);
    const error = rows.find((row) => row.label === "Absolute page error");
    expect(error.better).toBe("lower");
    expect(error.ours.share).toBe(0);
    expect(error.theirs.share).toBe(1);
    const bench = benchmarksChartSvg(FORMATS, "light");
    for (const label of ["Absolute page error", "Render time", "Recalc time"]) {
      const at = bench.indexOf(`>${label}<`);
      expect(at).toBeGreaterThan(0);
      expect(bench.slice(at, at + 200)).toMatch(/<text x="0" y="[\d.]+" font-size="10" font-style="italic" fill="#6b6b6b">lower is better<\/text>/);
    }
    expect(bench.match(/>lower is better<\/text>/g)).toHaveLength(tableRows(FORMATS).filter((row) => row.better === "lower").length);
    expect(bench).not.toContain("(lower)");
    expect(bench).not.toContain("marks rows");
    expect(rows.find((row) => row.label === "SSIM").ours.share).toBeCloseTo(0.8519, 4);
    expect(rows.find((row) => row.label === "Render time (avg)").theirs.share).toBe(1);
  });

  test("content is flush with the canvas edges: titles, legend and labels start at x = 0, nothing is end-anchored in the benchmarks chart", () => {
    for (const [name, svg] of SVGS()) {
      expect(svg).toMatch(/<text x="0" y="1\d" font-size="1[56]" font-weight="600"/);
      const xs = [...svg.matchAll(/<text x="([\d.]+)"/g)].map((m) => Number(m[1]));
      expect(Math.min(...xs)).toBe(0);
      if (name === "benchmarks") expect(svg).not.toContain("text-anchor");
    }
  });

  test("the landing page font is embedded, with a system fallback, and each SVG stays small", () => {
    for (const [, svg] of SVGS()) {
      expect(svg).toContain("@font-face{font-family:Geist;");
      expect(svg).toMatch(/src:url\(data:font\/woff2;base64,[A-Za-z0-9+/=]{2000,}\)/);
      expect(svg).toContain("font-family=\"Geist, -apple-system");
      expect(svg.length).toBeLessThan(60_000);
    }
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
    expect(evals("light")).toContain("#94a3b8");
    expect(evals("dark")).toContain("#64748b");
    const tokens = readFileSync(join(import.meta.dir, "..", "..", "..", "shared", "site-tokens.css"), "utf8");
    expect(tokens).toContain("--bo-acc: #047857");
    expect(tokens).toContain("--bo-acc: #34d399");
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
  process.env.BETTEROFFICE_AGENT_EVALS = SECTION;
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
