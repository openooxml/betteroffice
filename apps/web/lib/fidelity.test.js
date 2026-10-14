import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fidelityCharts, parseFidelity, sharedVersion } from "./fidelity.ts";

const README = readFileSync(join(import.meta.dir, "..", "..", "..", "README.md"), "utf8");

const block = (body) =>
  `intro\n<!-- BEGIN GENERATED VISUAL FIDELITY -->\n## Benchmarks\n${body}\n<!-- END GENERATED VISUAL FIDELITY -->\nafter`;

const table = (rows) =>
  `<table>\n<tr><th width="180"></th><th width="230" align="right">BetterOffice (<a href="https://example.test/commit/abc">abc12345</a>)</th><th width="230" align="right">LibreOffice (1.2.3)</th></tr>\n${rows}\n</table>`;

describe("fidelity summary", () => {
  test("reads every format the README block publishes", () => {
    const formats = parseFidelity(README);
    expect(formats.map((f) => f.format).sort()).toEqual(["docx", "pptx", "xlsx"]);
    for (const f of formats) {
      expect(f.commit).toMatch(/^[0-9a-f]{8}$/);
      expect(f.commitUrl).toContain(f.commit);
      expect(f.libreoffice).not.toBe("");
      expect(f.rows.find((row) => row.label === "Scored/total")).toBeDefined();
    }
  });

  test("keeps published values verbatim", () => {
    const formats = parseFidelity(
      block(`### DOCX\n\n${table('<tr><td>SSIM</td><td align="right">0.8519</td><td align="right">0.7694</td></tr>\n<tr><td>Render time (avg)</td><td align="right">430 ms</td><td align="right">773 ms</td></tr>')}`)
    );
    expect(formats).toEqual([
      {
        format: "docx",
        commit: "abc12345",
        commitUrl: "https://example.test/commit/abc",
        libreoffice: "1.2.3",
        rows: [
          { label: "SSIM", ours: "0.8519", theirs: "0.7694" },
          { label: "Render time (avg)", ours: "430 ms", theirs: "773 ms" },
        ],
      },
    ]);
  });

  test("yields nothing without the generated block", () => {
    expect(parseFidelity("# readme\nno benchmarks here")).toEqual([]);
  });
});

const rows = (pairs) =>
  pairs.map(([label, ours, theirs]) => `<tr><td>${label}</td><td align="right">${ours}</td><td align="right">${theirs}</td></tr>`).join("\n");
const one = (format, pairs) => parseFidelity(block(`### ${format}\n\n${table(rows(pairs))}`))[0];
const bars = (chart) => chart.bars.map((bar) => [bar.who, bar.text, Number(bar.share.toFixed(4))]);

describe("fidelity charts", () => {
  test("draws the headline rows in a fixed order and leaves the rest to the table", () => {
    const f = one("DOCX", [
      ["Exact page counts", "63/63", "44/63"],
      ["Absolute page error", "0", "57"],
      ["SSIM", "0.8519", "0.7694"],
      ["Scored/total", "63/63", "63/63"],
      ["Render time (avg)", "430 ms", "773 ms"],
      ["Parse success", "100.00%", "100.00%"],
    ]);
    const charts = fidelityCharts(f);
    expect(charts.map((c) => [c.label, c.better, c.hint, c.group])).toEqual([
      ["Visual similarity", "higher", "SSIM against Word, 0 to 1", "quality"],
      ["Exact page counts", "higher", "documents with Word's page count", "quality"],
      ["Render time", "lower", "average per document", "time"],
    ]);
    expect(charts.map(bars)).toEqual([
      [["BetterOffice", "0.8519", 0.8519], ["LibreOffice", "0.7694", 0.7694]],
      [["BetterOffice", "63/63", 1], ["LibreOffice", "44/63", 0.6984]],
      [["BetterOffice", "430 ms", 0.5563], ["LibreOffice", "773 ms", 1]],
    ]);
  });

  test("puts percentages on a 0 to 1 scale and times on the slower engine", () => {
    const charts = fidelityCharts(
      one("XLSX", [
        ["SSIM", "0.7661", "0.6892"],
        ["Recalc accuracy", "99.60%", "85.87%"],
        ["Recalc time (avg)", "64 ms", "405 ms"],
      ])
    );
    expect(charts.map((c) => [c.label, c.hint, c.group])).toEqual([
      ["Visual similarity", "SSIM against Excel, 0 to 1", "quality"],
      ["Formula accuracy", "results matching Excel after recalculation", "quality"],
      ["Recalculation time", "average per workbook", "time"],
    ]);
    expect(bars(charts[1])).toEqual([["BetterOffice", "99.60%", 0.996], ["LibreOffice", "85.87%", 0.8587]]);
    expect(bars(charts[2])).toEqual([["BetterOffice", "64 ms", 0.158], ["LibreOffice", "405 ms", 1]]);
  });

  test("draws no bars for a value that is not a number", () => {
    const f = one("PPTX", [
      ["SSIM", "n/a", "0.6"],
      ["Render time (avg)", "—", "1351 ms"],
      ["Scored/total", "1/2", "2/2"],
    ]);
    expect(fidelityCharts(f)).toEqual([]);
  });

  test("charts every published format", () => {
    for (const f of parseFidelity(README)) {
      const charts = fidelityCharts(f);
      expect(charts.length).toBeGreaterThan(0);
      expect(charts.map((c) => c.group)).toEqual(charts.map((c) => c.group).sort());
      for (const chart of charts) {
        const row = f.rows.find((r) => r.label === chart.row);
        expect(chart.bars.map((bar) => bar.text)).toEqual([row.ours, row.theirs]);
        for (const bar of chart.bars) expect(bar.share).toBeWithin(0, 1.0000001);
      }
    }
  });

  test("shares one version line only when every format ran on the same builds", () => {
    const docx = one("DOCX", [["SSIM", "0.8", "0.7"]]);
    const xlsx = one("XLSX", [["SSIM", "0.8", "0.7"]]);
    expect(sharedVersion([docx, xlsx])).toEqual({ commit: "abc12345", commitUrl: "https://example.test/commit/abc", libreoffice: "1.2.3" });
    expect(sharedVersion([docx, { ...xlsx, libreoffice: "9.9" }])).toBeNull();
    expect(sharedVersion([])).toBeNull();
  });
});
