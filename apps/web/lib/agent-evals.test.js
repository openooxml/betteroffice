import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { agentEvalCards, agentEvalTasks, agentEvalsSection, parseAgentEvals, publishedAgentEvals } from "./agent-evals.ts";

const README = readFileSync(join(import.meta.dir, "..", "..", "..", "README.md"), "utf8");
const table = (header, rows) => `| | ${header.join(" | ")} |\n|---|---|---|\n${rows.map((row) => `| ${row.join(" | ")} |`).join("\n")}`;
const block = (body) => `intro\n## Evals\n\n${body}\n\n## Later\n\n${table(["Generic file tools", "BetterOffice MCP"], [["PPTX", "1/2", "2/2"]])}`;
const one = (rows) => parseAgentEvals(block(table(["Generic file tools", "BetterOffice MCP"], rows)));

describe("agent eval summary", () => {
  test("reads the README's labels and ratios for every format", () => {
    const evals = parseAgentEvals(README);
    expect(evals.ours).toBe("BetterOffice MCP");
    expect(evals.theirs).toBe("Generic file tools");
    expect(evals.rows.map((row) => row.format)).toEqual(["docx", "xlsx", "pptx"]);
    for (const row of evals.rows) {
      expect(row.ours).toMatch(/^\d+\/\d+$/);
      expect(row.theirs).toMatch(/^\d+\/\d+$/);
    }
  });

  test("keeps values verbatim and finds our column on either side", () => {
    const rows = [{ format: "docx", ours: "10/10", theirs: "06/10" }];
    expect(one([["`DOCX`", "`06/10`", "**10/10**"]])).toEqual({ ours: "BetterOffice MCP", theirs: "Generic file tools", rows });
    expect(parseAgentEvals(block(table(["**BetterOffice MCP**", "`Generic file tools`"], [["DOCX", "**10/10**", "06/10"]])))).toEqual({
      ours: "BetterOffice MCP", theirs: "Generic file tools", rows,
    });
  });

  test("stops at the next section or marker and reads only the first table", () => {
    const first = table(["Generic file tools", "BetterOffice MCP"], [["DOCX", "6/10", "10/10"]]);
    expect(parseAgentEvals(block(first)).rows).toHaveLength(1);
    expect(parseAgentEvals(block(`no table\n<!-- marker -->\n${first}`))).toBeNull();
    expect(parseAgentEvals(block("no table"))).toBeNull();
    expect(parseAgentEvals(block(`${first}\n\n${first}`)).rows).toHaveLength(1);
  });

  test("yields nothing without the section, table or our column", () => {
    expect(parseAgentEvals("# readme\nno evals here")).toBeNull();
    expect(parseAgentEvals("intro ## Evals\nno table")).toBeNull();
    expect(parseAgentEvals(block(table(["Generic file tools", "Other tools"], [["DOCX", "1/2", "2/2"]])))).toBeNull();
    expect(parseAgentEvals(block("| | BetterOffice MCP |\n|---|---|\n| DOCX | 1/2 |"))).toBeNull();
    expect(parseAgentEvals(block("| | BetterOffice MCP | Generic file tools | Other tools |\n|---|---|---|---|\n| DOCX | 1/2 | 1/2 | 1/2 |"))).toBeNull();
  });

  test("reads the extracted section exactly as the whole README", () => {
    const section = agentEvalsSection(README);
    expect(section.startsWith("## Evals\n")).toBe(true);
    expect(section).not.toContain("<!--");
    expect(parseAgentEvals(section)).toEqual(parseAgentEvals(README));
  });

  test("reads the published build-time value", () => {
    const previous = process.env.BETTEROFFICE_AGENT_EVALS;
    try {
      process.env.BETTEROFFICE_AGENT_EVALS = README;
      expect(publishedAgentEvals()).toEqual(parseAgentEvals(README));
      delete process.env.BETTEROFFICE_AGENT_EVALS;
      expect(publishedAgentEvals()).toBeNull();
    } finally {
      if (previous === undefined) delete process.env.BETTEROFFICE_AGENT_EVALS;
      else process.env.BETTEROFFICE_AGENT_EVALS = previous;
    }
  });
});

describe("agent eval cards", () => {
  test("puts our toolset first with shares on a 0 to 1 scale", () => {
    const cards = agentEvalCards(one([["DOCX", "6/10", "10/10"], ["XLSX", "0/10", "10/10"]]));
    expect(cards).toEqual([
      { format: "docx", bars: [{ who: "BetterOffice MCP", ours: true, text: "10/10", share: 1 }, { who: "Generic file tools", ours: false, text: "6/10", share: 0.6 }] },
      { format: "xlsx", bars: [{ who: "BetterOffice MCP", ours: true, text: "10/10", share: 1 }, { who: "Generic file tools", ours: false, text: "0/10", share: 0 }] },
    ]);
  });

  test("draws only valid ratios in both columns", () => {
    const evals = one([["DOCX", "n/a", "10/10"], ["XLSX", "0/10", "100%"], ["PPTX", "8/0", "10/10"], ["VSDX", "11/10", "10/10"], ["DOCX", "6/10", "11/10"], ["XLSX", "0/10", "10/0"]]);
    expect(agentEvalCards(evals)).toEqual([]);
    expect(agentEvalTasks(evals)).toBe(0);
    expect(agentEvalCards(null)).toEqual([]);
    expect(agentEvalTasks(null)).toBe(0);
  });

  test("totals the tasks in the drawn README rows", () => {
    const evals = parseAgentEvals(README);
    expect(agentEvalCards(evals)).toHaveLength(3);
    expect(agentEvalTasks(evals)).toBe(30);
    expect(agentEvalTasks(one([["DOCX", "1/3", "2/4"], ["XLSX", "n/a", "10/10"]]))).toBe(4);
  });
});
