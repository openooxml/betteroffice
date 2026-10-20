import { expect, test } from "bun:test";
import type { XlsxExportCell, XlsxExportSheet, XlsxStructuredContent } from "@betteroffice/xlsx";
import { buildTotalsEdits } from "./demoAgent";

function cell(a1: string, value: XlsxExportCell["value"]): XlsxExportCell {
  return {
    id: a1,
    anchor: { kind: "cell", sheet: { index: 0, name: "Sheet1" }, a1 } as XlsxExportCell["anchor"],
    value,
    formula: null,
    displayText: "",
    numberFormat: "General",
    formulaResult: null,
    merge: null,
  };
}

function content(cells: XlsxExportCell[], truncated = false): XlsxStructuredContent {
  const sheet = {
    kind: "worksheet",
    anchor: { kind: "sheet", sheet: { index: 0, name: "Sheet1" } },
    truncated,
    cells,
  } as unknown as XlsxExportSheet;
  return { sheets: [sheet], truncated } as unknown as XlsxStructuredContent;
}

const number = (value: number) => ({ kind: "number" as const, value });
const text = (value: string) => ({ kind: "text" as const, value });

test("totals sum numeric cells and leave numeric-looking text alone", () => {
  const edits = buildTotalsEdits(content([
    cell("A1", number(1)), cell("A2", number(2)),
    cell("B1", text("3")), cell("B2", text("4")),
  ]));
  expect(edits).toEqual([{ sheet: 0, row: 2, col: 0, input: "=SUM(A1:A2)" }]);
});

test("totals never replace an occupied cell", () => {
  expect(buildTotalsEdits(content([
    cell("A1", number(1)), cell("A2", number(2)), cell("A3", text("note")),
  ]))).toEqual([]);
});

test("totals never replace error or unavailable cells", () => {
  for (const value of [
    { kind: "error" as const, value: "#DIV/0!" },
    { kind: "unavailable" as const, reason: "limit" },
  ]) {
    expect(buildTotalsEdits(content([
      cell("A1", number(1)), cell("A2", number(2)), cell("A3", value as XlsxExportCell["value"]),
    ]))).toEqual([]);
  }
});

test("a truncated export proposes nothing for that sheet", () => {
  expect(buildTotalsEdits(content([cell("A1", number(1)), cell("A2", number(2))], true))).toEqual([]);
});
