"use client";

import {
  defineXlsxPlugin,
  type XlsxPluginCommandResult,
  type XlsxPluginContext,
  type XlsxPluginGeometry,
} from "@betteroffice/xlsx-react";

const REVIEW_PLUGIN_ID = "demo.review";
const REVIEWED_FILL = "#d9ead3";
/** Larger selections are summarized without reading their cells. */
const MAX_READ_CELLS = 400;
const WRITE_NOT_GRANTED = {
  code: "write-not-granted",
  message: "Write access is not granted.",
};

interface ReviewSheet {
  sheetId: string;
  name: string;
  editable: boolean;
}

interface ReviewSelection {
  label: string;
  cells: number;
  filled: number | null;
}

interface ReviewState {
  version: string | null;
  sheets: ReviewSheet[];
  selection: ReviewSelection | null;
  message: string | null;
}

type ReviewContext = XlsxPluginContext<ReviewState>;

/** The selected cells as an inclusive range on the active sheet, or null. */
function selectedRange(context: ReviewContext) {
  const selection = context.snapshot.selection;
  const cells = selection?.cells;
  if (!selection || !cells) return null;
  return {
    sheetId: selection.sheetId,
    top: Math.min(cells.anchor.row, cells.focus.row),
    left: Math.min(cells.anchor.col, cells.focus.col),
    bottom: Math.max(cells.anchor.row, cells.focus.row),
    right: Math.max(cells.anchor.col, cells.focus.col),
  };
}

function target(range: NonNullable<ReturnType<typeof selectedRange>>) {
  return {
    sheetId: range.sheetId,
    range: {
      kind: "rowCol" as const,
      start: { row: range.top, col: range.left },
      end: { row: range.bottom, col: range.right },
    },
  };
}

async function refresh(context: ReviewContext): Promise<void> {
  const range = selectedRange(context);
  const count = range
    ? (range.bottom - range.top + 1) * (range.right - range.left + 1)
    : 0;
  const bounded = range !== null && count <= MAX_READ_CELLS;
  const read = await context.read.readCells({
    ranges: range && bounded ? [target(range)] : [],
  });
  if (!read.ok) return;
  const cells = read.ranges[0]?.cells.flat() ?? [];
  const first = cells[0]?.a1;
  const last = cells[cells.length - 1]?.a1;
  const selection: ReviewSelection | null = range
    ? {
        label:
          first && last && first !== last
            ? `${first}:${last}`
            : first ?? "selection",
        cells: count,
        filled: bounded
          ? cells.filter((cell) => cell.value.kind !== "empty").length
          : null,
      }
    : null;
  context.setState(
    (previous) => ({
      ...previous,
      version: read.version,
      sheets: read.sheets.map((sheet) => ({
        sheetId: sheet.sheetId,
        name: sheet.name,
        editable: sheet.editable,
      })),
      selection,
    }),
    read.version,
  );
}

async function markReviewed(
  context: ReviewContext,
): Promise<XlsxPluginCommandResult> {
  if (!context.edits) {
    context.setState((previous) => ({
      ...previous,
      message: WRITE_NOT_GRANTED.message,
    }));
    return { ok: false, failure: WRITE_NOT_GRANTED };
  }
  const range = selectedRange(context);
  const version = await context.read.version();
  if (!range || !version.ok) {
    const message = version.ok
      ? "Select cells first."
      : version.failure.message;
    context.setState((previous) => ({ ...previous, message }));
    return version.ok ? { ok: true, status: "noop" } : version;
  }
  const result = await context.edits.applyEdits({
    expectVersion: version.version,
    source: "agent",
    steps: [
      {
        op: "patchStyle",
        target: target(range),
        patch: { fillColor: REVIEWED_FILL },
      },
    ],
  });
  const message = result.ok
    ? "Marked the selection as reviewed."
    : `Refused (${result.failure.code}): ${result.failure.message}`;
  context.setState(
    (previous) => ({ ...previous, message }),
    "version" in result ? result.version : undefined,
  );
  return result.ok ? { ok: true, status: "executed" } : result;
}

function ReviewPanel({ context }: { context: ReviewContext }) {
  const { sheets, selection, message } = context.state;
  const writable = context.edits !== null && !context.snapshot.readOnly;
  const current = context.snapshot.selection?.sheetId;
  return (
    <div style={{ padding: 12, fontSize: 13, lineHeight: 1.45 }}>
      <p style={{ margin: "0 0 8px" }}>
        {sheets.length} sheets, version {context.snapshot.version.slice(-6)}
      </p>
      <ol style={{ margin: "0 0 12px", paddingLeft: 18 }}>
        {sheets.map((sheet) => (
          <li key={sheet.sheetId}>
            <button
              type="button"
              aria-current={sheet.sheetId === current ? "true" : undefined}
              style={{ textAlign: "left", textDecoration: "underline" }}
              onClick={() =>
                void context.run((action) =>
                  action.navigation
                    .selectCells(
                      {
                        sheetId: sheet.sheetId,
                        selection: {
                          anchor: { row: 0, col: 0 },
                          focus: { row: 0, col: 0 },
                        },
                      },
                      { expectVersion: action.snapshot.version },
                    )
                    .then(() => undefined),
                )
              }
            >
              {sheet.name}
              {sheet.editable ? "" : " (locked)"}
            </button>
          </li>
        ))}
      </ol>
      <p style={{ margin: "0 0 8px" }}>
        {selection
          ? `Selection ${selection.label}: ${selection.cells} cells${
              selection.filled === null ? "" : `, ${selection.filled} filled`
            }`
          : "No cells selected"}
      </p>
      <button
        type="button"
        disabled={!writable || !selection}
        title={
          writable ? undefined : "Needs write access and an editable editor"
        }
        onClick={() =>
          void context.run(async (action) => void (await markReviewed(action)))
        }
      >
        Mark selection reviewed
      </button>
      {message && (
        <p role="status" style={{ margin: "8px 0 0" }}>
          {message}
        </p>
      )}
    </div>
  );
}

function SelectionOverlay({
  context,
  geometry,
}: {
  context: ReviewContext;
  geometry: XlsxPluginGeometry;
}) {
  const range = selectedRange(context);
  const box = range
    ? geometry.getRangeRect({ sheetId: range.sheetId, range })
    : null;
  if (!box) return null;
  return (
    <div
      style={{
        position: "absolute",
        left: box.x - 3,
        top: box.y - 3,
        width: box.width + 6,
        height: box.height + 6,
        outline: "2px dashed rgba(37, 99, 235, 0.7)",
        borderRadius: 2,
      }}
    />
  );
}

export const reviewPlugin = defineXlsxPlugin<ReviewState>({
  id: REVIEW_PLUGIN_ID,
  createState: () => ({
    version: null,
    sheets: [],
    selection: null,
    message: null,
  }),
  initialize(context) {
    const visible = () => {
      if (document.visibilityState === "visible") void context.run(refresh);
    };
    document.addEventListener("visibilitychange", visible);
    context.onCleanup(() =>
      document.removeEventListener("visibilitychange", visible),
    );
  },
  async onEvent(context, event) {
    if (
      event.type === "load" ||
      event.type === "document-change" ||
      event.type === "selection-change"
    ) {
      await refresh(context);
    }
  },
  panel: {
    title: "Review",
    placement: "right",
    preferredSize: 260,
    render: ReviewPanel,
  },
  overlay: SelectionOverlay,
  commands: [
    {
      id: "mark-reviewed",
      label: "Mark reviewed",
      mutatesDocument: true,
      shortcuts: ["Mod+Alt+Shift+R"],
      getState: (context) =>
        context.edits
          ? { enabled: true }
          : { enabled: false, disabledReason: WRITE_NOT_GRANTED },
      execute: markReviewed,
    },
  ],
  toolbar: ["mark-reviewed"],
});
