"use client";

import {
  defineDocxPlugin,
  type DocxCommandResult,
  type DocxPluginContext,
  type DocxPluginGeometry,
  type DocxPluginSidebarItem,
} from "@betteroffice/docx-react";

const REVIEW_PLUGIN_ID = "demo.review";

interface ReviewParagraph {
  paraId: string;
  text: string;
}

interface ReviewState {
  version: string | null;
  paragraphs: ReviewParagraph[];
  message: string | null;
}

type ReviewContext = DocxPluginContext<ReviewState>;

async function refresh(context: ReviewContext): Promise<void> {
  const read = await context.read.readParagraphs({ view: "accepted" });
  if (!read.ok) return;
  const paragraphs = read.paragraphs
    .filter((paragraph) => paragraph.text.trim().length > 0)
    .slice(0, 5)
    .map(({ paraId, text }) => ({ paraId, text }));
  context.setState(
    (previous) => ({ ...previous, version: read.version, paragraphs }),
    read.version,
  );
}

async function markReviewed(
  context: ReviewContext,
): Promise<DocxCommandResult> {
  const read = await context.read.readParagraphs({ view: "accepted" });
  const target = read.ok
    ? read.paragraphs.find((paragraph) => paragraph.text.trim().length > 0)
    : undefined;
  if (!read.ok || !target || !context.edits) {
    const message = read.ok
      ? "Write access is not granted."
      : read.failure.message;
    context.setState((previous) => ({ ...previous, message }));
    return { ok: true, status: "noop" };
  }
  const result = await context.edits.applyEdits({
    expectVersion: read.version,
    source: "agent",
    steps: [
      {
        op: "insertText",
        target: { kind: "paragraph", story: "body", paraId: target.paraId },
        at: "end",
        text: " ✓",
      },
    ],
  });
  const message = result.ok
    ? "Marked the first paragraph as reviewed."
    : `Refused (${result.failure.code}): ${result.failure.message}`;
  context.setState(
    (previous) => ({ ...previous, message }),
    "version" in result ? result.version : undefined,
  );
  return result.ok
    ? { ok: true, status: "executed" }
    : { ok: true, status: "noop" };
}

function ReviewPanel({ context }: { context: ReviewContext }) {
  const { paragraphs, message } = context.state;
  const writable = context.edits !== null && !context.snapshot.readOnly;
  return (
    <div style={{ padding: 12, fontSize: 13, lineHeight: 1.45 }}>
      <p style={{ margin: "0 0 8px" }}>
        {paragraphs.length} paragraphs, version{" "}
        {context.snapshot.version.slice(-6)}
      </p>
      <ol style={{ margin: "0 0 12px", paddingLeft: 18 }}>
        {paragraphs.map((paragraph) => (
          <li key={paragraph.paraId}>
            <button
              type="button"
              style={{ textAlign: "left", textDecoration: "underline" }}
              onClick={() =>
                void context.run((action) =>
                  action.navigation
                    .scrollToParagraph(
                      { story: "body", paraId: paragraph.paraId },
                      { expectVersion: action.snapshot.version },
                    )
                    .then(() => undefined),
                )
              }
            >
              {paragraph.text.slice(0, 60)}
            </button>
          </li>
        ))}
      </ol>
      <button
        type="button"
        disabled={!writable}
        title={writable ? undefined : "Needs write access and an editable mode"}
        onClick={() =>
          void context.run(async (action) => void (await markReviewed(action)))
        }
      >
        Mark first paragraph reviewed
      </button>
      {message && (
        <p role="status" style={{ margin: "8px 0 0" }}>
          {message}
        </p>
      )}
    </div>
  );
}

function ReviewCard({
  context,
  item,
}: {
  context: ReviewContext;
  item: DocxPluginSidebarItem<ReviewState>;
}) {
  const paragraph = context.state.paragraphs.find(
    (candidate) => candidate.paraId === item.anchor.paraId,
  );
  return (
    <div
      style={{
        padding: 8,
        borderRadius: 6,
        border: "1px solid #c9d7f2",
        background: "#f4f8ff",
        fontSize: 12,
      }}
    >
      Review starts at &ldquo;{paragraph?.text.slice(0, 40)}&rdquo;
    </div>
  );
}

function SelectionOverlay({
  context,
  geometry,
}: {
  context: ReviewContext;
  geometry: DocxPluginGeometry;
}) {
  const range = context.snapshot.selection.displayRange;
  if (
    !range ||
    range.from === range.to ||
    range.layoutId !== geometry.layout.id
  ) {
    return null;
  }
  return (
    <>
      {geometry.dom
        .getRectsForRange(range.from, range.to)
        .map((rect, index) => {
          const box = geometry.toOverlayRect(rect);
          return (
            <div
              key={index}
              style={{
                position: "absolute",
                left: box.x,
                top: box.y,
                width: box.width,
                height: box.height,
                outline: "2px dashed rgba(37, 99, 235, 0.7)",
              }}
            />
          );
        })}
    </>
  );
}

export const reviewPlugin = defineDocxPlugin<ReviewState>({
  id: REVIEW_PLUGIN_ID,
  createState: () => ({ version: null, paragraphs: [], message: null }),
  initialize(context) {
    const timer = window.setInterval(() => void context.run(refresh), 30_000);
    context.onCleanup(() => window.clearInterval(timer));
  },
  async onEvent(context, event) {
    if (event.type === "load" || event.type === "document-change")
      await refresh(context);
  },
  panel: {
    title: "Review",
    placement: "right",
    preferredSize: 260,
    render: ReviewPanel,
  },
  overlay: SelectionOverlay,
  getSidebarItems(context) {
    const first = context.state.paragraphs[0];
    const version = context.state.version;
    return first && version
      ? [
          {
            id: "first-paragraph",
            anchor: { version, story: "body", paraId: first.paraId },
            render: ReviewCard,
          },
        ]
      : [];
  },
  commands: [
    {
      id: "mark-reviewed",
      label: "Mark reviewed",
      mutatesDocument: true,
      shortcuts: ["Mod+Alt+Shift+R"],
      getState: (context) =>
        context.edits
          ? { enabled: true }
          : {
              enabled: false,
              disabledReason: {
                code: "write-not-granted",
                message: "Write access is not granted.",
              },
            },
      execute: markReviewed,
    },
  ],
  toolbar: ["mark-reviewed"],
});
