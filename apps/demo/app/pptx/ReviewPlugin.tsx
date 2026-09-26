"use client";

import {
  definePptxPlugin,
  type PptxCommandResult,
  type PptxPluginContext,
  type PptxPluginGeometry,
} from "@betteroffice/pptx-react";

const REVIEW_PLUGIN_ID = "demo.review";
const REVIEWED = "Reviewed ✓";

interface ReviewSlide {
  slideId: string;
  title: string;
  notes: string;
}

interface ReviewState {
  version: string | null;
  slides: ReviewSlide[];
  message: string | null;
}

type ReviewContext = PptxPluginContext<ReviewState>;

async function refresh(context: ReviewContext): Promise<void> {
  const read = await context.read.readContent();
  if (!read.ok) return;
  const slides = read.slides.map((slide) => {
    const story = read.stories.find(
      (candidate) =>
        candidate.slideId === slide.id && candidate.text.trim().length > 0,
    );
    return {
      slideId: slide.id,
      title: story?.text.split("\n")[0].slice(0, 60) ?? "Untitled slide",
      notes: slide.notes ?? "",
    };
  });
  context.setState(
    (previous) => ({ ...previous, version: read.version, slides }),
    read.version,
  );
}

async function markReviewed(
  context: ReviewContext,
): Promise<PptxCommandResult> {
  const read = await context.read.readContent();
  const slideId =
    context.snapshot.selection?.slideId ??
    (read.ok ? read.slides[0]?.id : undefined);
  const slide = read.ok
    ? read.slides.find((candidate) => candidate.id === slideId)
    : undefined;
  if (!read.ok || !slide || !context.edits) {
    const message = read.ok
      ? "Write access is not granted."
      : read.failure.message;
    context.setState((previous) => ({ ...previous, message }));
    return { ok: true, status: "noop" };
  }
  const notes = slide.notes ?? "";
  const result = await context.edits.applyEdits({
    expectVersion: read.version,
    source: "agent",
    steps: [
      {
        op: "setSlideNotes",
        target: { slideId: slide.id },
        text: notes ? `${notes}\n${REVIEWED}` : REVIEWED,
      },
    ],
  });
  const message = result.ok
    ? "Marked the current slide as reviewed."
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
  const { slides, message } = context.state;
  const writable = context.edits !== null && !context.snapshot.readOnly;
  const current = context.snapshot.selection?.slideId;
  return (
    <div style={{ padding: 12, fontSize: 13, lineHeight: 1.45 }}>
      <p style={{ margin: "0 0 8px" }}>
        {slides.length} slides, version {context.snapshot.version.slice(-6)}
      </p>
      <ol style={{ margin: "0 0 12px", paddingLeft: 18 }}>
        {slides.map((slide) => (
          <li key={slide.slideId}>
            <button
              type="button"
              aria-current={slide.slideId === current ? "true" : undefined}
              style={{ textAlign: "left", textDecoration: "underline" }}
              onClick={() =>
                void context.run((action) =>
                  action.navigation
                    .goToSlide(
                      { slideId: slide.slideId },
                      { expectVersion: action.snapshot.version },
                    )
                    .then(() => undefined),
                )
              }
            >
              {slide.title}
              {slide.notes.endsWith(REVIEWED) ? " ✓" : ""}
            </button>
          </li>
        ))}
      </ol>
      <button
        type="button"
        disabled={!writable}
        title={
          writable ? undefined : "Needs write access and an editable editor"
        }
        onClick={() =>
          void context.run(async (action) => void (await markReviewed(action)))
        }
      >
        Mark current slide reviewed
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
  geometry: PptxPluginGeometry;
}) {
  const target = context.snapshot.selection?.target;
  const shapeId =
    target && target.kind !== "slide" ? target.shapeId : undefined;
  const box = shapeId ? geometry.getShapeRect(shapeId) : null;
  if (!box) return null;
  return (
    <div
      style={{
        position: "absolute",
        left: box.x - 4,
        top: box.y - 4,
        width: box.width + 8,
        height: box.height + 8,
        outline: "2px dashed rgba(37, 99, 235, 0.7)",
        borderRadius: 4,
      }}
    />
  );
}

export const reviewPlugin = definePptxPlugin<ReviewState>({
  id: REVIEW_PLUGIN_ID,
  createState: () => ({ version: null, slides: [], message: null }),
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
