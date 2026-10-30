import { readFileSync } from 'node:fs';

import {
  initWasm,
  openPresentation,
  type PptxEditRequest,
  type PptxExportOptions,
  type PptxFindRequest,
  type PptxReadRequest,
  type StorySnapshot,
  type TextParagraphDraft
} from '../../packages/pptx/src/index';
import { engineKind } from '../../packages/pptx/src/wasm/engine';
import type { Format } from './harness';

function storyResult({ id, length, paragraphs }: StorySnapshot) {
  return {
    id,
    length,
    text: paragraphs.map(({ runs }) => runs.map(({ text }) => text).join('')).join('\n'),
    paragraphs: paragraphs.map(({ id, alignment, level, runs }) => ({
      id,
      alignment,
      level,
      text: runs.map(({ text }) => text).join(''),
      runs: runs.map(({ text, style }) => ({
        text,
        style: {
          bold: style.bold,
          italic: style.italic,
          fontSizePt: style.fontSizePt,
          color: style.color,
          fontFamily: style.fontFamily,
          underline: style.underline,
          spacingPt: style.spacingPt ?? null,
          baselinePct: style.baselinePct ?? null,
          kernPt: style.kernPt ?? null,
          caps: style.caps ?? null
        }
      }))
    }))
  };
}

export const pptx: Format = {
  engineKind,
  versionPattern: /^[0-9a-f]{16}-[0-9]+$/,
  async open(bytes, script) {
    await initWasm(
      readFileSync(
        new URL('../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)
      )
    );
    const presentation = openPresentation(bytes, { clientId: script.clientId as number });
    return {
      operations: {
        version: () => presentation.version(),
        readContent: (args) => presentation.readContent(args as PptxReadRequest),
        findText: (args) => presentation.findText(args as unknown as PptxFindRequest),
        validateEdits: (args) => presentation.validateEdits(args as unknown as PptxEditRequest),
        applyEdits: (args) => presentation.applyEdits(args as unknown as PptxEditRequest),
        exportStructured: (args) => presentation.exportStructured(args as PptxExportOptions),
        exportMarkdown: (args) => presentation.exportMarkdown(args as PptxExportOptions),
        historyStep: (args) => {
          const result = presentation.historyStep(args.direction as 'undo' | 'redo');
          return { ...result, changedSlideIds: [...result.changedSlideIds].sort() };
        },
        duplicateSlide: (args) => presentation.duplicateSlide(
          args.slideId as string,
          args.toIndex as number
        ),
        bringShapeToFront: (args) => presentation.bringShapeToFront(
          args.slideId as string,
          args.shapeId as string
        ),
        addComment: (args) => presentation.addComment(
          args.slideId as string,
          args.comment as Parameters<typeof presentation.addComment>[1]
        ),
        setCommentPosition: (args) => presentation.setCommentPosition(
          args.commentId as string,
          args.position as { xEmu: number; yEmu: number }
        ),
        setStoryParagraphs: (args) => storyResult(
          presentation.setStoryParagraphs(
            args.storyId as string,
            args.paragraphs as TextParagraphDraft[]
          )
        ),
        readCommandContent: (args) => presentation.readCommandContent(args as PptxReadRequest),
        save: () => presentation.save()
      },
      dispose: () => presentation.dispose()
    };
  }
};
