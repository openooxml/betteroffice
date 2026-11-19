import type { DisplayPage } from '../layout/render/displayList';
import { displayPageRevision, retainedFramePageById, type RetainedFrame } from '../layout/render/frameDelta';
import type {
  DocxCommandEchoContext,
  DocxCommandPageIdentity,
  DocxCommandStagedPresentation,
  DocxEchoRect,
} from './residentCommandProtocol';

interface StagedEchoPage {
  page: DocxCommandPageIdentity;
  frameEpoch: number;
  displayPage: DisplayPage;
  revision: number;
  layoutRevision: number;
  pixelRatio: number;
  backgrounds: readonly { id: string; rect: DocxEchoRect; omitted: string }[];
}

/** The page bitmaps and echo backgrounds main holds, as this worker staged them, by page id. */
export class StagedEchoPages {
  private readonly pages = new Map<string, StagedEchoPage>();

  /** Records what `presentation` gives main; a full presentation replaces everything recorded before. */
  record(
    presentation: DocxCommandStagedPresentation, frame: RetainedFrame, layoutRevision: number, pixelRatio: number,
    echo: DocxCommandEchoContext | undefined, full: boolean
  ): void {
    if (full) this.pages.clear();
    for (const staged of presentation.pages) {
      const retained = retainedFramePageById(frame, BigInt(staged.page.pageId));
      if (!retained) {
        this.pages.delete(staged.page.pageId);
        continue;
      }
      const sent = new Set((staged.backgrounds ?? []).map(({ id }) => id));
      this.pages.set(staged.page.pageId, {
        page: staged.page, frameEpoch: staged.frameEpoch, displayPage: retained.page,
        revision: displayPageRevision(retained.page), layoutRevision, pixelRatio,
        backgrounds: (echo?.damage ?? []).filter((damage) => damage.page.pageId === staged.page.pageId && sent.has(damage.backgroundId))
          .map((damage) => ({ id: damage.backgroundId, rect: damage.rect, omitted: damage.omittedPrimitiveIds.join(',') })),
      });
    }
  }

  /**
   * `context` drawn over the page bitmap and background main already holds, when its page and
   * line are unchanged since they were staged; undefined when main holds none that fits.
   */
  reuse(
    context: DocxCommandEchoContext, frame: RetainedFrame, layoutRevision: number, pixelRatio: number
  ): DocxCommandEchoContext | undefined {
    const held = this.pages.get(context.page.pageId);
    const [damage] = context.damage;
    if (!held || !damage || context.damage.length !== 1 || damage.page.pageId !== context.page.pageId ||
      held.page.pageIndex !== context.page.pageIndex || held.layoutRevision !== layoutRevision ||
      held.pixelRatio !== pixelRatio) return undefined;
    const retained = retainedFramePageById(frame, BigInt(context.page.pageId));
    if (!retained || retained.page !== held.displayPage || displayPageRevision(retained.page) !== held.revision) return undefined;
    const omitted = damage.omittedPrimitiveIds.join(',');
    const background = held.backgrounds.find((candidate) => candidate.omitted === omitted &&
      candidate.rect.x === damage.rect.x && candidate.rect.y === damage.rect.y &&
      candidate.rect.width === damage.rect.width && candidate.rect.height === damage.rect.height);
    return background
      ? { ...context, frameEpoch: held.frameEpoch, damage: [{ ...damage, backgroundId: background.id }] }
      : undefined;
  }
}
