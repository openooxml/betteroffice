export class VerticalCaretGoal {
  private goalX: number | undefined;

  current(): number | undefined {
    return this.goalX;
  }

  retain(goalX: number): void {
    this.goalX = goalX;
  }

  reset(): void {
    this.goalX = undefined;
  }
}

export interface VerticalParagraph {
  paraId: string;
  length: number;
}

export interface VerticalParagraphLoc {
  story: string;
  paraId: string;
  offset: number;
}

export function paragraphVerticalMove(
  paragraphs: readonly VerticalParagraph[],
  head: VerticalParagraphLoc,
  direction: 'up' | 'down'
): VerticalParagraphLoc {
  const index = paragraphs.findIndex((paragraph) => paragraph.paraId === head.paraId);
  if (index < 0) return head;
  const target = paragraphs[index + (direction === 'up' ? -1 : 1)];
  return target
    ? {
        story: head.story,
        paraId: target.paraId,
        offset: Math.min(head.offset, target.length),
      }
    : head;
}
