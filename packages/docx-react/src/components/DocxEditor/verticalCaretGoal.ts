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
