/**
 * Two-mode caret arbitration (the Google Docs model): while typing, the worker
 * paints a solid caret line into the same presented frame as the glyphs and
 * the DOM blink caret is hidden; once input pauses or the selection moves, the
 * DOM caret takes over and the painted line is erased.
 */

/** Idle threshold: painted-caret mode ends this long after the last text input. */
export const CARET_PAINT_IDLE_MS = 500;

export const CARET_PAINT_FALLBACK_COLOR = '#000';

/** Resolve the DOM caret's `--doc-caret` token for the worker painter. */
export function resolveCaretPaintColor(element: Element | null): string {
  if (!element || typeof getComputedStyle !== 'function') return CARET_PAINT_FALLBACK_COLOR;
  const value = getComputedStyle(element).getPropertyValue('--doc-caret').trim();
  return value || CARET_PAINT_FALLBACK_COLOR;
}

export class PaintedCaretMachine {
  private active = false;
  private lastInputAt = Number.NEGATIVE_INFINITY;
  private hideUntil = Number.NEGATIVE_INFINITY;
  private interrupts = 0;

  /** True while the painted line owns the caret (DOM caret hidden). */
  isActive(): boolean {
    return this.active;
  }

  /** Local text input: opens (or extends) the paint window. */
  noteInput(now: number): void {
    this.lastInputAt = now;
  }

  /** Text input dispatched toward a painted frame: hide the DOM caret NOW.
   * The stale caret would otherwise sit one keystroke behind the atomically
   * presented frame until the reply commits; a briefly hidden caret is
   * invisible, a stale one is not. */
  noteDispatch(now: number): void {
    this.lastInputAt = now;
    this.hideUntil = now + CARET_PAINT_IDLE_MS;
  }

  /** A dispatch is still awaiting its painted frame. */
  isHolding(now: number): boolean {
    return now < this.hideUntil;
  }

  /** A frame built now should carry the painted caret. */
  shouldPaint(now: number): boolean {
    return now - this.lastInputAt <= CARET_PAINT_IDLE_MS;
  }

  /** Captured before a paint request; its reply activates only if unchanged. */
  token(): number {
    return this.interrupts;
  }

  /** A presented frame carried the painted line. Returns whether it activates
   * (false means the paint is stale and the caller must erase it). */
  framePainted(token: number): boolean {
    if (token !== this.interrupts) return false;
    this.hideUntil = Number.NEGATIVE_INFINITY;
    this.active = true;
    return true;
  }

  /** A presented frame carried no painted line (the worker already erased any
   * previous one). Restores the DOM caret: drops any dispatch hold and returns
   * whether active mode ended. */
  frameUnpainted(): boolean {
    this.hideUntil = Number.NEGATIVE_INFINITY;
    const wasActive = this.active;
    this.active = false;
    return wasActive;
  }

  /** Selection move, blur, IME start, or mode change: closes the paint window,
   * drops any dispatch hold, and invalidates in-flight paints. Returns whether
   * a painted line must be erased (DOM caret first, then erase). */
  interrupt(): boolean {
    this.interrupts += 1;
    this.lastInputAt = Number.NEGATIVE_INFINITY;
    this.hideUntil = Number.NEGATIVE_INFINITY;
    const wasActive = this.active;
    this.active = false;
    return wasActive;
  }

  /** Time until the idle threshold elapses. */
  msUntilIdle(now: number): number {
    return Math.max(0, this.lastInputAt + CARET_PAINT_IDLE_MS - now);
  }

  /** Fires the idle transition once the threshold has elapsed. Returns whether
   * it deactivated (DOM caret first, then erase). */
  idleTimeout(now: number): boolean {
    if (!this.active || now - this.lastInputAt < CARET_PAINT_IDLE_MS) return false;
    this.active = false;
    return true;
  }
}
