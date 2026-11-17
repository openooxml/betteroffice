type Handler = ((this: FakeWorkerScope, event: MessageEvent) => unknown) | null;

/** A dedicated worker scope: an EventTarget whose `onmessage` and `onmessageerror` behave like event handler attributes. */
export class FakeWorkerScope extends EventTarget {
  readonly posted: unknown[] = [];
  readonly listened: string[] = [];
  readonly assigned: string[] = [];
  readonly #handlers = new Map<string, { handler: Handler; listener: EventListener }>();

  postMessage(message: unknown, _transfer?: Transferable[]): void {
    this.posted.push(message);
  }

  override addEventListener(type: string, listener: EventListenerOrEventListenerObject | null, options?: AddEventListenerOptions | boolean): void {
    this.listened.push(type);
    super.addEventListener(type, listener, options);
  }

  get onmessage(): Handler { return this.#handlers.get('message')?.handler ?? null; }
  set onmessage(handler: Handler) { this.#assign('message', handler); }
  get onmessageerror(): Handler { return this.#handlers.get('messageerror')?.handler ?? null; }
  set onmessageerror(handler: Handler) { this.#assign('messageerror', handler); }

  #assign(type: string, handler: Handler): void {
    this.assigned.push(type);
    const slot = this.#handlers.get(type);
    if (slot && handler) {
      slot.handler = handler;
    } else if (slot) {
      super.removeEventListener(type, slot.listener);
      this.#handlers.delete(type);
    } else if (handler) {
      const entry = { handler, listener: (event: Event) => { entry.handler?.call(this, event as MessageEvent); } };
      this.#handlers.set(type, entry);
      super.addEventListener(type, entry.listener);
    }
  }
}
