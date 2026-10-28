/** One view as reported to the host. */
export interface WopiView {
  ViewId: number;
  UserName?: string;
  ReadOnly?: boolean;
}

/** Host actions the page can be asked to perform. */
export interface PostMessageHandlers {
  onSave(): void;
  onClose(): void;
  onBlurFocus(): void;
  onGrabFocus(): void;
  onGetViews(): WopiView[];
  onHostReady?(): void;
}

/** Minimal window surface the bridge needs, so tests can fake it. */
export interface PostMessageWindow {
  parent: { postMessage(message: string, targetOrigin: string): void };
  addEventListener(type: "message", listener: (event: MessageEvent) => void): void;
  removeEventListener(type: "message", listener: (event: MessageEvent) => void): void;
}

/** Normalizes a value to an http(s) origin, or empty when it is not one. */
export function normalizeHttpOrigin(raw: string): string {
  if (raw === "" || raw === "*") return "";
  try {
    const url = new URL(raw);
    if (url.protocol !== "http:" && url.protocol !== "https:") return "";
    return url.origin;
  } catch {
    return "";
  }
}

const INCOMING = new Set([
  "Host_PostmessageReady",
  "Action_Save",
  "Action_Close",
  "Blur_Focus",
  "Grab_Focus",
  "Get_Views",
]);

/** WOPI PostMessage bridge. Sends only to the configured origin, never to *. */
export class PostMessageBridge {
  readonly #window: PostMessageWindow;
  readonly #origin: string;
  readonly #handlers: PostMessageHandlers;
  readonly #now: () => number;
  #listening = false;
  #hostReady = false;

  readonly #listener = (event: MessageEvent): void => {
    this.handleMessage(event);
  };

  constructor(win: PostMessageWindow, origin: string, handlers: PostMessageHandlers, now: () => number = () => Date.now()) {
    this.#window = win;
    this.#origin = normalizeHttpOrigin(origin);
    this.#handlers = handlers;
    this.#now = now;
  }

  get enabled(): boolean {
    return this.#origin !== "";
  }

  get hostReady(): boolean {
    return this.#hostReady;
  }

  /** Starts accepting host messages. */
  start(): void {
    if (this.#listening) return;
    this.#listening = true;
    this.#window.addEventListener("message", this.#listener);
  }

  /** Stops accepting host messages. */
  dispose(): void {
    if (!this.#listening) return;
    this.#listening = false;
    this.#window.removeEventListener("message", this.#listener);
  }

  /** Sends one message as a JSON string to the host origin only. */
  send(messageId: string, values: Record<string, unknown> = {}): void {
    if (!this.enabled) return;
    const payload = JSON.stringify({ MessageId: messageId, SendTime: this.#now(), Values: values });
    this.#window.parent.postMessage(payload, this.#origin);
  }

  frameReady(): void {
    this.send("App_LoadingStatus", { Status: "Frame_Ready" });
  }

  bootFailed(): void {
    this.send("App_LoadingStatus", { Status: "Failed" });
  }

  documentLoaded(documentLoadedTime: number = this.#now()): void {
    this.send("App_LoadingStatus", { Status: "Document_Loaded", DocumentLoadedTime: documentLoadedTime });
  }

  modified(isModified: boolean): void {
    this.send("Doc_ModifiedStatus", { Modified: isModified });
  }

  uiClose(): void {
    this.send("UI_Close", {});
  }

  viewAdded(view: WopiView): void {
    this.send("View_Added", { ...view });
  }

  viewRemoved(view: WopiView): void {
    this.send("View_Removed", { ...view });
  }

  /** Validates and dispatches one incoming message event. */
  handleMessage(event: { origin?: unknown; data?: unknown }): void {
    if (!this.enabled) return;
    if (event.origin !== this.#origin) return;
    let data: unknown = event.data;
    if (typeof data === "string") {
      try {
        data = JSON.parse(data);
      } catch {
        return;
      }
    }
    if (typeof data !== "object" || data === null) return;
    const message = data as { MessageId?: unknown; SendTime?: unknown; Values?: unknown };
    if (typeof message.MessageId !== "string" || !INCOMING.has(message.MessageId)) return;
    if (message.SendTime !== undefined && (typeof message.SendTime !== "number" || !Number.isFinite(message.SendTime))) return;
    if (message.Values !== undefined && (typeof message.Values !== "object" || message.Values === null || Array.isArray(message.Values))) return;
    switch (message.MessageId) {
      case "Host_PostmessageReady":
        this.#hostReady = true;
        this.#handlers.onHostReady?.();
        return;
      case "Action_Save":
        this.#handlers.onSave();
        return;
      case "Action_Close":
        this.#handlers.onClose();
        return;
      case "Blur_Focus":
        this.#handlers.onBlurFocus();
        return;
      case "Grab_Focus":
        this.#handlers.onGrabFocus();
        return;
      case "Get_Views":
        this.send("Get_Views_Resp", { Views: this.#handlers.onGetViews() });
        return;
    }
  }
}
