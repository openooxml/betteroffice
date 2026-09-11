export type Format = "docx" | "xlsx" | "pptx";
export interface OpenFile {
  id: string;
  name: string;
  url: string;
}
export interface RecentFile {
  id: string;
  name: string;
  folder: string;
}
export interface Configuration {
  format: Format;
  name: string;
  recent: RecentFile[];
}

export interface DesktopActions {
  open(file?: OpenFile): Promise<void>;
  close(): Promise<boolean>;
  command(command: string): void;
}

declare global {
  interface Window {
    webkit?: {
      messageHandlers: {
        desktop: { postMessage(message: unknown): Promise<unknown> };
      };
    };
    desktop?: DesktopActions;
  }
}

export function isNative(): boolean {
  return Boolean(window.webkit?.messageHandlers.desktop);
}

export async function invoke<T>(message: {
  type: string;
  [key: string]: unknown;
}): Promise<T> {
  const handler = window.webkit?.messageHandlers.desktop;
  if (!handler) throw new Error("This action is available in the desktop app.");
  return (await handler.postMessage(message)) as T;
}

export function formatOf(name: string): Format {
  const format = name.split(".").at(-1)?.toLowerCase();
  if (format !== "docx" && format !== "xlsx" && format !== "pptx") {
    throw new Error(
      "Choose a Word document, Excel spreadsheet, or PowerPoint presentation."
    );
  }
  return format;
}

export function encodeBytes(bytes: Uint8Array): string {
  let text = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    text += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(text);
}

export function dispatchEditorShortcut(key: string, shiftKey = false): void {
  (document.activeElement ?? document.body).dispatchEvent(
    new KeyboardEvent("keydown", {
      key,
      metaKey: true,
      shiftKey,
      bubbles: true,
      cancelable: true,
    })
  );
}
