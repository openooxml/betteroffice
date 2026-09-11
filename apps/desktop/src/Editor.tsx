import { useCallback, useEffect, useRef, useState } from "react";
import { DocxEditor, type DocxEditorRef } from "@betteroffice/docx-react";
import { XlsxEditor, type XlsxEditorApi } from "@betteroffice/xlsx-react";
import { PptxEditor, type PptxEditorApi } from "@betteroffice/pptx-react";
import { configureDefaultFonts } from "@betteroffice/docx/layout";
import { setGoogleFontsEnabled } from "@betteroffice/docx/utils";
import { CollaborationProvider as DocxProvider } from "@betteroffice/docx/collaboration";
import { CollaborationProvider as XlsxProvider } from "@betteroffice/xlsx/collaboration";
import {
  CollaborationProvider as PptxProvider,
  type PptxFontFace,
} from "@betteroffice/pptx";
import type { Document } from "@betteroffice/docx/types/document";
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
} from "@betteroffice/fonts";
import { createRoomTransport } from "../../demo/app/collab/createRoomTransport";
import type { CollaborationReplica } from "../../demo/app/collab/types";
import type { Format } from "./bridge";
import "@betteroffice/docx-react/styles.css";

configureDefaultFonts({ load: () => import("@betteroffice/fonts") });
setGoogleFontsEnabled(false);

export interface PreparedFile {
  bytes: Uint8Array;
  format: Format;
  document?: Document;
  fonts?: PptxFontFace[];
}

export async function prepareFile(
  bytes: Uint8Array,
  format: Format
): Promise<PreparedFile> {
  if (format === "docx") {
    const { parseDocx } = await import("@betteroffice/docx/docx");
    return { bytes, format, document: await parseDocx(bytes) };
  }
  if (format === "xlsx") {
    const { initWasm, openWorkbook } = await import("@betteroffice/xlsx");
    await initWasm();
    openWorkbook(bytes).dispose();
    return { bytes, format };
  }
  const { initWasm, openPresentation } = await import("@betteroffice/pptx");
  await initWasm();
  const fonts = await Promise.all(
    [false, true].flatMap((bold) =>
      [false, true].map(async (italic) => {
        const face =
          resolveMetricCompatFace("Arial", bold, italic) ??
          resolveLastResortFace("Arial", bold, italic);
        return {
          family: "Arial",
          bold,
          italic,
          bytes: new Uint8Array(await loadBundledFontBytes(face)),
        };
      })
    )
  );
  openPresentation(bytes, { fonts }).dispose();
  return { bytes, format, fonts };
}

export interface EditorApi {
  serialize(): Promise<Uint8Array>;
  focus(): void;
  print(): void;
}
export interface Room {
  id: string;
  name: string;
}

export default function Editor({
  file,
  name,
  room,
  onSave,
  onChange,
  onError,
  onReady,
  onOpen,
  onStatus,
}: {
  file: PreparedFile;
  name: string;
  room: Room | null;
  onSave(bytes: Uint8Array): void;
  onChange(): void;
  onError(error: Error): void;
  onReady(api: EditorApi | null): void;
  onOpen(file: File): Promise<void>;
  onStatus(status: string): void;
}) {
  const serializing = useRef(false);
  const docx = useRef<DocxEditorRef>(null);
  const xlsx = useRef<XlsxEditorApi | null>(null);
  const pptx = useRef<PptxEditorApi | null>(null);
  const clientId = useRef(
    crypto.getRandomValues(new Uint32Array(1))[0] & 0x7fffffff || 1
  ).current;
  const [replica, setReplica] = useState<CollaborationReplica | null>(null);
  const [provider, setProvider] = useState<
    DocxProvider | XlsxProvider | PptxProvider | null
  >(null);
  const callbacks = useRef({ onSave, onChange, onError, onStatus });
  callbacks.current = { onSave, onChange, onError, onStatus };

  useEffect(() => {
    if (!room || !replica) return;
    const transport = createRoomTransport(
      "https://betteroffice-collaboration-relay.elia7.workers.dev",
      room.id
    );
    const options = { user: { name: room.name } };
    const next =
      file.format === "docx"
        ? new DocxProvider(replica, transport, options)
        : file.format === "xlsx"
        ? new XlsxProvider(replica, transport, options)
        : new PptxProvider(replica, transport, options);
    setProvider(next);
    const unsubscribe = next.onStatus(({ status, synced }) =>
      callbacks.current.onStatus(synced ? "Connected" : status)
    );
    const unsubscribeError = next.onError((error) =>
      callbacks.current.onError(error)
    );
    void next.connect();
    return () => {
      unsubscribe();
      unsubscribeError();
      next.destroy();
      setProvider(null);
    };
  }, [file.format, replica, room]);

  const changed = useCallback(() => callbacks.current.onChange(), []);
  const saveBytes = useCallback(
    (bytes: Uint8Array) => callbacks.current.onSave(bytes),
    []
  );
  const error = useCallback(
    (error: Error) => callbacks.current.onError(error),
    []
  );
  const workbookReady = useCallback(
    (api: XlsxEditorApi) => {
      xlsx.current = api;
      const unsubscribe = api.handle.onUpdate(changed);
      return () => {
        unsubscribe();
        xlsx.current = null;
      };
    },
    [changed]
  );
  const presentationReady = useCallback((api: PptxEditorApi) => {
    pptx.current = api;
  }, []);

  useEffect(() => {
    onReady({
      print() {
        if (file.format === "docx") docx.current?.print();
        else window.print();
      },
      focus() {
        if (file.format === "docx") docx.current?.focus();
        else if (file.format === "xlsx") xlsx.current?.focus();
        else pptx.current?.focus();
      },
      async serialize() {
        if (document.activeElement instanceof HTMLElement)
          document.activeElement.blur();
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
        serializing.current = true;
        try {
          const bytes =
            file.format === "docx"
              ? await docx.current?.save()
              : file.format === "xlsx"
              ? xlsx.current?.handle.save()
              : pptx.current?.save();
          if (!bytes)
            throw new Error(
              "The editor is still opening the file. Try again in a moment."
            );
          return new Uint8Array(bytes);
        } finally {
          serializing.current = false;
        }
      },
    });
    return () => onReady(null);
  }, [file.format, onReady]);

  if (file.format === "docx")
    return (
      <DocxEditor
        ref={docx}
        document={file.document}
        downloadOnSave={false}
        onSave={(buffer) => {
          if (!serializing.current) saveBytes(new Uint8Array(buffer));
        }}
        documentName={name}
        onChange={changed}
        onError={error}
        onOpen={onOpen}
        showToolbar
        showRuler
        showZoomControl
        colorMode="light"
        onPrint={() => undefined}
        collaboration={
          room
            ? {
                clientId,
                user: { name: room.name },
                onReplica: setReplica,
                presence:
                  provider instanceof DocxProvider ? provider : undefined,
              }
            : undefined
        }
      />
    );
  if (file.format === "xlsx")
    return (
      <XlsxEditor
        file={file.bytes}
        fileName={name}
        onSave={saveBytes}
        onReady={workbookReady}
        collaboration={
          room
            ? {
                clientId,
                onReplica: setReplica,
                provider:
                  provider instanceof XlsxProvider ? provider : undefined,
              }
            : undefined
        }
      />
    );
  return (
    <PptxEditor
      file={file.bytes}
      fileName={name}
      fonts={file.fonts ?? []}
      onSave={saveBytes}
      onChange={changed}
      onReady={presentationReady}
      onError={error}
      collaboration={
        room
          ? {
              clientId,
              onReplica: setReplica,
              presence: provider instanceof PptxProvider ? provider : undefined,
            }
          : undefined
      }
    />
  );
}
