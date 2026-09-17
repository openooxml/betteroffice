import type { VsdxEditorApi } from "@betteroffice/vsdx-react";

export type VsdxApiWindow = { __vsdxApi?: unknown };

export function shouldExposeVsdxApi(search: string): boolean {
  return new URLSearchParams(search).get("vsdxApi") === "1";
}

export function exposeVsdxApi(target: object, api: VsdxEditorApi): void {
  (target as VsdxApiWindow).__vsdxApi = api;
}

export function clearVsdxApi(target: object, api: VsdxEditorApi): void {
  const view = target as VsdxApiWindow;
  if (view.__vsdxApi === api) delete view.__vsdxApi;
}
