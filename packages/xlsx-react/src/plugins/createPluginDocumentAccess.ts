import type { WorkbookHandle, XlsxEditRequest, XlsxEditResult, XlsxEditRefusal } from '@betteroffice/xlsx';
import type { XlsxCommandController } from '../commands/createXlsxCommandStore';
import type { XlsxAdmission, XlsxPluginEditorAccess, XlsxPluginNavigator } from './createPluginClients';
import type { XlsxPluginReadClient, XlsxPluginSelection } from './types';

type ReadMethod<K extends keyof XlsxPluginReadClient> = (
  identity: object, ...args: Parameters<XlsxPluginReadClient[K]>
) => ReturnType<XlsxPluginReadClient[K]> | Awaited<ReturnType<XlsxPluginReadClient[K]>>;

type NavigationArguments<K extends keyof XlsxPluginNavigator> =
  Parameters<XlsxPluginNavigator[K]> extends [WorkbookHandle, ...infer A] ? [...A, string] : never;

export interface XlsxPluginDocumentAccess {
  identity(): object | null;
  version(identity: object): string | null;
  currentVersion(identity: object): string | null;
  subscribe(identity: object, listener: () => void): () => void;
  selection?(identity: object): XlsxPluginSelection;
  readOnly?(identity: object): boolean;
  admit<T>(identity: object, operation: () => T, fence?: boolean): Promise<XlsxAdmission<T>>;
  readVersion: ReadMethod<'version'>;
  readCells: ReadMethod<'readCells'>;
  findText: ReadMethod<'findText'>;
  validateEdits: ReadMethod<'validateEdits'>;
  readOnlyRefusal(identity: object): XlsxEditRefusal | null;
  applyEdits<Refusal>(identity: object, request: XlsxEditRequest,
    authorize: () => Refusal | null, commit: <T>(write: () => T) => T): XlsxEditResult | Refusal;
  admitEdits?<Refusal>(identity: object, request: XlsxEditRequest,
    authorize: (recovering: boolean) => Refusal | null,
    commit: <T>(write: () => T) => T,
    signals?: readonly AbortSignal[]): Promise<XlsxAdmission<XlsxEditResult | Refusal>>;
  navigator: {
    [K in keyof XlsxPluginNavigator]: (identity: object,
      ...args: NavigationArguments<K>
    ) => ReturnType<XlsxPluginNavigator[K]> | Promise<ReturnType<XlsxPluginNavigator[K]>>;
  };
  commands(): XlsxCommandController | null;
}

export function createWorkbookPluginDocument(access: XlsxPluginEditorAccess): XlsxPluginDocumentAccess {
  const handle = (identity: object) => identity as WorkbookHandle;
  const version = (identity: object): string | null => {
    try { return handle(identity).version(); } catch { return null; }
  };
  return {
    identity: () => access.handle(),
    version, currentVersion: version,
    subscribe: (identity, listener) => handle(identity).onUpdate(listener),
    admit: (identity, operation) => access.admit(handle(identity), operation),
    readVersion: (identity) => ({ ok: true, version: handle(identity).version() }),
    readCells: (identity, request) => handle(identity).readCells(request),
    findText: (identity, request) => handle(identity).findText(request),
    validateEdits: (identity, request) => access.readOnlyRefusal(handle(identity)) ?? handle(identity).validateEdits(request),
    readOnlyRefusal: (identity) => access.readOnlyRefusal(handle(identity)),
    applyEdits: (identity, ...args) => access.applyEdits(handle(identity), ...args),
    ...(access.admitEdits ? { admitEdits: <Refusal>(identity: object, request: XlsxEditRequest,
      authorize: (recovering: boolean) => Refusal | null, commit: <T>(write: () => T) => T) =>
      access.admitEdits!(handle(identity), request, authorize, commit) } : {}),
    navigator: {
      selectCells: (identity, target, focus, live) => access.navigator.selectCells(handle(identity), target, focus, live),
      scrollToCell: (identity, target, align, live) => access.navigator.scrollToCell(handle(identity), target, align, live),
    },
    commands: access.commands,
  };
}
