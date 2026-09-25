import { useCallback, useContext, useMemo, useSyncExternalStore } from 'react';
import { useTranslation } from '../i18n';
import { xlsxCommandController, type XlsxChromeContext } from './createXlsxCommandStore';
import { commandLabelKey, commandShortcut, formatChord, isPluginCommandId } from './descriptors';
import { XlsxCommandContext } from './XlsxCommandProvider';
import type {
  XlsxCommandArgs,
  XlsxCommandDescriptor,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandState,
  XlsxCommandStore,
  XlsxPluginCommandDescriptor,
  XlsxPluginCommandId,
  XlsxPluginCommandState,
} from './types';

/** The nearest editor's command store. */
export function useXlsxCommands(): XlsxCommandStore {
  const store = useContext(XlsxCommandContext);
  if (!store) {
    throw new Error('useXlsxCommands must be used inside XlsxEditor chrome or an XlsxCommandProvider');
  }
  return store;
}

function parse<T>(key: string | undefined): T | undefined {
  return key === undefined ? undefined : (JSON.parse(key) as T);
}

/** Subscribes to one command's state, optionally for specific arguments. */
export function useXlsxCommandState<K extends XlsxCommandId>(
  id: K,
  args?: XlsxCommandArgs[K]
): XlsxCommandState<K>;
/** Subscribes to a contributed command's state. */
export function useXlsxCommandState(id: XlsxPluginCommandId, args?: null): XlsxPluginCommandState;
export function useXlsxCommandState(
  id: XlsxCommandId | XlsxPluginCommandId,
  args?: unknown
): XlsxCommandState | XlsxPluginCommandState {
  const store = useXlsxCommands();
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const subscribe = useCallback(
    (listener: () => void) => {
      const release = xlsxCommandController(store)?.hold(id as XlsxCommandId, parse(argsKey));
      const unsubscribe = store.subscribe(listener);
      return () => {
        unsubscribe();
        release?.();
      };
    },
    [store, id, argsKey]
  );
  const getSnapshot = useCallback(
    () => store.getState(id as XlsxCommandId, parse(argsKey)),
    [store, id, argsKey]
  );
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}

/** A command bound to this component, with its localized label and shortcut. */
export interface XlsxBoundCommand<K extends XlsxCommandId> {
  id: K;
  descriptor: XlsxCommandDescriptor<K>;
  state: XlsxCommandState<K>;
  /** The label for the bound arguments, such as "Format as currency". */
  label: string;
  shortcut: string | null;
  /** Executes with `args`, falling back to the arguments bound to the hook. */
  execute(args?: XlsxCommandArgs[K]): Promise<XlsxCommandResult>;
}

/** A contributed command bound to this component. */
export interface XlsxBoundPluginCommand {
  id: XlsxPluginCommandId;
  /** Null while no active plugin contributes the command. */
  descriptor: XlsxPluginCommandDescriptor | null;
  state: XlsxPluginCommandState;
  /** The contributed label, or the id while no plugin contributes it. */
  label: string;
  shortcut: string | null;
  execute(): Promise<XlsxCommandResult>;
}

/** Binds one command for a custom control. */
export function useXlsxCommand<K extends XlsxCommandId>(
  id: K,
  args?: XlsxCommandArgs[K]
): XlsxBoundCommand<K>;
/** Binds a contributed command for a custom control. */
export function useXlsxCommand(id: XlsxPluginCommandId, args?: null): XlsxBoundPluginCommand;
export function useXlsxCommand(
  id: XlsxCommandId | XlsxPluginCommandId,
  args?: unknown
): XlsxBoundCommand<XlsxCommandId> | XlsxBoundPluginCommand {
  const store = useXlsxCommands();
  const state = useXlsxCommandState(id as XlsxCommandId, args as never);
  const { t } = useTranslation();
  const descriptor = store.getDescriptor(id as XlsxCommandId) as
    | XlsxCommandDescriptor
    | XlsxPluginCommandDescriptor
    | null;
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const execute = useCallback(
    (callArgs?: unknown) =>
      store.execute(
        id as XlsxCommandId,
        (callArgs !== undefined ? callArgs : (parse(argsKey) ?? null)) as never
      ),
    [store, id, argsKey]
  );
  return useMemo(() => {
    if (isPluginCommandId(id)) {
      const contributed = descriptor as XlsxPluginCommandDescriptor | null;
      const chord = contributed?.shortcuts[0]?.chord;
      return {
        id,
        descriptor: contributed,
        state,
        label: contributed?.label ?? id,
        shortcut: chord ? formatChord(chord) : null,
        execute: () => execute(),
      };
    }
    return {
      id,
      descriptor: descriptor as XlsxCommandDescriptor,
      state: state as XlsxCommandState,
      label: t(commandLabelKey(id, parse(argsKey))),
      shortcut: commandShortcut(id, parse(argsKey)),
      execute,
    };
  }, [id, descriptor, state, t, argsKey, execute]);
}

/** The editor's locale, following its changes; `null` without an editor. */
export function useXlsxChrome(): XlsxChromeContext | null {
  const controller = xlsxCommandController(useXlsxCommands());
  const subscribe = useCallback(
    (listener: () => void) => controller?.subscribeChrome(listener) ?? (() => {}),
    [controller]
  );
  const getSnapshot = useCallback(() => controller?.chrome() ?? null, [controller]);
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}
