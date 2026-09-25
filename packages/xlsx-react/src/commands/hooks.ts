import { useCallback, useContext, useMemo, useSyncExternalStore } from 'react';
import { useTranslation } from '../i18n';
import { xlsxCommandController, type XlsxChromeContext } from './createXlsxCommandStore';
import { commandLabelKey, commandShortcut } from './descriptors';
import { XlsxCommandContext } from './XlsxCommandProvider';
import type {
  XlsxCommandArgs,
  XlsxCommandDescriptor,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandState,
  XlsxCommandStore,
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
): XlsxCommandState<K> {
  const store = useXlsxCommands();
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const subscribe = useCallback(
    (listener: () => void) => {
      const release = xlsxCommandController(store)?.hold(id, parse(argsKey));
      const unsubscribe = store.subscribe(listener);
      return () => {
        unsubscribe();
        release?.();
      };
    },
    [store, id, argsKey]
  );
  const getSnapshot = useCallback(
    () => store.getState(id, parse<XlsxCommandArgs[K]>(argsKey)),
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

/** Binds one command for a custom control. */
export function useXlsxCommand<K extends XlsxCommandId>(
  id: K,
  args?: XlsxCommandArgs[K]
): XlsxBoundCommand<K> {
  const store = useXlsxCommands();
  const state = useXlsxCommandState(id, args);
  const { t } = useTranslation();
  const descriptor = store.getDescriptor(id);
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const execute = useCallback(
    (callArgs?: XlsxCommandArgs[K]) =>
      store.execute(
        id,
        (callArgs !== undefined ? callArgs : (parse(argsKey) ?? null)) as XlsxCommandArgs[K]
      ),
    [store, id, argsKey]
  );
  return useMemo(
    () => ({
      id,
      descriptor,
      state,
      label: t(commandLabelKey(id, parse<XlsxCommandArgs[K]>(argsKey))),
      shortcut: commandShortcut(id, parse<XlsxCommandArgs[K]>(argsKey)),
      execute,
    }),
    [id, descriptor, state, t, argsKey, execute]
  );
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
