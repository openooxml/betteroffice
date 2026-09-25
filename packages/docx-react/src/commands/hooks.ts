import { useCallback, useContext, useMemo, useRef, useSyncExternalStore } from 'react';
import { useTranslation } from '../i18n';
import { docxCommandController, type DocxChromeContext } from './createDocxCommandStore';
import { DocxCommandContext } from './DocxCommandProvider';
import { commandShortcut, formatChord, isPluginCommandId } from './descriptors';
import type {
  DocxCommandArgs,
  DocxCommandDescriptor,
  DocxCommandId,
  DocxCommandResult,
  DocxCommandState,
  DocxCommandStore,
  DocxPluginCommandDescriptor,
  DocxPluginCommandId,
  DocxPluginCommandState,
} from './types';

/** The nearest editor's command store. */
export function useDocxCommands(): DocxCommandStore {
  const store = useContext(DocxCommandContext);
  if (!store) {
    throw new Error('useDocxCommands must be used inside DocxEditor chrome or a DocxCommandProvider');
  }
  return store;
}

/** Subscribes to one command's state, optionally for specific arguments. */
export function useDocxCommandState<K extends DocxCommandId>(
  id: K,
  args?: DocxCommandArgs[K]
): DocxCommandState<K>;
/** Subscribes to a contributed command's state. */
export function useDocxCommandState(id: DocxPluginCommandId, args?: null): DocxPluginCommandState;
export function useDocxCommandState(
  id: DocxCommandId | DocxPluginCommandId,
  args?: unknown
): DocxCommandState | DocxPluginCommandState {
  const store = useDocxCommands();
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const subscribe = useCallback(
    (listener: () => void) => {
      const release = docxCommandController(store)?.hold(
        id as DocxCommandId,
        argsKey === undefined ? undefined : JSON.parse(argsKey)
      );
      const unsubscribe = store.subscribe(listener);
      return () => {
        unsubscribe();
        release?.();
      };
    },
    [store, id, argsKey]
  );
  const getSnapshot = useCallback(
    () =>
      store.getState(id as DocxCommandId, argsKey === undefined ? undefined : JSON.parse(argsKey)),
    [store, id, argsKey]
  );
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}

/** A command bound to this component, with its localized label and shortcut. */
export interface DocxBoundCommand<K extends DocxCommandId> {
  id: K;
  descriptor: DocxCommandDescriptor<K>;
  state: DocxCommandState<K>;
  label: string;
  shortcut: string | null;
  /** Executes with `args`, falling back to the arguments bound to the hook. */
  execute(args?: DocxCommandArgs[K]): Promise<DocxCommandResult>;
}

/** A contributed command bound to this component. */
export interface DocxBoundPluginCommand {
  id: DocxPluginCommandId;
  /** Null while no active plugin contributes the command. */
  descriptor: DocxPluginCommandDescriptor | null;
  state: DocxPluginCommandState;
  /** The contributed label, or the id while no plugin contributes it. */
  label: string;
  shortcut: string | null;
  execute(): Promise<DocxCommandResult>;
}

/** Binds one command for a custom control. */
export function useDocxCommand<K extends DocxCommandId>(
  id: K,
  args?: DocxCommandArgs[K]
): DocxBoundCommand<K>;
/** Binds a contributed command for a custom control. */
export function useDocxCommand(id: DocxPluginCommandId, args?: null): DocxBoundPluginCommand;
export function useDocxCommand(
  id: DocxCommandId | DocxPluginCommandId,
  args?: unknown
): DocxBoundCommand<DocxCommandId> | DocxBoundPluginCommand {
  const store = useDocxCommands();
  const state = useDocxCommandState(id as DocxCommandId, args as never);
  const { t } = useTranslation();
  const descriptor = store.getDescriptor(id as DocxCommandId) as
    | DocxCommandDescriptor
    | DocxPluginCommandDescriptor
    | null;
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const execute = useCallback(
    (callArgs?: unknown) =>
      store.execute(
        id as DocxCommandId,
        (callArgs !== undefined
          ? callArgs
          : argsKey !== undefined
            ? JSON.parse(argsKey)
            : null) as never
      ),
    [store, id, argsKey]
  );
  return useMemo(() => {
    if (isPluginCommandId(id)) {
      const contributed = descriptor as DocxPluginCommandDescriptor | null;
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
    const builtIn = descriptor as DocxCommandDescriptor;
    return {
      id,
      descriptor: builtIn,
      state: state as DocxCommandState,
      label: t(builtIn.labelKey),
      shortcut: commandShortcut(id, argsKey === undefined ? undefined : JSON.parse(argsKey)),
      execute,
    };
  }, [id, descriptor, state, t, argsKey, execute]);
}

/** The editor's locale, color mode and theme, following their changes. */
export function useDocxChrome(): DocxChromeContext | null {
  const controller = docxCommandController(useDocxCommands());
  const subscribe = useCallback(
    (listener: () => void) => controller?.subscribeChrome(listener) ?? (() => {}),
    [controller]
  );
  const getSnapshot = useCallback(() => controller?.chrome() ?? null, [controller]);
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}

/**
 * A ref for chrome rendered outside the editor's tree, such as a portalled
 * popup, so shortcuts pressed inside it reach the surrounding editor.
 */
export function useCommandChromeRef(): (element: HTMLElement | null) => void {
  const store = useContext(DocxCommandContext);
  const release = useRef<(() => void) | null>(null);
  return useCallback(
    (element: HTMLElement | null) => {
      release.current?.();
      release.current = null;
      const controller = store ? docxCommandController(store) : null;
      if (element && controller) release.current = controller.registerChrome(element);
    },
    [store]
  );
}
