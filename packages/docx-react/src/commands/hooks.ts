import { useCallback, useContext, useMemo, useRef, useSyncExternalStore } from 'react';
import { useTranslation } from '../i18n';
import { docxCommandController, type DocxChromeContext } from './createDocxCommandStore';
import { DocxCommandContext } from './DocxCommandProvider';
import { commandShortcut } from './descriptors';
import type {
  DocxCommandArgs,
  DocxCommandDescriptor,
  DocxCommandId,
  DocxCommandResult,
  DocxCommandState,
  DocxCommandStore,
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
): DocxCommandState<K> {
  const store = useDocxCommands();
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const subscribe = useCallback(
    (listener: () => void) => {
      const release = docxCommandController(store)?.hold(
        id,
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
    () => store.getState(id, argsKey === undefined ? undefined : JSON.parse(argsKey)),
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

/** Binds one command for a custom control. */
export function useDocxCommand<K extends DocxCommandId>(
  id: K,
  args?: DocxCommandArgs[K]
): DocxBoundCommand<K> {
  const store = useDocxCommands();
  const state = useDocxCommandState(id, args);
  const { t } = useTranslation();
  const descriptor = store.getDescriptor(id);
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const execute = useCallback(
    (callArgs?: DocxCommandArgs[K]) =>
      store.execute(
        id,
        (callArgs !== undefined
          ? callArgs
          : argsKey !== undefined
            ? JSON.parse(argsKey)
            : null) as DocxCommandArgs[K]
      ),
    [store, id, argsKey]
  );
  return useMemo(
    () => ({
      id,
      descriptor,
      state,
      label: t(descriptor.labelKey),
      shortcut: commandShortcut(id, argsKey === undefined ? undefined : JSON.parse(argsKey)),
      execute,
    }),
    [id, descriptor, state, t, argsKey, execute]
  );
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
