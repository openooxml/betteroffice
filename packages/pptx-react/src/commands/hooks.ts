import { useCallback, useContext, useMemo, useRef, useSyncExternalStore } from 'react';
import { useTranslation } from '../i18n';
import { pptxCommandController, type PptxChromeContext } from './createPptxCommandStore';
import { PptxCommandContext } from './PptxCommandProvider';
import { commandShortcut, defaultArgs } from './descriptors';
import type {
  PptxCommandArgs,
  PptxCommandDescriptor,
  PptxCommandId,
  PptxCommandResult,
  PptxCommandState,
  PptxCommandStore,
} from './types';

/**
 * The nearest editor's command store.
 * @experimental
 */
export function usePptxCommands(): PptxCommandStore {
  const store = useContext(PptxCommandContext);
  if (!store) {
    throw new Error(
      'usePptxCommands must be used inside PptxEditor chrome or a PptxCommandProvider'
    );
  }
  return store;
}

function parse(argsKey: string | undefined) {
  return argsKey === undefined ? undefined : JSON.parse(argsKey);
}

/**
 * Subscribes to one command's state, optionally for specific arguments.
 * @experimental
 */
export function usePptxCommandState<K extends PptxCommandId>(
  id: K,
  args?: PptxCommandArgs[K]
): PptxCommandState<K> {
  const store = usePptxCommands();
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const subscribe = useCallback(
    (listener: () => void) => {
      const release = pptxCommandController(store)?.hold(id, parse(argsKey));
      const unsubscribe = store.subscribe(listener);
      return () => {
        unsubscribe();
        release?.();
      };
    },
    [store, id, argsKey]
  );
  const getSnapshot = useCallback(() => store.getState(id, parse(argsKey)), [store, id, argsKey]);
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}

/**
 * A command bound to this component, with its localized label and shortcut.
 * @experimental
 */
export interface PptxBoundCommand<K extends PptxCommandId> {
  id: K;
  descriptor: PptxCommandDescriptor<K>;
  state: PptxCommandState<K>;
  label: string;
  /** The keyboard shortcut as people read it on this platform, such as `Ctrl+B`. */
  shortcut: string | null;
  /** Executes with `args`, falling back to the arguments bound to the hook. */
  execute(args?: PptxCommandArgs[K]): Promise<PptxCommandResult>;
}

/**
 * Binds one command for a custom control.
 * @experimental
 */
export function usePptxCommand<K extends PptxCommandId>(
  id: K,
  args?: PptxCommandArgs[K]
): PptxBoundCommand<K> {
  const store = usePptxCommands();
  const state = usePptxCommandState(id, args);
  const { t } = useTranslation();
  const descriptor = store.getDescriptor(id);
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const execute = useCallback(
    (callArgs?: PptxCommandArgs[K]) =>
      store.execute(
        id,
        callArgs !== undefined
          ? callArgs
          : argsKey !== undefined
          ? JSON.parse(argsKey)
          : defaultArgs(id)
      ),
    [store, id, argsKey]
  );
  return useMemo(
    () => ({
      id,
      descriptor,
      state,
      label: t(descriptor.labelKey),
      shortcut: commandShortcut(id, parse(argsKey)),
      execute,
    }),
    [id, descriptor, state, t, argsKey, execute]
  );
}

/** The editor's locale, following its changes; `null` outside an editor. */
export function usePptxChrome(): PptxChromeContext | null {
  const controller = pptxCommandController(usePptxCommands());
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
  const store = useContext(PptxCommandContext);
  const release = useRef<(() => void) | null>(null);
  return useCallback(
    (element: HTMLElement | null) => {
      release.current?.();
      release.current = null;
      const controller = store ? pptxCommandController(store) : null;
      if (element && controller) release.current = controller.registerChrome(element);
    },
    [store]
  );
}
