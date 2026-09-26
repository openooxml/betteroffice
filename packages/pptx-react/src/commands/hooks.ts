import { useCallback, useContext, useMemo, useRef, useSyncExternalStore } from 'react';
import { useTranslation } from '../i18n';
import { pptxCommandController, type PptxChromeContext } from './createPptxCommandStore';
import { PptxCommandContext } from './PptxCommandProvider';
import { commandShortcut, defaultArgs, formatChord, isPluginCommandId } from './descriptors';
import type {
  PptxCommandArgs,
  PptxCommandDescriptor,
  PptxCommandId,
  PptxCommandResult,
  PptxCommandState,
  PptxCommandStore,
  PptxPluginCommandDescriptor,
  PptxPluginCommandId,
  PptxPluginCommandResult,
  PptxPluginCommandState,
} from './types';

/** The nearest editor's command store. */
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

/** Subscribes to one command's state, optionally for specific arguments. */
export function usePptxCommandState<K extends PptxCommandId>(
  id: K,
  args?: PptxCommandArgs[K]
): PptxCommandState<K>;
/** Subscribes to a contributed command's state. */
export function usePptxCommandState(id: PptxPluginCommandId, args?: null): PptxPluginCommandState;
export function usePptxCommandState(
  id: PptxCommandId | PptxPluginCommandId,
  args?: unknown
): PptxCommandState | PptxPluginCommandState {
  const store = usePptxCommands();
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const subscribe = useCallback(
    (listener: () => void) => {
      const release = pptxCommandController(store)?.hold(id as PptxCommandId, parse(argsKey));
      const unsubscribe = store.subscribe(listener);
      return () => {
        unsubscribe();
        release?.();
      };
    },
    [store, id, argsKey]
  );
  const getSnapshot = useCallback(
    () => store.getState(id as PptxCommandId, parse(argsKey)),
    [store, id, argsKey]
  );
  return useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
}

/** A command bound to this component, with its localized label and shortcut. */
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

/** A contributed command bound to this component. */
export interface PptxBoundPluginCommand {
  id: PptxPluginCommandId;
  /** Null while no active plugin contributes the command. */
  descriptor: PptxPluginCommandDescriptor | null;
  state: PptxPluginCommandState;
  /** The contributed label, or the id while no plugin contributes it. */
  label: string;
  shortcut: string | null;
  execute(): Promise<PptxPluginCommandResult>;
}

/** Binds one command for a custom control. */
export function usePptxCommand<K extends PptxCommandId>(
  id: K,
  args?: PptxCommandArgs[K]
): PptxBoundCommand<K>;
/** Binds a contributed command for a custom control. */
export function usePptxCommand(id: PptxPluginCommandId, args?: null): PptxBoundPluginCommand;
export function usePptxCommand(
  id: PptxCommandId | PptxPluginCommandId,
  args?: unknown
): PptxBoundCommand<PptxCommandId> | PptxBoundPluginCommand {
  const store = usePptxCommands();
  const state = usePptxCommandState(id as PptxCommandId, args as never);
  const { t } = useTranslation();
  const descriptor = store.getDescriptor(id as PptxCommandId) as
    | PptxCommandDescriptor
    | PptxPluginCommandDescriptor
    | null;
  const argsKey = args === undefined ? undefined : JSON.stringify(args);
  const execute = useCallback(
    (callArgs?: unknown) =>
      store.execute(
        id as PptxCommandId,
        (callArgs !== undefined
          ? callArgs
          : argsKey !== undefined
          ? JSON.parse(argsKey)
          : defaultArgs(id as PptxCommandId)) as never
      ),
    [store, id, argsKey]
  );
  return useMemo(() => {
    if (isPluginCommandId(id)) {
      const contributed = descriptor as PptxPluginCommandDescriptor | null;
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
    const builtIn = descriptor as PptxCommandDescriptor;
    return {
      id,
      descriptor: builtIn,
      state: state as PptxCommandState,
      label: t(builtIn.labelKey),
      shortcut: commandShortcut(id, parse(argsKey)),
      execute,
    };
  }, [id, descriptor, state, t, argsKey, execute]);
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
