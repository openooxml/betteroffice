import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { resolveNavigationTarget, type YrsSession } from '@betteroffice/docx/yrs';
import { grantsCommand, grantsEditBatch, grantsWrite } from '../../../../shared/plugin-host/grants';
import type { InvocationRefusal, PluginInvocation } from '../../../../shared/plugin-host/runtime';
import {
  UNAVAILABLE_DOCX_COMMANDS,
  type DocxCommandController,
  type DocxCommandScope,
} from '../commands/createDocxCommandStore';
import { DOCX_COMMAND_DESCRIPTORS, isPluginCommandId } from '../commands/descriptors';
import type { DocxCommandStore } from '../commands/types';
import {
  applyEditBatch,
  flushEditorInput,
  modeRefusal,
} from '../components/DocxEditor/editorBatches';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';
import { isLayoutQueued, sourceVersionOf } from '../components/DocxEditor/internals/layoutProvenance';
import {
  requestWorkerOpenReplica,
  workerOpenSourceVersion,
} from '../components/DocxEditor/internals/workerOpenReplica';
import {
  handedOverRequest,
  workerProposalAuthority,
  type WorkerProposalAuthority,
} from '../components/DocxEditor/internals/workerProposalAuthority';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import { currentPreviewKey, renderedPreviewKey } from './proposalPreview';
import type {
  DocxPluginCommandClient,
  DocxPluginEditClient,
  DocxPluginFailureCode,
  DocxPluginGrant,
  DocxPluginNavigation,
  DocxPluginNavigationFailureCode,
  DocxPluginReadClient,
  DocxPluginRefusal,
  DocxPluginSnapshot,
} from './types';

/** The editor as plugin clients reach it; every member is read at call time. */
export interface DocxPluginEditorAccess {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  /** The editor mode writes are checked against; `viewing` also while `readOnly`. */
  writeMode(): EditorMode;
  commands(): DocxCommandController | null;
  layout(): { queries: DisplayListQueries | null; complete: boolean; failed: boolean };
  subscribeLayout(listener: () => void): () => void;
}

const LAYOUT_WAIT_MS = 30_000;
const navigationRequests = new WeakMap<object, AbortController>();

const MESSAGES: Record<DocxPluginFailureCode, string> = {
  'plugin-unavailable': 'The plugin is no longer active',
  'document-replaced': 'The document was replaced',
  aborted: 'This plugin context was superseded',
  'permission-denied': 'The host has not granted this operation',
  'read-only': 'The editor is read-only',
  'input-failed': 'Pending input could not be applied',
  'unsupported-policy': 'Plugins cannot run this operation yet',
  'plugin-failed': 'The plugin failed',
};

export function pluginRefusal(
  code: DocxPluginFailureCode,
  message = MESSAGES[code]
): DocxPluginRefusal {
  return { ok: false, failure: { code, message } };
}

function refusalOf(invocation: PluginInvocation<DocxPluginSnapshot>): DocxPluginRefusal | null {
  const refusal: InvocationRefusal | null = invocation.refusal();
  return refusal ? pluginRefusal(refusal) : null;
}

/** The command scope of one invocation: its activation and the plugin's live grant. */
export function pluginCommandScope(
  invocation: PluginInvocation<DocxPluginSnapshot>,
  grant: () => DocxPluginGrant,
  subscribe: (listener: () => void) => () => void
): DocxCommandScope {
  return {
    deny(id) {
      const refusal = invocation.refusal();
      if (refusal) return refusal;
      if (isPluginCommandId(id)) {
        return id.startsWith(`plugin:${invocation.pluginId}/`) ? null : 'permission-denied';
      }
      const current = grant();
      if (!grantsCommand(current, id)) return 'permission-denied';
      if (DOCX_COMMAND_DESCRIPTORS[id].mutatesDocument && !grantsWrite(current)) {
        return 'permission-denied';
      }
      return null;
    },
    subscribe,
  };
}

function navigationFailure(code: DocxPluginNavigationFailureCode, message: string) {
  return { ok: false as const, failure: { code, message } };
}

/** A unique body paragraph and its display position, or why there is none. */
export function resolveParagraph(
  session: YrsSession,
  target: { story: string; paraId: string }
): ReturnType<typeof resolveNavigationTarget> {
  const { story, paraId } = target ?? {};
  return resolveNavigationTarget(session, story, paraId);
}

export interface DocxPluginClients {
  read: DocxPluginReadClient;
  commands: DocxPluginCommandClient;
  edits: DocxPluginEditClient | null;
  navigation: DocxPluginNavigation;
}

/** Clients bound to one invocation: they refuse once it is superseded or its plugin ends. */
export function createPluginClients(
  invocation: PluginInvocation<DocxPluginSnapshot>,
  access: DocxPluginEditorAccess,
  grant: () => DocxPluginGrant,
  store: DocxCommandStore
): DocxPluginClients {
  /** Why `session` may not be touched now; call in the continuation that touches it. */
  const invalid = (session: YrsSession): DocxPluginRefusal | null => {
    const refused = refusalOf(invocation);
    if (refused) return refused;
    const editor = access.pagedEditorRef.current;
    return editor && editor.getYrsSession() === session ? null : pluginRefusal('document-replaced');
  };

  /** Flushes input, then runs `use` synchronously against the still-current session. */
  const whenFlushed = async <T>(
    use: (session: YrsSession, editor: PagedEditorRef) => T
  ): Promise<T | DocxPluginRefusal> => {
    const before = refusalOf(invocation);
    if (before) return before;
    const flush = await flushEditorInput(access.pagedEditorRef);
    const refused = refusalOf(invocation);
    if (refused) return refused;
    if (!flush.ok) {
      return pluginRefusal(flush.code === 'editor-unavailable' ? 'plugin-unavailable' : flush.code);
    }
    return invalid(flush.session) ?? use(flush.session, flush.editor);
  };

  const replicaReady = async (session: YrsSession): Promise<DocxPluginRefusal | null> => {
    const before = invalid(session);
    if (before) return before;
    try {
      await requestWorkerOpenReplica(session);
    } catch {
      return invalid(session) ?? pluginRefusal('input-failed');
    }
    return invalid(session);
  };

  const read: DocxPluginReadClient = {
    version: () => {
      const session = access.pagedEditorRef.current?.getYrsSession();
      if (session && workerProposalAuthority(session)) {
        return Promise.resolve(invalid(session) ?? { ok: true as const, version: session.version() });
      }
      return whenFlushed((session) => ({ ok: true as const, version: session.version() }));
    },
    readParagraphs: async (request) => {
      const session = access.pagedEditorRef.current?.getYrsSession();
      const authority = session ? workerProposalAuthority(session) : null;
      if (session && authority) {
        const before = invalid(session);
        if (before) return before;
        let fallbackRefusal: DocxPluginRefusal | null = null;
        try {
          const result = await authority.readParagraphs(request, async (request) => {
            const result = await whenFlushed((session) => session.readParagraphs(request));
            if ('version' in result) return result;
            fallbackRefusal = result;
            throw result;
          });
          return invalid(session) ?? result;
        } catch (error) {
          const refused = invalid(session);
          if (refused) return refused;
          if (fallbackRefusal && error === fallbackRefusal) return fallbackRefusal;
          throw error;
        }
      }
      return whenFlushed((session) => session.readParagraphs(request));
    },
    findText: async (request) => {
      const session = access.pagedEditorRef.current?.getYrsSession();
      if (session && workerProposalAuthority(session)) {
        const refused = await replicaReady(session);
        if (refused) return refused;
      }
      return whenFlushed((session) => session.findText(request));
    },
    validateEdits: (request) =>
      whenFlushed(
        (session) =>
          modeRefusal(session, access.writeMode(), request) ?? session.validateEdits(request)
      ),
  };

  const batchDenial = (history: 'separate' | 'none' | undefined): DocxPluginRefusal | null =>
    refusalOf(invocation) ??
    (grantsEditBatch(grant(), history) ? null : pluginRefusal('permission-denied'));

  const edits: DocxPluginEditClient | null = grantsEditBatch(grant(), 'separate')
    ? {
        async applyEdits(request) {
          const denied = batchDenial(request.history);
          if (denied) return denied;
          const outcome = await applyEditBatch(
            access.pagedEditorRef,
            access.writeMode,
            request,
            () => batchDenial(request.history),
            (write) => invocation.commit(write)
          );
          if (!('flush' in outcome)) return outcome.result;
          return (
            refusalOf(invocation) ??
            pluginRefusal(
              outcome.flush.code === 'editor-unavailable'
                ? 'plugin-unavailable'
                : outcome.flush.code
            )
          );
        },
      }
    : null;

  const commands = {
    getDescriptor: (id: never) => store.getDescriptor(id),
    getState: (id: never, args?: never) => store.getState(id, args),
    subscribe: (listener: () => void) => store.subscribe(listener),
    execute: async (id: never, args: never) => refusalOf(invocation) ?? store.execute(id, args),
  } as unknown as DocxPluginCommandClient;

  const navigationVersion = (session: YrsSession, version: string): string =>
    handedOverRequest(session, { expectVersion: version }).expectVersion;

  /** Resolves the target against the current version; the resolution or why it failed. */
  const locate = (
    session: YrsSession,
    target: { story: string; paraId: string },
    version: string
  ) => {
    if (session.version() !== navigationVersion(session, version)) {
      return navigationFailure('stale-version', 'The document changed after that version');
    }
    const resolved = resolveParagraph(session, target);
    return typeof resolved === 'string'
      ? navigationFailure(resolved, `The paragraph cannot be shown (${resolved})`)
      : resolved;
  };

  const locateWorker = async (
    session: YrsSession,
    target: { story: string; paraId: string },
    version: string,
    authority: WorkerProposalAuthority
  ) => {
    const before = invalid(session);
    if (before) return before;
    if (session.version() !== navigationVersion(session, version)) {
      return navigationFailure('stale-version', 'The document changed after that version');
    }
    let resolved: Awaited<ReturnType<WorkerProposalAuthority['navigationTarget']>>;
    try {
      resolved = await authority.navigationTarget(target?.story, target?.paraId, () =>
        resolveParagraph(session, target)
      );
    } catch (error) {
      const refused = invalid(session);
      if (refused) return refused;
      throw error;
    }
    const refused = invalid(session);
    if (refused) return refused;
    const expected = navigationVersion(session, version);
    if (session.version() !== expected || navigationVersion(session, resolved.version) !== expected) {
      return navigationFailure('stale-version', 'The document changed after that version');
    }
    return typeof resolved.target === 'string'
      ? navigationFailure(resolved.target, `The paragraph cannot be shown (${resolved.target})`)
      : resolved.target;
  };

  const layoutMatches = (
    session: YrsSession,
    version: string,
    queries: DisplayListQueries | null
  ): boolean =>
    queries !== null &&
    !isLayoutQueued(session) &&
    workerOpenSourceVersion(session, sourceVersionOf(queries)) === navigationVersion(session, version) &&
    renderedPreviewKey(queries) === currentPreviewKey(session);

  const settledLayout = (
    session: YrsSession,
    position: number,
    version: string,
    signal: AbortSignal
  ): Promise<boolean> =>
    new Promise((resolve) => {
      let done = false;
      let timer: ReturnType<typeof setTimeout> | undefined;
      let unsubscribe = () => {};
      const finish = (ready: boolean) => {
        if (done) return;
        done = true;
        if (timer !== undefined) clearTimeout(timer);
        unsubscribe();
        signal.removeEventListener('abort', cancelled);
        resolve(ready);
      };
      const cancelled = () => finish(false);
      const check = () => {
        if (done) return;
        const expected = navigationVersion(session, version);
        if (signal.aborted || invalid(session) || session.version() !== expected) {
          finish(false);
          return;
        }
        const { queries, complete, failed } = access.layout();
        if (failed) {
          finish(false);
          return;
        }
        if (queries && layoutMatches(session, version, queries)) {
          const source = queries.sourceState();
          if (source.status === 'error') {
            finish(false);
            return;
          }
          if (source.status === 'ready' && (queries.anchorRect(position) || complete)) {
            finish(true);
            return;
          }
        }
      };
      signal.addEventListener('abort', cancelled, { once: true });
      unsubscribe = access.subscribeLayout(check);
      if (done) unsubscribe();
      check();
      if (!done) timer = setTimeout(cancelled, LAYOUT_WAIT_MS);
    });

  const navigation: DocxPluginNavigation = {
    async scrollToParagraph(target, options) {
      const previous = navigationRequests.get(invocation.activation);
      const request = new AbortController();
      navigationRequests.set(invocation.activation, request);
      previous?.abort();
      const abort = () => request.abort();
      invocation.signal.addEventListener('abort', abort, { once: true });
      invocation.lifetimeSignal.addEventListener('abort', abort, { once: true });
      if (invocation.signal.aborted || invocation.lifetimeSignal.aborted) abort();
      try {
        const current = access.pagedEditorRef.current?.getYrsSession();
        const authority = current ? workerProposalAuthority(current) : null;
        const first = current && authority
          ? {
              session: current,
              located: await locateWorker(current, target, options.expectVersion, authority),
            }
          : await whenFlushed((session) => ({
              session,
              located: locate(session, target, options.expectVersion),
            }));
        if (!('session' in first)) return first;
        if ('ok' in first.located) return first.located;
        const session = first.session;
        const settled = await settledLayout(
          session,
          first.located.position,
          options.expectVersion,
          request.signal
        );
        const refused = invalid(session);
        if (refused) return refused;
        const currentAuthority = workerProposalAuthority(session);
        const located = currentAuthority
          ? first.located
          : locate(session, target, options.expectVersion);
        if ('ok' in located) return located;
        if (currentAuthority) {
          const refused = invalid(session);
          if (refused) return refused;
          if (session.version() !== navigationVersion(session, options.expectVersion)) {
            return navigationFailure('stale-version', 'The document changed after that version');
          }
        }
        if (
          !settled || request.signal.aborted ||
          !layoutMatches(session, options.expectVersion, access.layout().queries)
        ) {
          return navigationFailure(
            'layout-unavailable',
            request.signal.aborted
              ? 'A newer paragraph navigation superseded this request'
              : 'No rendered layout shows this version yet'
          );
        }
        if (options.focus && workerProposalAuthority(session)) {
          const refused = await replicaReady(session);
          if (refused) return refused;
          if (request.signal.aborted) {
            return navigationFailure(
              'layout-unavailable',
              'A newer paragraph navigation superseded this request'
            );
          }
        }
        const editor = access.pagedEditorRef.current!;
        const outcome = editor.revealDisplayPosition(located.position);
        if (outcome !== 'scrolled') {
          return navigationFailure(
            outcome,
            outcome === 'unsupported'
              ? 'The paragraph has no rendered position'
              : 'The layout cannot show this paragraph yet'
          );
        }
        if (options.focus) {
          session.setSelection(located.loc);
          editor.syncYrsInputState(false);
          editor.focus();
        }
        return { ok: true };
      } finally {
        invocation.signal.removeEventListener('abort', abort);
        invocation.lifetimeSignal.removeEventListener('abort', abort);
        if (navigationRequests.get(invocation.activation) === request) {
          navigationRequests.delete(invocation.activation);
        }
      }
    },
  };

  return { read, commands, edits, navigation };
}

/** The store plugin clients use: the editor's commands scoped to the plugin, or none. */
export function scopedCommandStore(
  controller: DocxCommandController | null,
  scope: DocxCommandScope
): DocxCommandStore {
  return controller?.scoped(scope) ?? UNAVAILABLE_DOCX_COMMANDS;
}
