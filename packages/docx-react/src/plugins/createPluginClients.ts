import { createYrsSidebarProjection } from '@betteroffice/docx/layout/render';
import type { YrsLoc, YrsSession } from '@betteroffice/docx/yrs';
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
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
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
  /** Resolves true once a rendered layout shows `version`, false if none does soon. */
  settledLayout(version: string): Promise<boolean>;
}

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
): { loc: YrsLoc; position: number } | DocxPluginNavigationFailureCode {
  const { story, paraId } = target ?? {};
  if (typeof story !== 'string' || typeof paraId !== 'string') return 'missing-target';
  if (!session.storyIds().includes(story)) return 'missing-target';
  const matches = session.paragraphs(story).filter((paragraph) => paragraph.paraId === paraId);
  if (matches.length === 0) return 'missing-target';
  if (matches.length > 1) return 'ambiguous-target';
  const loc = { story, paraId, offset: 0 };
  const point = createYrsSidebarProjection(session).locToDisplayPoint(loc);
  if (!point || point.hfRid) return 'unsupported';
  return { loc, position: point.position };
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

  const read: DocxPluginReadClient = {
    version: () => whenFlushed((session) => ({ ok: true as const, version: session.version() })),
    readParagraphs: (request) => whenFlushed((session) => session.readParagraphs(request)),
    findText: (request) => whenFlushed((session) => session.findText(request)),
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

  /** Resolves the target against the current version; the resolution or why it failed. */
  const locate = (
    session: YrsSession,
    target: { story: string; paraId: string },
    version: string
  ) => {
    if (session.version() !== version) {
      return navigationFailure('stale-version', 'The document changed after that version');
    }
    const resolved = resolveParagraph(session, target);
    return typeof resolved === 'string'
      ? navigationFailure(resolved, `The paragraph cannot be shown (${resolved})`)
      : resolved;
  };

  const navigation: DocxPluginNavigation = {
    async scrollToParagraph(target, options) {
      const first = await whenFlushed((session) => ({
        session,
        located: locate(session, target, options.expectVersion),
      }));
      if (!('session' in first)) return first;
      if ('ok' in first.located) return first.located;
      const settled = await access.settledLayout(options.expectVersion);
      const session = first.session;
      const refused = invalid(session);
      if (refused) return refused;
      const located = locate(session, target, options.expectVersion);
      if ('ok' in located) return located;
      if (!settled) {
        return navigationFailure('layout-unavailable', 'No rendered layout shows this version yet');
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
