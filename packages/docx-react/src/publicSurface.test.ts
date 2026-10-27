import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  DOCX_REF_ASYNC_TWINS,
  DOCX_REF_ASYNC_TWIN_EXEMPTIONS,
  DOCX_REF_REPLICA_ACCESS,
  DocxAsyncOnlyError,
  DocxReplicaNotReadyError,
} from './components/DocxEditor/hooks/useDocxEditorRefApi';
import { DocxWorkerError } from './components/DocxEditor/internals/docxWorkerError';
import { DocxCommandAdmissionError } from './commands/createDocxCommandStore';
import { LEGACY_SYNC_API_PROP } from './components/DocxEditor/peerless/contracts';
import type {
  DocxCommandAdmissionCode,
  DocxCommentInsertion,
  DocxDocumentChange,
  DocxParagraphMatch,
  DocxParagraphList,
  DocxSelectionInfo,
} from './index';

const editor = readFileSync(resolve(import.meta.dir, 'components/DocxEditor.tsx'), 'utf8');
const plugins = readFileSync(resolve(import.meta.dir, 'plugins/types.ts'), 'utf8');
const index = readFileSync(resolve(import.meta.dir, 'index.ts'), 'utf8');

function memberDocs(source: string, name: string): Map<string, string> {
  const body = source.match(new RegExp(`^export interface ${name}\\b[^\\n]*\\{\\r?\\n([\\s\\S]*?)^}`, 'm'))?.[1];
  if (body === undefined) throw new Error(`Missing interface ${name}`);
  const docs = new Map<string, string>();
  for (const member of body.matchAll(/^  (?:readonly )?(\[[\w$]+\]|[\w$]+)\??\s*(?=[:(<])/gm)) {
    const before = body.slice(0, member.index);
    const comment = before.slice(before.lastIndexOf('/**'));
    const end = comment.indexOf('*/') + 2;
    const key = member[1] === '[LEGACY_SYNC_API_PROP]' ? LEGACY_SYNC_API_PROP : member[1]!;
    docs.set(key, end > 1 && comment.slice(end).trim() === '' ? comment.slice(0, end) : '');
  }
  return docs;
}

test('deprecated ref members have public non-deprecated twins or explicit exemptions', () => {
  const docs = memberDocs(editor, 'DocxEditorRef');
  for (const [member, doc] of docs) {
    if (!doc.includes('@deprecated')) continue;
    expect(member in DOCX_REF_ASYNC_TWINS || DOCX_REF_ASYNC_TWIN_EXEMPTIONS.has(member)).toBe(true);
  }
  for (const [member, twins] of Object.entries(DOCX_REF_ASYNC_TWINS)) {
    expect(docs.get(member)).toContain('@deprecated');
    for (const twin of typeof twins === 'string' ? [twins] : twins) {
      expect(docs.has(twin)).toBe(true);
      expect(docs.get(twin)).not.toContain('@deprecated');
      expect(docs.get(member)).toContain(`{@link ${twin}}`);
    }
  }
  expect([...docs.keys()].sort()).toEqual(Object.keys(DOCX_REF_REPLICA_ACCESS).sort());
});

test('deprecated props and plugin geometry retain their replacements', () => {
  const props = memberDocs(editor, 'DocxEditorBaseProps');
  expect(props.has('experimentalWorkerOpen')).toBe(true);
  expect(props.get('experimentalWorkerOpen')).toContain('@deprecated');
  expect(editor).toMatch(/^  experimentalWorkerOpen\?: boolean;$/m);
  const legacy = memberDocs(editor, 'DocxLegacyEditorProps');
  expect(legacy.get(LEGACY_SYNC_API_PROP)).toContain('@deprecated');
  expect(legacy.get(LEGACY_SYNC_API_PROP)).toContain('0.6.0');
  const propExemptions = new Set(['pluginOverlays', 'pluginSidebarItems', 'pluginRenderedDomContext']);
  for (const [member, doc] of props) {
    if (!doc.includes('@deprecated') || propExemptions.has(member)) continue;
    if (member === 'experimentalWorkerOpen') continue;
    expect(member).toBe('onChange');
    expect(doc).toContain('{@link onDocumentChange}');
    expect(props.has('onDocumentChange')).toBe(true);
    expect(props.get('onDocumentChange')).not.toContain('@deprecated');
  }
  const geometry = memberDocs(plugins, 'DocxPluginGeometry');
  for (const [member, doc] of geometry) {
    if (!doc.includes('@deprecated')) continue;
    const twins = DOCX_REF_ASYNC_TWINS[member as keyof typeof DOCX_REF_ASYNC_TWINS];
    expect(twins).toBeDefined();
    for (const twin of typeof twins === 'string' ? [twins] : twins) {
      expect(geometry.has(twin)).toBe(true);
      expect(geometry.get(twin)).not.toContain('@deprecated');
      expect(doc).toContain(`{@link ${twin}}`);
    }
  }
});

test('errors and new public types are exported from the root', () => {
  const [AsyncOnly, NotReady]: [typeof import('./index').DocxAsyncOnlyError, typeof import('./index').DocxReplicaNotReadyError] = [DocxAsyncOnlyError, DocxReplicaNotReadyError];
  const asyncOnly = new AsyncOnly('getDocument', 'readParagraphs');
  expect(asyncOnly).toBeInstanceOf(Error);
  expect(asyncOnly).not.toBeInstanceOf(NotReady);
  expect(new NotReady('getDocument')).not.toBeInstanceOf(AsyncOnly);
  const WorkerError: typeof import('./index').DocxWorkerError = DocxWorkerError;
  const cause = new Error('Worker stopped');
  for (const stage of ['open', 'layout', 'render'] as const) {
    const workerError = new WorkerError(stage, cause);
    expect(workerError).toBeInstanceOf(Error);
    expect(workerError.name).toBe('DocxWorkerError');
    expect(workerError.stage).toBe(stage);
    expect(workerError.cause).toBe(cause);
  }
  expect(new WorkerError('open').cause).toBeUndefined();
  expect(index).toMatch(/export\s*\{[^}]*\bDocxWorkerError\b[^}]*\}\s*from\s*['"]\.\/components\/DocxEditor\/internals\/docxWorkerError['"]/);
  const Admission: typeof import('./index').DocxCommandAdmissionError = DocxCommandAdmissionError;
  const codes: DocxCommandAdmissionCode[] = ['input-failed', 'document-replaced', 'target-changed', 'editor-unavailable'];
  for (const code of codes) {
    const refusal = new Admission(code);
    expect(refusal).toBeInstanceOf(Error);
    expect(refusal).toMatchObject({ name: 'DocxCommandAdmissionError', code, message: code });
  }
  // @ts-expect-error
  const notAdmission: DocxCommandAdmissionCode = 'command-failed';
  expect(String(notAdmission)).toBe('command-failed');
  expect(index).toMatch(/export\s*\{[^}]*\bDocxCommandAdmissionError\b[^}]*\bDocxCommandAdmissionCode\b[^}]*\}\s*from\s*['"]\.\/commands\/createDocxCommandStore['"]/);
  const names = ['DocxParagraphList', 'DocxCommandAdmissionError', 'DocxCommandAdmissionCode', 'DocxWorkerError', 'DocxWorkerEditorChunkLoadError', 'DocxSyncApiError', 'DocxEditorNotReadyError', 'DocxAsyncOnlyError', 'DocxReplicaNotReadyError', 'DocxParagraphMatch', 'DocxSelectionInfo', 'DocxCommentInsertion', 'DocxDocumentChange', 'DocxEditorBaseProps', 'DocxAsyncEditorProps', 'DocxLegacyEditorProps', 'DocxInThreadEditorProps'];
  for (const name of names) expect(index).toMatch(new RegExp(`\\b${name}\\b`));
  const match: DocxParagraphMatch = { paraId: 'p', match: 'text', before: '', after: '' };
  const selection: DocxSelectionInfo = { paraId: null, selectedText: '', paragraphText: '', before: '', after: '' };
  const insertion: DocxCommentInsertion = { paraId: 'p', text: 'text', author: 'Author' };
  const change: DocxDocumentChange = { version: 'v' };
  expect([match.paraId, selection.paraId, insertion.paraId, change.version]).toEqual(['p', null, 'p', 'v']);
});


test('plugin notification drains stay internal', () => {
  expect(index).not.toMatch(/\b(?:DocxPluginHost|createDocxPluginHost|notificationsIdle)\b/);
  expect(plugins).not.toContain('notificationsIdle');
  expect(editor).not.toContain('notificationsIdle');
});

test('paragraph list is exported by both public roots', () => {
  const list: DocxParagraphList = { marker: '1.', level: 0, numId: '1', format: 'decimal' };
  const core: import('@betteroffice/docx').DocxParagraphList = list;
  expect(core).toEqual(list);
  expect(index).toMatch(/export type \{[^}]*\bDocxParagraphList\b[^}]*\} from '@betteroffice\/docx\/yrs'/);
});
