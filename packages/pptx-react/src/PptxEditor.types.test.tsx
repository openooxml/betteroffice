import { expect, test } from 'bun:test';
import { createElement, memo } from 'react';
import type { ComponentProps } from 'react';
import type { AsyncCollaborationReplica, CollaborationReplica, PresentationHandle } from '@betteroffice/pptx';
import { PptxEditor } from './PptxEditor';
import { definePptxPlugin } from './plugins/definePptxPlugin';
import type { PptxAsyncEditorProps, PptxEditorApi, PptxEditorProps, PptxWorkerViewerApi } from './PptxEditor';
import type { PptxAsyncEditorAccess, PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';
import type { PptxCommandAdmissionCode } from './index';
import type { PptxWorkerEditorApi } from './worker/createWorkerEditorApi';

type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends
  (<T>() => T extends B ? 1 : 2) ? true : false;
function assertType<T extends true>(_value?: T): void {}

test('default and in-thread async props preserve async contextual APIs', () => {
  const check = () => [
    <PptxEditor fonts={[]} onReady={(api) => {
      assertType<Equal<typeof api, PptxAsyncEditorApi>>();
      assertType<Equal<ReturnType<typeof api.goToSlide>, Promise<boolean>>>();
      assertType<Equal<ReturnType<typeof api.clearSelection>, Promise<void>>>();
      // @ts-expect-error
      api.handle;
      // @ts-expect-error
      api.save();
    }} />,
    <PptxEditor fonts={[]} experimentalWorkerOpen={false} onReady={(api) => {
      assertType<Equal<typeof api, PptxAsyncEditorApi>>();
    }} />,
    <PptxEditor fonts={[]} readOnly onChange={() => {}} onReady={(api) => {
      assertType<Equal<typeof api, PptxAsyncEditorApi>>();
    }} />,
    <PptxEditor fonts={[]} collaboration={{ clientId: 1, onReplica: (replica) => {
      assertType<Equal<typeof replica, AsyncCollaborationReplica | null>>();
    } }} />,
  ];
  expect(check).toBeFunction();
});

test('plugin props preserve async and legacy contextual APIs', () => {
  const plugins = [definePptxPlugin({ id: 'review', createState: () => null })];
  const check = () => [
    <PptxEditor fonts={[]} plugins={plugins} pluginGrants={{}} onPluginError={() => {}}
      onReady={(api) => { assertType<Equal<typeof api, PptxAsyncEditorApi>>(); }} />,
    <PptxEditor fonts={[]} experimentalWorkerOpen={false} plugins={plugins}
      onReady={(api) => { assertType<Equal<typeof api, PptxAsyncEditorApi>>(); }} />,
    { fonts: [], plugins, pluginGrants: {}, onPluginError: () => {} } satisfies PptxAsyncEditorProps,
    { fonts: [], plugins } satisfies PptxEditorProps,
    <PptxEditor legacySyncApi fonts={[]} plugins={plugins}
      onReady={(api) => { assertType<Equal<typeof api, PptxEditorApi>>(); }} />,
    <PptxEditor legacySyncApi experimentalWorkerOpen fonts={[]} plugins={plugins}
      onReady={(api) => { assertType<Equal<typeof api, PptxWorkerEditorApi>>(); }} />,
    <PptxEditor legacySyncApi experimentalWorkerOpen={false} fonts={[]} plugins={plugins}
      onReady={(api) => { assertType<Equal<typeof api, PptxEditorApi>>(); }} />,
  ];
  expect(check).toBeFunction();
});

test('only legacy sync props expose sync contextual APIs', () => {
  const check = () => [
    <PptxEditor legacySyncApi fonts={[]} onReady={(api) => {
      assertType<Equal<typeof api, PptxEditorApi>>();
      assertType<Equal<ReturnType<typeof api.save>, Uint8Array>>();
    }} />,
    <PptxEditor legacySyncApi experimentalWorkerOpen fonts={[]} onReady={(api) => {
      assertType<Equal<typeof api, PptxWorkerEditorApi>>();
      assertType<Equal<ReturnType<typeof api.goToSlide>, boolean>>();
    }} />,
    <PptxEditor legacySyncApi fonts={[]} experimentalWorkerOpen={false} onReady={(api) => {
      assertType<Equal<typeof api, PptxEditorApi>>();
      api.handle.snapshot();
      api.save();
    }} />,
    <PptxEditor legacySyncApi fonts={[]} readOnly onReady={(api) => {
      assertType<Equal<typeof api, PptxEditorApi>>();
    }} />,
    <PptxEditor legacySyncApi experimentalWorkerOpen fonts={[]} readOnly onReady={(api) => {
      assertType<Equal<typeof api, PptxWorkerViewerApi>>();
    }} />,
    <PptxEditor legacySyncApi fonts={[]} collaboration={{ clientId: 1, onReplica: (replica) => {
      assertType<Equal<typeof replica, CollaborationReplica | null>>();
    } }} onReady={(api) => { assertType<Equal<typeof api, PptxEditorApi>>(); }} />,
  ];
  expect(check).toBeFunction();
});

test('exported union narrows and forwards through React wrappers', () => {
  assertType<Equal<ComponentProps<typeof PptxEditor>, PptxEditorProps>>();
  const Wrapped = memo(PptxEditor);
  function Forward(props: PptxEditorProps) { return <PptxEditor {...props} />; }
  const check = (props: PptxEditorProps) => {
    if (!props.legacySyncApi) {
      assertType<Equal<typeof props, PptxAsyncEditorProps>>();
      props.onReady?.({} as PptxAsyncEditorApi);
    }
    const options: PptxEditorProps = { fonts: [], onReady: (api) => {
      assertType<Equal<typeof api, PptxAsyncEditorApi>>();
      return () => api.focus();
    } };
    return [<Forward {...options} />, <Wrapped fonts={[]} onReady={(api) => {
      assertType<Equal<typeof api, PptxAsyncEditorApi>>();
    }} />, createElement<PptxAsyncEditorProps>(PptxEditor, { fonts: [], onReady: (api) => {
      assertType<Equal<typeof api, PptxAsyncEditorApi>>();
    } }), createElement(PptxEditor, options), <PptxEditor {...props} />];
  };
  expect(check).toBeFunction();
});

test('rejects sync callbacks without a literal legacy opt-in', () => {
  const check = (flag: boolean, sync: (api: PptxEditorApi) => void) => [
    // @ts-expect-error
    <PptxEditor fonts={[]} onReady={sync} />,
    // @ts-expect-error
    <PptxEditor fonts={[]} experimentalWorkerOpen={false} onReady={sync} />,
    // @ts-expect-error
    <PptxEditor fonts={[]} legacySyncApi={flag} onReady={sync} />,
    // @ts-expect-error
    { fonts: [], onReady: sync } satisfies PptxEditorProps,
  ];
  expect(check).toBeFunction();
});

test('async editor parity types', () => {
  const check = () => [
    <PptxEditor fonts={[]} collaboration={{ onReplica: (replica) => {
      assertType<Equal<typeof replica, AsyncCollaborationReplica | null>>();
    } }} />,
    <PptxEditor fonts={[]} collaboration={{ clientId: 7 }} />,
  ];
  assertType<Equal<AsyncCollaborationReplica['encodeStateAsUpdateChunks'],
    ((remoteStateVector: Uint8Array, maxChunkBytes: number) => Promise<Uint8Array[]>) | undefined>>();
  assertType<Equal<AsyncCollaborationReplica['applyUpdateChunks'],
    ((updates: readonly Uint8Array[]) => Promise<unknown>) | undefined>>();
  assertType<Equal<Awaited<ReturnType<PptxAsyncEditorApi['exportStructured']>>, ReturnType<PresentationHandle['exportStructured']>>>();
  assertType<Equal<Awaited<ReturnType<PptxAsyncEditorApi['exportMarkdown']>>, ReturnType<PresentationHandle['exportMarkdown']>>>();
  assertType<Equal<PptxAsyncEditorApi['exportStructured'], PptxAsyncEditorAccess['exportStructured']>>();
  assertType<Equal<PptxCommandAdmissionCode, 'input-failed' | 'document-replaced' | 'target-changed' | 'gesture-active'
    | 'editor-unavailable' | 'read-only' | 'invalid-arguments'>>();
  expect(check).toBeFunction();
});
