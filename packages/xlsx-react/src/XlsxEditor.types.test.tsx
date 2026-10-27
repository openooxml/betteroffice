import { expect, test } from 'bun:test';
import { createElement, memo } from 'react';
import type { ComponentProps } from 'react';
import { XlsxEditor } from './XlsxEditor';
import type { XlsxAsyncEditorProps, XlsxEditorApi, XlsxEditorProps, XlsxWorkerViewerApi } from './XlsxEditor';
import type { XlsxAsyncEditorApi } from './worker/createAsyncEditorApi';
import type { XlsxWorkerEditorApi } from './worker/createWorkerEditorApi';

type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends
  (<T>() => T extends B ? 1 : 2) ? true : false;
function assertType<T extends true>(_value?: T): void {}

test('default and in-thread props preserve the async contextual API', () => {
  const check = (readOnly: boolean) => [
    <XlsxEditor onReady={(api) => {
      assertType<Equal<typeof api, XlsxAsyncEditorApi>>();
      assertType<Equal<ReturnType<typeof api.exportStructured>, Promise<ReturnType<import('@betteroffice/xlsx').WorkbookHandle['exportStructured']>>>>();
      // @ts-expect-error
      api.handle;
      assertType<Equal<ReturnType<typeof api.save>, Promise<Uint8Array>>>();
    }} />,
    <XlsxEditor experimentalWorkerOpen={false} onReady={(api) => {
      assertType<Equal<typeof api, XlsxAsyncEditorApi>>();
    }} />,
    <XlsxEditor readOnly={readOnly} plugins={[]} collaboration={{}} onReady={(api) => {
      assertType<Equal<typeof api, XlsxAsyncEditorApi>>();
    }} />,
  ];
  expect(check).toBeFunction();
});

test('only legacy sync props expose sync contextual APIs', () => {
  const check = (readOnly: boolean) => [
    <XlsxEditor legacySyncApi onReady={(api) => {
      assertType<Equal<typeof api, XlsxEditorApi>>();
      assertType<Equal<ReturnType<typeof api.save>, Uint8Array>>();
    }} />,
    <XlsxEditor legacySyncApi experimentalWorkerOpen onReady={(api) => {
      assertType<Equal<typeof api, XlsxWorkerEditorApi>>();
    }} />,
    <XlsxEditor legacySyncApi experimentalWorkerOpen={false} onReady={(api) => {
      assertType<Equal<typeof api, XlsxEditorApi>>();
      api.handle.sheetInfo();
      api.save();
    }} />,
    <XlsxEditor legacySyncApi readOnly onReady={(api) => {
      assertType<Equal<typeof api, XlsxEditorApi>>();
    }} />,
    <XlsxEditor legacySyncApi experimentalWorkerOpen readOnly onReady={(api) => {
      assertType<Equal<typeof api, XlsxWorkerViewerApi>>();
    }} />,
    <XlsxEditor legacySyncApi readOnly={readOnly} onReady={(api) => {
      assertType<Equal<typeof api, XlsxEditorApi>>();
    }} />,
    <XlsxEditor legacySyncApi experimentalWorkerOpen readOnly={readOnly} onReady={(api) => {
      assertType<Equal<typeof api, XlsxWorkerEditorApi | XlsxWorkerViewerApi>>();
    }} />,
    <XlsxEditor legacySyncApi experimentalWorkerOpen={false} collaboration={{}} onReady={(api) => {
      assertType<Equal<typeof api, XlsxEditorApi>>();
    }} />,
  ];
  expect(check).toBeFunction();
});

test('exported union narrows and forwards through React wrappers', () => {
  assertType<Equal<ComponentProps<typeof XlsxEditor>, XlsxEditorProps>>();
  const Wrapped = memo(XlsxEditor);
  function Forward(props: XlsxEditorProps) { return <XlsxEditor {...props} />; }
  const check = (props: XlsxEditorProps) => {
    if (!props.legacySyncApi) {
      assertType<Equal<typeof props, XlsxAsyncEditorProps>>();
      props.onReady?.({} as XlsxAsyncEditorApi);
    }
    const options: XlsxEditorProps = { onReady: (api) => {
      assertType<Equal<typeof api, XlsxAsyncEditorApi>>();
      return () => api.focus();
    } };
    return [<Forward {...options} />, <Wrapped onReady={(api) => {
      assertType<Equal<typeof api, XlsxAsyncEditorApi>>();
    }} />, createElement<XlsxAsyncEditorProps>(XlsxEditor, { onReady: (api) => {
      assertType<Equal<typeof api, XlsxAsyncEditorApi>>();
    } }), createElement(XlsxEditor, options), <XlsxEditor {...props} />];
  };
  expect(check).toBeFunction();
});

test('rejects sync callbacks without a literal legacy opt-in', () => {
  const check = (flag: boolean, sync: (api: XlsxEditorApi) => void) => [
    // @ts-expect-error
    <XlsxEditor onReady={sync} />,
    // @ts-expect-error
    <XlsxEditor experimentalWorkerOpen={false} onReady={sync} />,
    // @ts-expect-error
    <XlsxEditor legacySyncApi={flag} onReady={sync} />,
    // @ts-expect-error
    { onReady: sync } satisfies XlsxEditorProps,
  ];
  expect(check).toBeFunction();
});

test('accepts async host props and document change types', () => {
  const props = { file: new ArrayBuffer(0), initialZoom: 1.5, loadingIndicator: <span>Opening</span>,
    style: { height: 400 }, downloadOnSave: false,
    onDocumentChange: (change: import('./index').XlsxDocumentChange) => { assertType<Equal<typeof change.version, string>>(); },
  } satisfies XlsxAsyncEditorProps;
  expect(props.downloadOnSave).toBe(false);
});

test('keeps new host options off deprecated legacy props', () => {
  const legacy = {
    legacySyncApi: true,
    // @ts-expect-error
    initialZoom: 2,
  } satisfies import('./XlsxEditor').XlsxLegacyEditorProps;
  expect(legacy.legacySyncApi).toBe(true);
});
