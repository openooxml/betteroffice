import { describe, expect, it } from 'bun:test';
import { PEERLESS_FLAG, XlsxEditor, type XlsxEditorCollaborationOptions } from './XlsxEditor';
import type { XlsxAsyncEditorApi, XlsxAsyncEditorProps } from './index';
import { defineXlsxPlugin } from './plugins/defineXlsxPlugin';
import type { XlsxPlugin } from './plugins/types';

type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2) ? true : false;
function expectType<T extends true>(value: T) { expect(value as boolean).toBe(true); }
const file = new Uint8Array([1, 2, 3]);
const asyncProps = { [PEERLESS_FLAG]: true } as const;
const plugin = defineXlsxPlugin({ id: 'writer', createState: () => null });

describe('XlsxEditor prop types', () => {
  it('types the async editor before the existing engine overloads', () => {
    const check = (readOnly: boolean, plugins: readonly XlsxPlugin[], collaboration: XlsxEditorCollaborationOptions | undefined) => [
      <XlsxEditor {...asyncProps} file={file} onReady={(api) => {
        expectType<Equal<typeof api, XlsxAsyncEditorApi>>(true);
        expectType<Equal<typeof api.hydrated, boolean>>(true);
        expectType<Equal<typeof api.failure, Error | null>>(true);
        expectType<Equal<ReturnType<typeof api.whenHydrated>, Promise<void>>>(true);
        return () => {};
      }} />,
      <XlsxEditor {...asyncProps} file={file} readOnly plugins={[]} onReady={(api) => {
        expectType<Equal<typeof api, XlsxAsyncEditorApi>>(true);
      }} />,
      <XlsxEditor {...asyncProps} file={file} readOnly={readOnly} plugins={plugins} collaboration={collaboration} experimentalWorkerOpen onReady={(api) => {
        expectType<Equal<typeof api, XlsxAsyncEditorApi>>(true);
      }} />,
      <XlsxEditor {...asyncProps} file={file} collaboration={{}} onReady={(api) => {
        expectType<Equal<typeof api, XlsxAsyncEditorApi>>(true);
      }} />,
      // @ts-expect-error
      <XlsxEditor {...asyncProps} file={file} experimentalWorkerOpen={false} />,
    ];
    const props: XlsxAsyncEditorProps = { [PEERLESS_FLAG]: true, file, plugins: [plugin], collaboration: {}, onError: () => {}, onReady: (api) => {
      expectType<Equal<typeof api, XlsxAsyncEditorApi>>(true);
    } };
    expectType<Equal<XlsxAsyncEditorProps['collaboration'], XlsxEditorCollaborationOptions | undefined>>(true);
    expectType<Equal<XlsxAsyncEditorProps['experimentalWorkerOpen'], true | undefined>>(true);
    expect(<XlsxEditor {...props} />).toBeDefined();
    expect(check).toBeDefined();
  });

});
