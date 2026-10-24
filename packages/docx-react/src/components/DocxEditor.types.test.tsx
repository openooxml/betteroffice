import { expect, test } from 'bun:test';
import { memo, type ComponentProps } from 'react';
import { DocxEditor } from './DocxEditor';
import type { DocxAsyncEditorProps, DocxEditorBaseProps, DocxEditorProps, DocxInThreadEditorProps, DocxLegacyEditorProps } from '../index';
import { LEGACY_SYNC_API_PROP } from './DocxEditor/peerless/contracts';

test('Document callbacks require a literal sync opt-in', () => {
  const onChange: NonNullable<DocxEditorBaseProps['onChange']> = () => {};
  const check = (flag: boolean) => [
    <DocxEditor />,
    <DocxEditor onDocumentChange={() => {}} />,
    <DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }} onChange={onChange} />,
    <DocxEditor experimentalWorkerOpen={false} onChange={onChange} />,
    // @ts-expect-error
    <DocxEditor onChange={onChange} />,
    // @ts-expect-error
    <DocxEditor {...{ [LEGACY_SYNC_API_PROP]: flag }} onChange={onChange} />,
    // @ts-expect-error
    { onChange } satisfies DocxEditorProps,
    // @ts-expect-error
    { experimentalWorkerOpen: true, onChange } satisfies DocxEditorProps,
    // @ts-expect-error
    <DocxEditor experimentalWorkerOpen={flag} onChange={onChange} />,
  ];
  const asyncProps: DocxAsyncEditorProps = { onDocumentChange() {} };
  const legacy: DocxLegacyEditorProps = { [LEGACY_SYNC_API_PROP]: true, onChange };
  const inThread: DocxInThreadEditorProps = { experimentalWorkerOpen: false, onChange };
  expect([asyncProps, legacy, inThread]).toHaveLength(3);
  expect(check).toBeFunction();
});

test('experimentalWorkerOpen accepts a computed boolean', () => {
  const check = (flag: boolean) => (
    <DocxEditor experimentalWorkerOpen={flag} documentBuffer={null} onDocumentChange={() => {}} onSave={() => {}} />
  );
  expect(check).toBeFunction();
});

test('the exported props union forwards through React wrappers', () => {
  const Wrapped = memo(DocxEditor);
  function Forward(props: DocxEditorProps) { return <DocxEditor {...props} />; }
  const check = (props: ComponentProps<typeof DocxEditor>) => [<Wrapped {...props} />, <Forward {...props} />];
  expect(check).toBeFunction();
});

test('file and fileName are accepted on every props variant', () => {
  const onChange: NonNullable<DocxEditorBaseProps['onChange']> = () => {};
  const check = (bytes: Uint8Array, blob: Blob, upload: File, buffer: ArrayBuffer) => [
    <DocxEditor file={bytes} fileName="a.docx" />,
    <DocxEditor file={blob} />,
    <DocxEditor file={upload} fileName={upload.name} onDocumentChange={() => {}} />,
    <DocxEditor file={null} />,
    <DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }} file={buffer} fileName="a.docx" onChange={onChange} />,
    <DocxEditor experimentalWorkerOpen={false} file={buffer} fileName="a.docx" onChange={onChange} />,
    // @ts-expect-error
    <DocxEditor file="a.docx" />,
  ];
  const asyncProps: DocxAsyncEditorProps = { file: new Uint8Array(), fileName: 'a.docx' };
  const legacy: DocxLegacyEditorProps = { [LEGACY_SYNC_API_PROP]: true, file: new ArrayBuffer(0), fileName: 'a.docx' };
  const inThread: DocxInThreadEditorProps = { experimentalWorkerOpen: false, file: new Blob(), fileName: 'a.docx' };
  expect([asyncProps, legacy, inThread]).toHaveLength(3);
  expect(check).toBeFunction();
});
