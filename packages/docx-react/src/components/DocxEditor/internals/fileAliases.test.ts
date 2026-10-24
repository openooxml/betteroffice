import { afterEach, expect, spyOn, test } from 'bun:test';
import type { DocxInput } from '@betteroffice/docx/utils';
import {
  resetDocxFileAliasWarningsForTests,
  resolveDocxFileAliases as resolveAliases,
  warnDocxFileAliasConflicts,
} from './fileAliases';

interface AliasInput {
  documentBuffer?: DocxInput | null;
  documentName?: string;
  file?: DocxInput | null;
  fileName?: string;
  readOnly?: boolean;
}

const resolveDocxFileAliases = (props: AliasInput) => resolveAliases(props);

afterEach(() => resetDocxFileAliasWarningsForTests());

test('file and fileName stand in for documentBuffer and documentName', () => {
  for (const file of [new ArrayBuffer(2), new Uint8Array([1]), new Blob(['x']), new File(['x'], 'a.docx')]) {
    const { props, conflicts } = resolveDocxFileAliases({ file, fileName: 'Report.docx', readOnly: true });
    expect(props).toEqual({ documentBuffer: file, documentName: 'Report.docx', readOnly: true });
    expect(props.documentBuffer).toBe(file);
    expect('file' in props || 'fileName' in props).toBe(false);
    expect(conflicts).toEqual([]);
  }
  expect(resolveDocxFileAliases({ file: null }).props).toEqual({ documentBuffer: null });
});

test('props without aliases pass through unchanged', () => {
  const input = { documentBuffer: new Uint8Array([1]), documentName: 'a.docx' };
  expect(resolveDocxFileAliases(input)).toEqual({ props: input, conflicts: [] });
  expect(resolveDocxFileAliases(input).props).toBe(input);
});

test('documentBuffer and documentName win over the aliases and name the conflict', () => {
  const buffer = new Uint8Array([1]);
  const file = new Uint8Array([2]);
  const { props, conflicts } = resolveDocxFileAliases({ documentBuffer: buffer, file, documentName: 'old.docx', fileName: 'new.docx' });
  expect(props).toEqual({ documentBuffer: buffer, documentName: 'old.docx' });
  expect(props.documentBuffer).toBe(buffer);
  expect(conflicts).toEqual([
    '[DocxEditor] Both documentBuffer and file are set; using documentBuffer. Pass only one.',
    '[DocxEditor] Both documentName and fileName are set; using documentName. Pass only one.',
  ]);
  const nullBuffer = resolveDocxFileAliases({ documentBuffer: null, file });
  expect(nullBuffer.props).toEqual({ documentBuffer: null });
  expect(nullBuffer.conflicts).toHaveLength(1);
  const mixed = resolveDocxFileAliases({ documentBuffer: buffer, fileName: 'new.docx' });
  expect(mixed).toEqual({ props: { documentBuffer: buffer, documentName: 'new.docx' }, conflicts: [] });
});

test('conflict warnings print once in development and never in production', () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const env = process.env.NODE_ENV;
  try {
    const { conflicts } = resolveDocxFileAliases({ documentBuffer: new Uint8Array(), file: new Uint8Array() });
    process.env.NODE_ENV = 'production';
    warnDocxFileAliasConflicts(conflicts);
    expect(warn).not.toHaveBeenCalled();
    process.env.NODE_ENV = 'development';
    warnDocxFileAliasConflicts(conflicts);
    warnDocxFileAliasConflicts(conflicts);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(String(warn.mock.calls[0]?.[0])).toContain('documentBuffer and file');
  } finally {
    if (env === undefined) delete process.env.NODE_ENV;
    else process.env.NODE_ENV = env;
    warn.mockRestore();
  }
});
