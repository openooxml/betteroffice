import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Document } from '@betteroffice/docx/types/document';
import { parseDocx } from '@betteroffice/docx/docx';
import { useImageActions } from './useImageActions';

// Two sections; the body sectPr inherits the first section's header and footer.
const probe = await parseDocx(
  Uint8Array.from(readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx')))
    .buffer
);

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

describe('useImageActions', () => {
  test('note properties reach both views of the last section', () => {
    const pushed: Document[] = [];
    const { result } = renderHook(() =>
      useImageActions({
        document: probe,
        pmImageContext: null,
        displayListQueries: null,
        pagedEditorRef: { current: null },
        focusActiveEditor: () => {},
        pushDocument: (doc) => {
          pushed.push(doc);
        },
      })
    );
    act(() => result.current.handleApplyFootnoteProperties({ numStart: 2 }, { numStart: 3 }));
    expect(pushed).toHaveLength(1);
    const body = pushed[0].package.document;
    const notes = { footnotePr: { numStart: 2 }, endnotePr: { numStart: 3 } };
    expect(body.finalSectionProperties).toMatchObject(notes);
    expect(body.finalSectionProperties?.headerReferences).toBeUndefined();
    expect(body.sections?.at(-1)?.properties).toMatchObject({
      ...notes,
      headerReferences: [{ type: 'default', rId: 'rId3' }],
    });
  });
});
