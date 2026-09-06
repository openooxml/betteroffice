import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Document } from '@betteroffice/docx/types/document';
import { parseDocx } from '@betteroffice/docx/docx';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import { usePageSetupControls } from './usePageSetupControls';

// Two sections; the body sectPr inherits the first section's header and footer.
const probe = await parseDocx(
  Uint8Array.from(readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx')))
    .buffer
);
const inheritedHeader = [{ type: 'default', rId: 'rId3' }];

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

function render() {
  const changed: Document[] = [];
  const hook = renderHook(() =>
    usePageSetupControls({
      document: probe,
      readOnly: false,
      handleDocumentChange: (doc) => {
        changed.push(doc);
      },
      pagedEditorRef: { current: null },
    })
  );
  return { hook, changed };
}

describe('usePageSetupControls', () => {
  test('a margin change reaches the layout request through the resolved last section', () => {
    const { hook, changed } = render();
    act(() => hook.result.current.handleLeftMarginChange(2880));
    expect(changed).toHaveLength(1);
    const request = buildResidentRegionLayoutRequest(changed[0], 24, {});
    expect(request.regions.sections.at(-1)?.properties.marginLeft).toBe(2880);
    const body = changed[0].package.document;
    expect(body.finalSectionProperties?.marginLeft).toBe(2880);
    expect(body.sections?.at(-1)?.properties).toMatchObject({
      marginLeft: 2880,
      headerReferences: inheritedHeader,
    });
  });

  test('page setup keeps the authored sectPr free of the inherited references', () => {
    const { hook, changed } = render();
    act(() => hook.result.current.handlePageSetupApply({ pageWidth: 11906, marginTop: 720 }));
    expect(changed).toHaveLength(1);
    const body = changed[0].package.document;
    expect(body.finalSectionProperties).toMatchObject({ pageWidth: 11906, marginTop: 720 });
    expect(body.finalSectionProperties?.headerReferences).toBeUndefined();
    expect(body.sections?.at(-1)?.properties).toMatchObject({
      pageWidth: 11906,
      marginTop: 720,
      headerReferences: inheritedHeader,
    });
    expect(body.sections?.[0]).toBe(probe.package.document.sections![0]);
  });
});
