import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Document, HeaderFooter, SectionProperties } from '@betteroffice/docx/types/document';
import { parseDocx } from '@betteroffice/docx/docx';
import type { PartEditTarget } from '../partEdit';
import { useHeaderFooterEditing } from './useHeaderFooterEditing';

// Two sections; the header and footer are authored on the first section's
// sectPr only, so the body sectPr inherits them.
const probe = await parseDocx(
  Uint8Array.from(readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx')))
    .buffer
);
type HeaderReference = NonNullable<SectionProperties['headerReferences']>[number];
const inheritedHeader: HeaderReference[] = [{ type: 'default', rId: 'rId3' }];

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

function render(document: Document, partEditTarget: PartEditTarget | null = null) {
  const pushed: Document[] = [];
  const targets: Array<PartEditTarget | null> = [];
  const hook = renderHook(() =>
    useHeaderFooterEditing({
      document,
      pushDocument: (doc) => {
        pushed.push(doc);
      },
      partEditTarget,
      setPartEditTarget: (next) => {
        targets.push(typeof next === 'function' ? next(partEditTarget) : next);
      },
    })
  );
  return { hook, pushed, targets };
}

function withSectionProperties(
  document: Document,
  patch: (properties: SectionProperties) => SectionProperties,
  headers = document.package.headers
): Document {
  const body = document.package.document;
  return {
    ...document,
    package: {
      ...document.package,
      headers,
      document: {
        ...body,
        sections: body.sections?.map((section) => ({
          ...section,
          properties: patch(section.properties),
        })),
      },
    },
  };
}

describe('useHeaderFooterEditing on a last section that inherits its header', () => {
  test('the probe authors the header on the first section only', () => {
    const body = probe.package.document;
    expect(body.sections).toHaveLength(2);
    expect(body.finalSectionProperties?.headerReferences).toBeUndefined();
    expect(body.sections?.at(-1)?.properties.headerReferences).toEqual(inheritedHeader);
  });

  test('double-clicking the header on the last page opens the inherited header', () => {
    const { hook, pushed, targets } = render(probe);
    expect(hook.result.current.headerContent).toBe(probe.package.headers?.get('rId3') ?? null);
    act(() => hook.result.current.handleHeaderFooterDoubleClick('header', 2));
    expect(pushed).toEqual([]);
    expect(targets).toEqual([{ kind: 'header', isFirstPage: false, pageIndex: 1 }]);
  });

  test('removing the header drops the inherited reference from the resolved last section', () => {
    const { hook, pushed } = render(probe, { kind: 'header', isFirstPage: false, pageIndex: 1 });
    act(() => hook.result.current.handleRemoveHeaderFooter());
    expect(pushed).toHaveLength(1);
    const next = pushed[0].package;
    expect(next.headers?.has('rId3')).toBe(false);
    expect(next.document.sections?.at(-1)?.properties.headerReferences).toEqual([]);
    expect(next.document.finalSectionProperties?.headerReferences).toBeUndefined();
  });

  test('an inherited titlePg opens the first-page variant on page 1', () => {
    const first: HeaderFooter = {
      type: 'header',
      hdrFtrType: 'first',
      content: [{ type: 'paragraph', content: [] }],
    };
    const document = withSectionProperties(
      probe,
      (properties) => ({
        ...properties,
        titlePg: true,
        headerReferences: [
          ...(properties.headerReferences ?? []),
          { type: 'first' as const, rId: 'rIdFirst' },
        ],
      }),
      new Map([...(probe.package.headers ?? []), ['rIdFirst', first]])
    );
    const { hook, pushed, targets } = render(document);
    expect(hook.result.current.firstPageHeaderContent).toBe(first);
    act(() => hook.result.current.handleHeaderFooterDoubleClick('header', 1));
    expect(pushed).toEqual([]);
    expect(targets).toEqual([{ kind: 'header', isFirstPage: true, pageIndex: 0 }]);
  });

  test('a header materialised where none exists is referenced on both views of the last section', () => {
    const document = withSectionProperties(probe, (properties) => ({
      ...properties,
      headerReferences: undefined,
    }));
    const { hook, pushed, targets } = render(document);
    expect(hook.result.current.headerContent).toBeNull();
    act(() => hook.result.current.handleHeaderFooterDoubleClick('header', 2));
    expect(pushed).toHaveLength(1);
    const next = pushed[0].package;
    const created: HeaderReference = { type: 'default', rId: 'rId_new_header_default' };
    expect(next.headers?.has(created.rId)).toBe(true);
    expect(next.relationships?.get(created.rId)?.target).toBe('header2.xml');
    expect(next.document.sections?.at(-1)?.properties.headerReferences).toEqual([created]);
    expect(next.document.finalSectionProperties?.headerReferences).toEqual([created]);
    expect(next.document.sections?.[0]?.properties.headerReferences).toBeUndefined();
    expect(targets).toEqual([{ kind: 'header', isFirstPage: false, pageIndex: 1 }]);
  });
});
