import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import {
  OPAQUE_SEED_BUDGET_BYTES,
  OpaqueSeedBudgetError,
  documentToYrs,
} from './documentToYrs';
import { createYrsSession } from './index';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

function opaqueDocument(xmlBytes: number): Document {
  const xml = `<w:object><o:OLEObject Type="Embed" ProgID="Eq"/>${'x'.repeat(xmlBytes)}</w:object>`;
  return {
    package: {
      document: {
        content: [
          {
            type: 'paragraph',
            paraId: '00000001',
            content: [{ type: 'run', content: [{ type: 'opaqueDrawing', kind: 'object', xml }] }],
          },
        ],
      },
    },
  };
}

function opaqueBytes(xmlBytes: number): Uint8Array<ArrayBuffer> {
  const filler = 'x'.repeat(xmlBytes);
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:r><w:object><o:OLEObject Type="Embed" ProgID="Eq"/><w:filler>${filler}</w:filler></w:object></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

it('projected seeder refuses an opaque payload over the aggregate budget', async () => {
  const session = await createYrsSession({ clientId: 74610 });
  try {
    expect(() => documentToYrs(session, opaqueDocument(OPAQUE_SEED_BUDGET_BYTES + 1))).toThrow(
      OpaqueSeedBudgetError
    );
    let error: unknown;
    try {
      documentToYrs(session, opaqueDocument(OPAQUE_SEED_BUDGET_BYTES + 1));
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(OpaqueSeedBudgetError);
    if (error instanceof OpaqueSeedBudgetError) {
      expect(error.budget).toBe(OPAQUE_SEED_BUDGET_BYTES);
      expect(error.total).toBeGreaterThan(OPAQUE_SEED_BUDGET_BYTES);
    }
    documentToYrs(session, opaqueDocument(16));
    expect(
      session.storySegments('body').filter((segment) => segment.kind === 'embed')
    ).toHaveLength(1);
  } finally {
    session.destroy();
  }
});

it('native seeder refuses an opaque payload over the aggregate budget', async () => {
  const session = await createYrsSession({ clientId: 74611 });
  try {
    expect(() => session.seedFromDocx(opaqueBytes(OPAQUE_SEED_BUDGET_BYTES + 1024))).toThrow(
      OpaqueSeedBudgetError
    );
  } finally {
    session.destroy();
  }
});

it('both seeders refuse many small opaque payloads over the aggregate budget', async () => {
  const each = 120 * 1024;
  const count = Math.floor(OPAQUE_SEED_BUDGET_BYTES / each) + 1;
  const paragraphs = Array.from({ length: count }, (_, index) => ({
    type: 'paragraph' as const,
    paraId: `p${index}`,
    content: [
      {
        type: 'run' as const,
        content: [{ type: 'opaqueDrawing' as const, kind: 'object', xml: `<w:object>${'y'.repeat(each)}</w:object>` }],
      },
    ],
  }));
  const document: Document = { package: { document: { content: paragraphs } } };
  const projected = await createYrsSession({ clientId: 74612 });
  try {
    expect(() => documentToYrs(projected, document)).toThrow(OpaqueSeedBudgetError);
  } finally {
    projected.destroy();
  }
  const filler = 'y'.repeat(each);
  const body = paragraphs
    .map(() => `<w:p><w:r><w:object>${filler}</w:object></w:r></w:p>`)
    .join('');
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`));
  parts.set('_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`));
  const native = await createYrsSession({ clientId: 74613 });
  try {
    expect(() => native.seedFromDocx(new Uint8Array(rezipPartsToArrayBuffer(parts)))).toThrow(
      OpaqueSeedBudgetError
    );
  } finally {
    native.destroy();
  }
});
